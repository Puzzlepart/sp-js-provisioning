import { ISiteGroup } from '@pnp/sp/presets/all'
import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { apply } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import {
  IListState,
  readList,
  readLists,
  roleAssignmentsOf
} from '../../debug/lib/readback/lists'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

const TASKS = 'SPJS Itest Oppgaver'
const DOCS = 'SPJS Itest Dokumenter'

/** The attributes of a field's schema XML (`<Field A="..." ...>`). */
function attributesOf(schemaXml: string): Record<string, string> {
  const attributes: Record<string, string> = {}
  const head = /^<Field\b([^>]*)>/.exec(schemaXml)
  const pattern = /(\w+)="([^"]*)"/g
  let match: RegExpExecArray | null
  while (head && (match = pattern.exec(head[1])))
    attributes[match[1]] = match[2]
  return attributes
}

/** A GUID without braces, lower case. */
const bare = (guid: string) => guid.replace(/[{}]/g, '').toLowerCase()

/** Folder paths of a `Folders` tree, relative to the list's root folder. */
function folderPaths(
  folders: Array<{ Name: string; Folders?: any[] }> = [],
  parent = ''
): string[] {
  return folders.reduce<string[]>((paths, folder) => {
    const path = parent ? `${parent}/${folder.Name}` : folder.Name
    return [...paths, path, ...folderPaths(folder.Folders, path)]
  }, [])
}

/** The id of the site content type a list content type was made from. */
const parentOf = (listContentTypeId: string) =>
  listContentTypeId.slice(0, -34).toUpperCase()

/** A site group's title. */
const titleOf = async (group: ISiteGroup) =>
  (await group.select('Title')<{ Title: string }>()).Title

/**
 * A connection whose web records every request the engine sends through it (method and path
 * below the web), so a test can see what a run asked for, not only what it left behind.
 *
 * @param conn - The scratch web
 * @param requests - Where the requests go
 */
function recording(conn: ISiteConnection, requests: string[]): ISiteConnection {
  return Object.create(conn, {
    web: {
      get: () => {
        const web = conn.sp.web
        web.on.pre(function (url: string, init: any, result: any) {
          requests.push(
            `${(init && init.method) || 'GET'} ${decodeURIComponent(
              url.slice(conn.url.length)
            )}`
          )
          return Promise.resolve<[string, any, any]>([url, init, result])
        })
        return web
      }
    }
  })
}

/**
 * The Lists handler (structure) on a scratch web: a generic list and a library with list
 * properties, content type bindings, list fields (a LookupMulti to the other list through
 * `{listid:}` and `{siteid}`), field refs to site columns from the same template, views, a folder
 * tree and list permissions granted to the scratch web's associated groups. DataRows have their
 * own suite. Known defects (field recreation, unconditional `breakRoleInheritance`,
 * `RemoveExistingViews` on every run) are pinned as `todo` (docs/specs/integration-testing.md, D4).
 *
 * The tests share one scratch web and run in order: the ones after the first build on the lists it
 * provisioned, so run the file whole (a name pattern that skips the first one fails them).
 */
describe('Lists', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('lists')
  let roles: {
    fullControl: string
    edit: string
    contribute: string
    read: string
  }
  let groups: { owners: string; members: string; visitors: string }
  let userTitle: string | undefined

  before(async () => {
    const { conn } = scratch()
    const roleName = async (kind: 2 | 3 | 5 | 6) =>
      (
        await conn.web.roleDefinitions.getByType(kind).select('Name')<{
          Name: string
        }>()
      ).Name
    roles = {
      fullControl: await roleName(5),
      edit: await roleName(6),
      contribute: await roleName(3),
      read: await roleName(2)
    }
    groups = {
      owners: await titleOf(conn.web.associatedOwnerGroup),
      members: await titleOf(conn.web.associatedMemberGroup),
      visitors: await titleOf(conn.web.associatedVisitorGroup)
    }
    userTitle =
      conn.mode === 'cookies'
        ? (await conn.web.currentUser.select('Title')<{ Title: string }>())
            .Title
        : undefined
  })

  it('creates a list and a library with fields, field refs, views, folders and permissions, and changes nothing on a second apply', async () => {
    const { conn } = scratch()
    const webAssignments = roleAssignmentsOf(
      await conn.web.roleAssignments.expand(
        'Member',
        'RoleDefinitionBindings'
      )()
    )
    const { first, second, afterFirst } = await applyTwice(
      conn,
      () => loadFixture('lists'),
      (c) => readLists(c, [TASKS, DOCS])
    )
    assert.deepEqual(
      [...first.warnings, ...second.warnings].map((entry) => entry.message),
      [],
      'the runs logged warnings'
    )
    const tasksId = (
      await conn.web.lists.getByTitle(TASKS).select('Id')<{ Id: string }>()
    ).Id

    for (const lc of loadFixture('lists').Lists) {
      const list: IListState = afterFirst[lc.Title]
      const label = (what: string) => `${lc.Title}: ${what}`

      // List properties, AdditionalSettings included (each one is read back).
      assert.equal(list.Title, lc.Title)
      assert.equal(list.Description, lc.Description)
      assert.equal(list.BaseTemplate, lc.Template)
      assert.equal(list.ContentTypesEnabled, lc.ContentTypesEnabled)
      for (const [key, value] of Object.entries(lc.AdditionalSettings))
        assert.equal((list as any)[key], value, label(key))

      // Bindings to OOTB content types, each bound once; nothing removed (SharePoint's Folder stays).
      const parents = list.contentTypes.map((ct) => parentOf(ct.Id))
      assert.deepEqual(
        [...parents].sort(),
        [
          ...lc.ContentTypeBindings.map((b: any) =>
            b.ContentTypeID.toUpperCase()
          ),
          '0X0120'
        ].sort(),
        label('content types')
      )

      // List fields from XML: internal name from Name, title from DisplayName.
      for (const xml of lc.Fields) {
        const attributes = attributesOf(xml)
        const field = list.fields[attributes.Name]
        assert.ok(field, label(`field ${attributes.Name}`))
        assert.equal(field.Id, bare(attributes.ID))
        assert.equal(field.Title, attributes.DisplayName)
        assert.equal(field.TypeAsString, attributes.Type)
        if (attributes.Type === 'LookupMulti') {
          // `{listid:...}` and `{siteid}` resolved to the other list and the scratch web.
          assert.equal(field.LookupList, bare(tasksId), label('LookupList'))
          assert.equal(
            field.LookupWebId,
            bare(conn.webInfo.Id),
            label('LookupWebId')
          )
          assert.equal(field.LookupField, attributes.ShowField)
          assert.equal(field.AllowMultipleValues, true)
        }
      }

      // Field refs to the template's site columns, with the list's own title, Required and Hidden.
      for (const reference of lc.FieldRefs) {
        const field = list.fields[reference.Name]
        assert.ok(field, label(`field reference ${reference.Name}`))
        assert.equal(field.Id, reference.ID)
        assert.equal(
          field.Title,
          reference.DisplayName,
          label(`${reference.Name} title`)
        )
        assert.equal(
          field.Required,
          !!reference.Required,
          label(`${reference.Name} Required`)
        )
        assert.equal(
          field.Hidden,
          !!reference.Hidden,
          label(`${reference.Name} Hidden`)
        )
      }

      // Views: one per title, with their fields and settings; the template's default view wins.
      for (const lvc of lc.Views) {
        const views = list.views.filter((view) => view.Title === lvc.Title)
        assert.equal(views.length, 1, label(`views titled ${lvc.Title}`))
        const [view] = views
        const settings = lvc.AdditionalSettings
        assert.deepEqual(view.fields, lvc.ViewFields, label('view fields'))
        assert.equal(view.ViewQuery, settings.ViewQuery)
        assert.equal(view.RowLimit, settings.RowLimit)
        assert.equal(view.Paged, settings.Paged)
        assert.equal(view.DefaultView, settings.DefaultView)
        assert.equal(view.Scope, settings.Scope ?? 0, label('Scope'))
        assert.equal(view.PersonalView, false)
      }
      const defaults = list.views.filter((view) => view.DefaultView)
      assert.equal(defaults.length, 1, label('default views'))
      const templateDefault = lc.Views.find(
        (view: any) => view.AdditionalSettings.DefaultView
      )
      if (templateDefault)
        assert.equal(defaults[0].Title, templateDefault.Title)

      // The folder tree, once.
      assert.deepEqual(
        list.folders,
        folderPaths(lc.Folders).sort(),
        label('folders')
      )
    }

    // Security on the list: inheritance broken, the web's assignments copied, the visitors also
    // contribute.
    const tasks = afterFirst[TASKS]
    assert.equal(tasks.HasUniqueRoleAssignments, true)
    assert.deepEqual(tasks.roleAssignments, {
      ...webAssignments,
      [groups.visitors]: [
        ...(webAssignments[groups.visitors] || []),
        roles.contribute
      ].sort()
    })
    // On the library: not copied, so only the template's groups (and the user who broke the
    // inheritance, whom SharePoint adds with full control); the visitors have no access.
    const library = afterFirst[DOCS]
    assert.equal(library.HasUniqueRoleAssignments, true)
    assert.deepEqual(library.roleAssignments[groups.owners], [
      roles.fullControl
    ])
    assert.deepEqual(library.roleAssignments[groups.members], [roles.edit])
    assert.equal(library.roleAssignments[groups.visitors], undefined)
    if (userTitle)
      assert.deepEqual(
        Object.keys(library.roleAssignments).sort(),
        [groups.owners, groups.members, userTitle].sort()
      )
  })

  it(
    'keeps the values in list Fields columns on a second apply',
    {
      todo:
        'D4, pinned: processField deletes and recreates every list field on every run, so ' +
        'the column values of existing items are lost (Text emptied, LookupMulti cleared)'
    },
    async () => {
      const { conn } = scratch()
      const item = await conn.web.lists.getByTitle(TASKS).items.add({
        Title: 'SPJS Itest oppgave',
        SPJSItestCode: 'K-1',
        SPJSItestCategory: 'Plan'
      })
      const root: {
        ServerRelativeUrl: string
      } = await conn.web.lists
        .getByTitle(DOCS)
        .rootFolder.select('ServerRelativeUrl')()
      const added = await conn.web
        .getFolderByServerRelativePath(root.ServerRelativeUrl)
        .files.addUsingPath('spjs-itest.txt', 'SPJS Itest', {
          Overwrite: true
        })
      const file = () =>
        conn.web.getFileByServerRelativePath(added.ServerRelativeUrl)
      await (await file().getItem()).update({
        SPJSItestRelatedId: [item.Id],
        SPJSItestCategory: 'Plan'
      })
      const values = async () => ({
        item: await conn.web.lists
          .getByTitle(TASKS)
          .items.getById(item.Id)
          .select('SPJSItestCode', 'SPJSItestCategory')<any>(),
        file: await file().listItemAllFields.select(
          'SPJSItestRelatedId',
          'SPJSItestCategory'
        )<any>()
      })
      const stored = await values()
      assert.equal(stored.item.SPJSItestCode, 'K-1')
      assert.deepEqual(stored.file.SPJSItestRelatedId, [item.Id])

      assertClean(await apply(conn, loadFixture('lists')), 'the run')

      const kept = await values()
      // Field refs are updated in place: their values survive.
      assert.equal(kept.item.SPJSItestCategory, 'Plan')
      assert.equal(kept.file.SPJSItestCategory, 'Plan')
      // List Fields are recreated: these do not (today null and []).
      assert.deepEqual(
        {
          text: kept.item.SPJSItestCode,
          lookupMulti: kept.file.SPJSItestRelatedId
        },
        { text: 'K-1', lookupMulti: [item.Id] },
        'the values in the list Fields columns'
      )
    }
  )

  it('keeps role assignments made on a list after the first apply', async () => {
    // SharePoint ignores breakRoleInheritance on a list that already has unique permissions (even
    // with clearSubscopes), so the unconditional call (pinned below) loses nothing today.
    const { conn } = scratch()
    const visitors = await conn.web.associatedVisitorGroup.select('Id')<{
      Id: number
    }>()
    const read = await conn.web.roleDefinitions.getByType(2).select('Id')<{
      Id: number
    }>()
    await conn.web.lists
      .getByTitle(DOCS)
      .roleAssignments.add(visitors.Id, read.Id)
    const granted = await readList(conn, DOCS)
    assert.deepEqual(granted.roleAssignments[groups.visitors], [roles.read])

    assertClean(await apply(conn, loadFixture('lists')), 'the run')

    assert.deepEqual(
      (await readList(conn, DOCS)).roleAssignments,
      granted.roleAssignments
    )
  })

  it(
    'does not break role inheritance again on a list that already has unique permissions',
    {
      todo:
        'D4, pinned: processListSecurity calls breakRoleInheritance on every run, whether ' +
        'or not the list already has unique permissions'
    },
    async () => {
      const { conn } = scratch()
      for (const title of [TASKS, DOCS])
        assert.equal(
          (await readList(conn, title)).HasUniqueRoleAssignments,
          true
        )
      const requests: string[] = []
      assertClean(
        await apply(recording(conn, requests), loadFixture('lists')),
        'the run'
      )
      assert.ok(
        requests.some((request) => /addroleassignment/i.test(request)),
        'the requests were recorded'
      )
      assert.deepEqual(
        requests.filter((request) => /breakroleinheritance/i.test(request)),
        []
      )
    }
  )

  it(
    'replaces the existing views with the template views once, and keeps them on a second apply',
    {
      todo:
        "D4, pinned: RemoveExistingViews deletes every view on every run, the template's own " +
        'included, and adds them again: each re-apply gives the views new ids (links to them ' +
        'break) and removes views users made since'
    },
    async () => {
      const { conn } = scratch()
      const title = 'SPJS Itest Visninger'
      const template = () => ({
        Lists: [
          {
            Title: title,
            Description: 'RemoveExistingViews',
            Template: 100,
            ContentTypesEnabled: false,
            RemoveExistingViews: true,
            Views: [
              {
                Title: 'SPJS Itest Alle',
                ViewFields: ['LinkTitle'],
                AdditionalSettings: { DefaultView: true, RowLimit: 30 }
              }
            ]
          }
        ]
      })
      const views = async () =>
        (
          await conn.web.lists
            .getByTitle(title)
            .views.select('Id', 'Title', 'DefaultView')<
            Array<{ Id: string; Title: string; DefaultView: boolean }>
          >()
        ).map(({ Id, Title, DefaultView }) => ({ Id, Title, DefaultView }))

      assertClean(await apply(conn, template()), 'the first run')
      const first = await views()
      assert.deepEqual(
        first.map((view) => [view.Title, view.DefaultView]),
        [['SPJS Itest Alle', true]]
      )
      assertClean(await apply(conn, template()), 'the second run')
      // The template's own view is kept (same id), not deleted and recreated.
      assert.deepEqual(await views(), first)
    }
  )
})
