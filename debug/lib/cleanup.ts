import { Web } from '@pnp/sp/presets/all'
import { ISiteConnection } from './connect'
import { dateOfStamp, deleteWebTree, listScratchWebs, stampIn } from './scratch'

/**
 * Removing what the tests and the runner leave behind, as a plan first: the command prints it, and
 * only `--apply` carries it out.
 */

/**
 * Something to remove, and how.
 */
export interface ICleanupItem {
  kind:
    | 'web'
    | 'list'
    | 'content type'
    | 'site field'
    | 'custom action'
    | 'navigation node'
    | 'page'
    | 'site group'
  name: string
  remove: () => Promise<unknown>
}

/**
 * Parses an age such as `2h`, `30m` or `1d` into milliseconds.
 *
 * @param value - The age
 */
export function parseAge(value: string): number {
  const match = /^(\d+)\s*(m|h|d)$/i.exec(value.trim())
  if (!match)
    throw new Error(`Invalid age "${value}" (expected e.g. 30m, 2h, 1d)`)
  const unit = { m: 60_000, h: 3_600_000, d: 86_400_000 }[
    match[2].toLowerCase() as 'm' | 'h' | 'd'
  ]
  return Number(match[1]) * unit
}

/**
 * Leftover scratch webs below the reserved site, and the site groups that carry a stamp (a scratch
 * web's associated groups and fixture groups named after it), older than `olderThan` unless `all`.
 * The age threshold keeps a run from deleting the webs of another run that is still going.
 *
 * @param parent - The reserved site
 * @param options - Age threshold, or all
 */
export async function scratchLeftovers(
  parent: ISiteConnection,
  options: { olderThan: number; all?: boolean; now?: Date }
): Promise<ICleanupItem[]> {
  const now = (options.now || new Date()).getTime()
  const isOld = (stamp: string) =>
    options.all || now - dateOfStamp(stamp).getTime() > options.olderThan
  const webs = (await listScratchWebs(parent)).filter((web) => isOld(web.stamp))
  const groups = (
    await parent.web.siteGroups.select('Id', 'Title')<
      Array<{ Id: number; Title: string }>
    >()
  ).filter((group) => {
    const stamp = /^SPJS Itest /.test(group.Title) && stampIn(group.Title)
    return stamp && isOld(stamp)
  })
  const origin = new URL(parent.url).origin
  return [
    ...webs.map(
      (web): ICleanupItem => ({
        kind: 'web',
        name: web.serverRelativeUrl,
        remove: () =>
          deleteWebTree(Web([parent.web, `${origin}${web.serverRelativeUrl}`]))
      })
    ),
    ...groups.map(
      (group): ICleanupItem => ({
        kind: 'site group',
        name: group.Title,
        remove: () => parent.web.siteGroups.removeById(group.Id)
      })
    )
  ]
}

/**
 * The ids of the associated owner, member and visitor groups of a web and of its site
 * collection's root web (those they have). Fails rather than leave a group unprotected: only a
 * 404 counts as no group.
 *
 * @param conn - The web
 */
async function associatedGroupIds(conn: ISiteConnection): Promise<Set<number>> {
  const webs = [conn.web, conn.sp.site.rootWeb]
  const ids = await Promise.all(
    webs
      .map((web) => [
        web.associatedOwnerGroup,
        web.associatedMemberGroup,
        web.associatedVisitorGroup
      ])
      .reduce((all, groups) => [...all, ...groups], [])
      .map((group) =>
        group
          .select('Id')<{ Id?: number }>()
          .then((info) => info.Id)
          .catch((error) => {
            if (error && error.status === 404) return undefined
            throw error
          })
      )
  )
  return new Set(ids.filter((id): id is number => typeof id === 'number'))
}

/**
 * Everything on one web whose title (or internal name, or group) starts with a prefix, in the
 * order it can be removed: lists before the content types and site fields they use; site groups
 * last, as lists may grant them permissions.
 *
 * @param conn - The web
 * @param prefix - E.g. `SPJS Itest`; internal names are matched without spaces (`SPJSItest`)
 */
export async function prefixedArtefacts(
  conn: ISiteConnection,
  prefix: string
): Promise<ICleanupItem[]> {
  if (prefix.trim().length < 4)
    throw new Error(`Prefix "${prefix}" is too short to clean up by`)
  const internal = prefix.replace(/\s+/g, '')
  const starts = (value: string | undefined, start: string) =>
    !!value && value.toLowerCase().startsWith(start.toLowerCase())
  const { web } = conn
  const [lists, contentTypes, fields, actions, quickLaunch, topNav, groups] =
    await Promise.all([
      web.lists.select('Id', 'Title')<Array<{ Id: string; Title: string }>>(),
      web.contentTypes.select('Id', 'Name', 'Group')<
        Array<{ Id: { StringValue: string }; Name: string; Group: string }>
      >(),
      web.fields.select('Id', 'InternalName', 'Group')<
        Array<{ Id: string; InternalName: string; Group: string }>
      >(),
      web.userCustomActions.select('Id', 'Title')<
        Array<{ Id: string; Title: string }>
      >(),
      web.navigation.quicklaunch.select('Id', 'Title')<
        Array<{ Id: number; Title: string }>
      >(),
      web.navigation.topNavigationBar.select('Id', 'Title')<
        Array<{ Id: number; Title: string }>
      >(),
      web.siteGroups.select('Id', 'Title')<Array<{ Id: number; Title: string }>>()
    ])
  const pages = await web.lists
    .filter('BaseTemplate eq 119')
    .select('Id')<Array<{ Id: string }>>()
    .then((libraries) =>
      libraries.length
        ? web.lists
            .getById(libraries[0].Id)
            .items.select('Id', 'Title', 'FileLeafRef')<
              Array<{ Id: number; Title: string; FileLeafRef: string }>
            >()
            .then((items) => ({ libraryId: libraries[0].Id, items }))
        : { libraryId: '', items: [] }
    )
  const items: ICleanupItem[] = []
  for (const list of lists.filter((l) => starts(l.Title, prefix)))
    items.push({
      kind: 'list',
      name: list.Title,
      remove: () => web.lists.getById(list.Id).delete()
    })
  for (const page of pages.items.filter(
    (p) => starts(p.Title, prefix) || starts(p.FileLeafRef, internal)
  ))
    items.push({
      kind: 'page',
      name: page.FileLeafRef,
      remove: () =>
        web.lists.getById(pages.libraryId).items.getById(page.Id).delete()
    })
  for (const ct of contentTypes.filter(
    (c) => c.Group === prefix || starts(c.Name, prefix)
  ))
    items.push({
      kind: 'content type',
      name: ct.Name,
      remove: () => web.contentTypes.getById(ct.Id.StringValue).delete()
    })
  for (const field of fields.filter(
    (f) => f.Group === prefix || starts(f.InternalName, internal)
  ))
    items.push({
      kind: 'site field',
      name: field.InternalName,
      remove: () => web.fields.getById(field.Id).delete()
    })
  for (const action of actions.filter((a) => starts(a.Title, prefix)))
    items.push({
      kind: 'custom action',
      name: action.Title,
      remove: () => web.userCustomActions.getById(action.Id).delete()
    })
  for (const node of quickLaunch.filter((n) => starts(n.Title, prefix)))
    items.push({
      kind: 'navigation node',
      name: `QuickLaunch: ${node.Title}`,
      remove: () => web.navigation.quicklaunch.getById(node.Id).delete()
    })
  for (const node of topNav.filter((n) => starts(n.Title, prefix)))
    items.push({
      kind: 'navigation node',
      name: `TopNavigationBar: ${node.Title}`,
      remove: () => web.navigation.topNavigationBar.getById(node.Id).delete()
    })
  // Groups with a stamp belong to a scratch web (perhaps of a run still going): `--webs` removes
  // them by age. The web's own associated groups stay, whatever their title: a reserved site
  // titled `SPJS Itest` has `SPJS Itest Owners`.
  const associated = await associatedGroupIds(conn)
  for (const group of groups.filter(
    (g) =>
      starts(g.Title, prefix) && !stampIn(g.Title) && !associated.has(g.Id)
  ))
    items.push({
      kind: 'site group',
      name: group.Title,
      remove: () => web.siteGroups.removeById(group.Id)
    })
  return items
}
