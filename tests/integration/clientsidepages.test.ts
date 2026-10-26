import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { apply, IApplyResult } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import {
  IPageState,
  readClientSidePages,
  withoutVolatile
} from '../../debug/lib/readback/clientsidepages'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

/** The file name prefix the fixtures' pages share. */
const PREFIX = 'SPJS-Itest-'
/** The fixtures' page with `Overwrite: false`. */
const KEPT = 'SPJS-Itest-Kept.aspx'
/** The fixtures' page with `Overwrite: true`. */
const OVERWRITTEN = 'SPJS-Itest-Overwritten.aspx'

/** The control ids the handler turns into a text control. */
const TEXT_IDS = ['Text', 'PageText']

/**
 * The controls of a template page: its vertical section and every column of its sections.
 *
 * @param page - A template page
 */
function controlsOf(page: Record<string, any>): Array<Record<string, any>> {
  const columns = (page.Sections || []).reduce(
    (all: any[], section: any) => all.concat(section.Columns),
    []
  )
  return columns.reduce(
    (all: any[], column: any) => all.concat(column.Controls),
    [...(page.VerticalSection || [])]
  )
}

/**
 * The web part component ids a template uses (lower case, without the text controls).
 *
 * @param template - The template
 */
function webPartIdsOf(template: Record<string, any>): string[] {
  const ids = template.ClientSidePages.reduce(
    (all: string[], page: any) =>
      all.concat(
        controlsOf(page)
          .map((control) => control.Id)
          .filter((id: string) => !TEXT_IDS.includes(id))
          .map((id: string) => id.toLowerCase())
      ),
    []
  )
  return ids.filter((id: string, index: number) => ids.indexOf(id) === index)
}

/**
 * The web part definitions on the web, by component id (lower case, without braces): the
 * fixtures name OOTB web parts by literal component ids, which must exist before the handler can
 * add them (a missing one is only a logged warning).
 *
 * @param conn - The web
 * @param templates - The templates whose web parts must be there
 */
async function assertWebPartsAvailable(
  conn: ISiteConnection,
  templates: Array<Record<string, any>>
): Promise<Record<string, { title: string }>> {
  const parts: Array<{
    Id: string
    Manifest: string
  }> = await conn.web.getClientsideWebParts()
  const definitions: Record<string, { title: string }> = {}
  for (const part of parts) {
    const id = part.Id.replace(/^\{|\}$/g, '').toLowerCase()
    // Extensions and libraries are listed too, without preconfigured entries.
    const entries = JSON.parse(part.Manifest).preconfiguredEntries
    const entry = entries && entries[0]
    definitions[id] = { title: entry && entry.title && entry.title.default }
  }
  const wanted = templates.reduce(
    (all: string[], template) => all.concat(webPartIdsOf(template)),
    []
  )
  assert.deepEqual(
    wanted.filter((id) => !definitions[id]),
    [],
    'web part components the fixture uses are missing from getClientsideWebParts()'
  )
  return definitions
}

/**
 * A copy of a template value with `{sitename}` and `{site}` as the engine should resolve them on
 * the scratch web: the web's title, and the site collection's server-relative URL (the engine
 * reads `{site}` from `pageContext.site`).
 *
 * @param value - A template value
 * @param conn - The scratch web
 */
function resolved(value: any, conn: ISiteConnection): any {
  if (typeof value === 'string')
    return value
      .replace(/{sitename}/g, conn.webInfo.Title)
      .replace(/{site}/g, conn.siteInfo.ServerRelativeUrl)
  if (Array.isArray(value)) return value.map((item) => resolved(item, conn))
  if (value && typeof value === 'object') {
    const copy: Record<string, any> = {}
    for (const key of Object.keys(value)) copy[key] = resolved(value[key], conn)
    return copy
  }
  return value
}

/**
 * What a template control should read back as: a text control's HTML, or a web part's component
 * and definition id with the template's properties and the non-empty parts of its server processed
 * content.
 *
 * @param control - A template control
 * @param conn - The scratch web
 */
function expectedControl(
  control: Record<string, any>,
  conn: ISiteConnection
): Record<string, any> {
  if (TEXT_IDS.includes(control.Id)) return { type: 'text', text: control.Text }
  const expected: Record<string, any> = {
    type: 'webpart',
    webPartId: control.Id.toLowerCase(),
    definitionId: control.Id.toLowerCase(),
    properties: resolved(control.Properties, conn)
  }
  const content = control.ServerProcessedContent || {}
  const filled = Object.keys(content).filter(
    (key) => Object.keys(content[key]).length > 0
  )
  if (filled.length) {
    expected.serverProcessedContent = {}
    for (const key of filled)
      expected.serverProcessedContent[key] = resolved(content[key], conn)
  }
  return expected
}

/**
 * What a template page should read back as, as far as the template declares it.
 *
 * @param page - A template page
 * @param conn - The scratch web
 */
function expectedPage(
  page: Record<string, any>,
  conn: ISiteConnection
): Record<string, any> {
  return {
    title: page.Title,
    layout: page.PageLayoutType,
    commentsDisabled: !!page.CommentsDisabled,
    // The handler saves with publish: a page is published and not left checked out.
    published: true,
    checkedOut: false,
    sections: (page.Sections || []).map((section: any) => ({
      columns: section.Columns.map((column: any) => ({
        factor: column.Factor,
        controls: column.Controls.map((control: any) =>
          expectedControl(control, conn)
        )
      }))
    })),
    verticalSection: page.VerticalSection
      ? page.VerticalSection.map((control: any) =>
          expectedControl(control, conn)
        )
      : null
  }
}

/**
 * The read-back reduced to the members the expectation names: an object keeps only the
 * expectation's keys (SharePoint adds the web part's default properties), an array is reduced item
 * by item (an extra item is kept whole, so it shows in the diff), anything else is kept as is.
 *
 * @param actual - The read-back
 * @param expected - The expectation
 */
function projected(actual: any, expected: any): any {
  if (Array.isArray(actual) && Array.isArray(expected))
    return actual.map((item, index) =>
      index < expected.length ? projected(item, expected[index]) : item
    )
  if (
    actual &&
    expected &&
    typeof actual === 'object' &&
    typeof expected === 'object' &&
    !Array.isArray(expected)
  ) {
    const copy: Record<string, any> = {}
    for (const key of Object.keys(expected))
      if (key in actual) copy[key] = projected(actual[key], expected[key])
    return copy
  }
  return actual
}

/**
 * Asserts a page reads back as the template page declares it.
 *
 * @param actual - The read-back
 * @param page - The template page
 * @param conn - The scratch web
 */
function assertPage(
  actual: Partial<IPageState> | null,
  page: Record<string, any>,
  conn: ISiteConnection
): void {
  assert.ok(actual, `the page ${page.Name} was not created`)
  const expected = expectedPage(page, conn)
  assert.deepEqual(projected(actual, expected), expected, `page ${page.Name}`)
}

/**
 * The warnings the ClientSidePages handler logged (a web part it did not find is one).
 *
 * @param result - A run
 */
function handlerWarnings(result: IApplyResult): string[] {
  return result.warnings
    .filter((entry) => entry.source === 'ClientSidePages')
    .map((entry) => entry.message)
}

/**
 * The template page with the file name.
 *
 * @param template - The template
 * @param name - The page's `Name`
 */
function pageOf(
  template: Record<string, any>,
  name: string
): Record<string, any> {
  return template.ClientSidePages.find((page: any) => page.Name === name)
}

/**
 * The ClientSidePages handler on a scratch web: modern Article pages with sections, columns, a
 * vertical section, text controls and OOTB web parts (by literal component id), web part
 * properties with `{sitename}` and `{site}`, server processed content with `{sitename}`, comments
 * on and off, and `Overwrite` both ways.
 */
describe('ClientSidePages', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('clientsidepages')

  it('creates the pages with their sections, columns and controls, skips the Overwrite false page and recreates the Overwrite true page on a second apply', async () => {
    const { conn } = scratch()
    const template = loadFixture('clientsidepages')
    const definitions = await assertWebPartsAvailable(conn, [template])
    const names: string[] = template.ClientSidePages.map(
      (page: any) => page.Name
    )
    const recreated = names.filter((name) => pageOf(template, name).Overwrite)
    assert.deepEqual(recreated, [OVERWRITTEN])

    // The file identity of each page after each run, beside the read-back: a recreated page has a
    // new UniqueId, so applyTwice compares it without that.
    const identities: Array<Record<string, string | undefined>> = []
    const { first, second, afterFirst } = await applyTwice(
      conn,
      () => loadFixture('clientsidepages'),
      async (c) => {
        const state = await readClientSidePages(c, names, PREFIX)
        const ids: Record<string, string | undefined> = {}
        const pages: Record<string, Partial<IPageState> | null> = {}
        for (const name of names) {
          const page = state.pages[name]
          ids[name] = page ? page.uniqueId : undefined
          pages[name] = recreated.includes(name) ? withoutVolatile(page) : page
        }
        identities.push(ids)
        return { pages, files: state.files }
      }
    )

    assert.ok(first.order.includes('ClientSidePages'), 'the handler ran')
    for (const page of template.ClientSidePages)
      assertPage(afterFirst.pages[page.Name], page, conn)

    // A web part's title comes from its definition's preconfigured entry.
    for (const name of names) {
      const page = afterFirst.pages[name]
      const parts = page.sections
        .reduce((all, section) => all.concat(section.columns), [])
        .reduce((all, column) => all.concat(column.controls), [])
        .filter((control) => control.type === 'webpart')
      for (const part of parts)
        assert.equal(
          part.title,
          definitions[part.webPartId].title,
          `the title of ${part.webPartId} on ${name}`
        )
    }

    // One file per page: neither run created a second one.
    assert.deepEqual(afterFirst.files, [...names].sort())
    // Overwrite false: the second run skipped the page (same file; applyTwice also compared its
    // TimeLastModified). Overwrite true: the second run deleted it and created it again.
    assert.equal(identities[1][KEPT], identities[0][KEPT])
    assert.notEqual(identities[1][OVERWRITTEN], identities[0][OVERWRITTEN])
    // Every web part was found (a missing one is a warning, and the page is saved without it).
    assert.deepEqual(handlerWarnings(first), [], 'the first run warned')
    assert.deepEqual(handlerWarnings(second), [], 'the second run warned')
  })

  it('leaves the Overwrite false page as it is and replaces the Overwrite true page when the template changes', async () => {
    const { conn } = scratch()
    const b = loadFixture('clientsidepages.b')
    await assertWebPartsAvailable(conn, [b])
    const names = [KEPT, OVERWRITTEN]

    assertClean(await apply(conn, loadFixture('clientsidepages')), 'the A run')
    const before = await readClientSidePages(conn, names, PREFIX)
    const changed = await apply(conn, b)
    assertClean(changed, 'the B run')
    assert.deepEqual(handlerWarnings(changed), [], 'the B run warned')
    const after = await readClientSidePages(conn, names, PREFIX)

    // B changes the kept page's title, comments and content: none of it is applied.
    assert.ok(before.pages[KEPT], `the page ${KEPT} was not created`)
    assert.deepEqual(after.pages[KEPT], before.pages[KEPT])
    // The overwritten page is B's, in a new file.
    assertPage(after.pages[OVERWRITTEN], pageOf(b, OVERWRITTEN), conn)
    assert.notEqual(
      after.pages[OVERWRITTEN].uniqueId,
      before.pages[OVERWRITTEN].uniqueId
    )
    assert.deepEqual(after.files, [...names].sort())
  })
})
