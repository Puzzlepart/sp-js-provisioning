import assert from 'node:assert/strict'
import { before, describe, it } from 'node:test'
import { apply, IApplyResult } from '../../debug/lib/apply'
import { ISiteConnection } from '../../debug/lib/connect'
import { env } from '../../debug/lib/env'
import {
  IDataRowItem,
  IDataRowsList,
  readDataRows
} from '../../debug/lib/readback/lists-datarows'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, assertClean, skipReason, withScratchWeb } from './helpers'

const SOURCE = 'SPJS Itest DataRows Source'
const OVERWRITE = 'SPJS Itest DataRows Overwrite'
const SKIP = 'SPJS Itest DataRows Skip'

/** Values for the fixture's `{parameter:...}` tokens. */
const parameters = {
  SPJSItestText: 'from a parameter',
  SPJSItestSlug: 'from-parameter'
}

/** The columns the fixture's rows write, by list. */
const COLUMNS: Record<string, string[]> = {
  [SOURCE]: ['Title'],
  [OVERWRITE]: [
    'Title',
    'SPJSItestDrText',
    'SPJSItestDrNumber',
    'SPJSItestDrFlag',
    'SPJSItestDrDate',
    'SPJSItestDrChoice',
    'SPJSItestDrChoices',
    'SPJSItestDrLink',
    'SPJSItestDrPerson',
    'SPJSItestDrLookup',
    'SPJSItestDrLookups'
  ],
  [SKIP]: ['Title', 'SPJSItestDrNumber', 'SPJSItestDrText']
}

const read = (conn: ISiteConnection) => readDataRows(conn, COLUMNS)

/**
 * An item's raw values, with a DateTime as the web's local wall time (what a user picked).
 *
 * @param item - The item
 */
function valuesOf(item: IDataRowItem): Record<string, any> {
  const values: Record<string, any> = {}
  for (const name of Object.keys(item.values)) {
    const value = item.values[name]
    values[name] = value && value.local !== undefined ? value.local : value
  }
  return values
}

/**
 * The item with the given title.
 *
 * @param list - The list
 * @param title - The title
 */
function byTitle(list: IDataRowsList, title: string): IDataRowItem {
  const item = list.items.find((candidate) => candidate.values.Title === title)
  assert.ok(item, `no item ${title}`)
  return item
}

/**
 * The Lists handler's DataRows on a scratch web: every value shape the handler writes (Text,
 * Number, Boolean, a date-only DateTime, Choice, MultiChoice, URL, User, Lookup and LookupMulti
 * by value and by id), tokens in text and URL values, and the upsert by a Title and by a Number
 * `KeyColumn` with `Overwrite` and `Skip`. The columns the rows must keep across runs are site
 * columns added through `FieldRefs`: a column from the list's `Fields` is deleted and recreated on
 * every run (pinned below), so only the `Overwrite` list has its lookups there.
 */
describe('Lists (DataRows)', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('lists-datarows')
  let user: string | undefined
  let userId: number
  let userTitle: string
  const template = () => loadFixture('lists-datarows', { user })

  before(async () => {
    const { conn } = scratch()
    user =
      conn.mode === 'cookies'
        ? (
            await conn.web.currentUser.select('LoginName')<{
              LoginName: string
            }>()
          ).LoginName
        : env.user
    if (user) {
      // The engine ensures a User value the same way.
      const ensured = await conn.web.ensureUser(user)
      userId = ensured.Id
      userTitle = ensured.Title
    }
  })

  it('writes every value shape, upserts by Title and by a Number key, and changes nothing on a second apply', async (t) => {
    if (!user) {
      t.skip('app-only without INTEGRATION_USER: no user for the User column')
      return
    }
    const { conn } = scratch()
    const site = conn.webInfo.Title
    const { afterFirst } = await applyTwice(conn, template, read, {
      parameters
    })

    const source = afterFirst[SOURCE]
    assert.deepEqual(
      source.items.map((item) => item.values.Title),
      ['SPJS Itest Source A', 'SPJS Itest Source B', 'SPJS Itest Source C']
    )
    const sourceId = (title: string) => byTitle(source, title).Id
    const [a, b, c] = ['A', 'B', 'C'].map((letter) =>
      sourceId(`SPJS Itest Source ${letter}`)
    )

    const overwrite = afterFirst[OVERWRITE]
    assert.deepEqual(overwrite.types, {
      Title: 'Text',
      SPJSItestDrText: 'Text',
      SPJSItestDrNumber: 'Number',
      SPJSItestDrFlag: 'Boolean',
      SPJSItestDrDate: 'DateTime',
      SPJSItestDrChoice: 'Choice',
      SPJSItestDrChoices: 'MultiChoice',
      SPJSItestDrLink: 'URL',
      SPJSItestDrPerson: 'User',
      SPJSItestDrLookup: 'Lookup',
      SPJSItestDrLookups: 'LookupMulti'
    })
    assert.equal(overwrite.items.length, 3)
    // Tokens resolve in text and URL values; braces that are no engine token stay as written.
    assert.deepEqual(valuesOf(byTitle(overwrite, 'SPJS Itest Row A')), {
      Title: 'SPJS Itest Row A',
      SPJSItestDrText: `${site} / ${parameters.SPJSItestText}`,
      SPJSItestDrNumber: 12.5,
      SPJSItestDrFlag: true,
      SPJSItestDrDate: '2026-05-17T00:00:00',
      SPJSItestDrChoice: 'Beta',
      SPJSItestDrChoices: ['Alfa', 'Gamma'],
      SPJSItestDrLink: {
        Url: `https://contoso.com/spjs-itest/${parameters.SPJSItestSlug}`,
        Description: 'SPJS Itest link'
      },
      SPJSItestDrPerson: userId,
      // `{ lookupValue }` by the show field; a LookupMulti mixes it with a bare id.
      SPJSItestDrLookup: a,
      SPJSItestDrLookups: [b, 3]
    })
    assert.deepEqual(valuesOf(byTitle(overwrite, 'SPJS Itest Row B')), {
      Title: 'SPJS Itest Row B',
      SPJSItestDrText: 'Kept {Title} and {unknown:value}',
      SPJSItestDrNumber: 7,
      SPJSItestDrFlag: false,
      SPJSItestDrDate: '2026-12-31T00:00:00',
      SPJSItestDrChoice: 'Gamma',
      // A single value for a MultiChoice is written as a one-item array.
      SPJSItestDrChoices: ['Beta'],
      // A string URL: SharePoint fills the description with the URL.
      SPJSItestDrLink: {
        Url: 'https://fabrikam.com/spjs-itest',
        Description: 'https://fabrikam.com/spjs-itest'
      },
      // `{ login }`.
      SPJSItestDrPerson: userId,
      // A bare id and `{ lookupId }`.
      SPJSItestDrLookup: 2,
      SPJSItestDrLookups: [1]
    })
    assert.deepEqual(valuesOf(byTitle(overwrite, 'SPJS Itest Row C')), {
      Title: 'SPJS Itest Row C',
      SPJSItestDrText: null,
      SPJSItestDrNumber: null,
      // `1` is a Boolean true.
      SPJSItestDrFlag: true,
      SPJSItestDrDate: null,
      SPJSItestDrChoice: null,
      SPJSItestDrChoices: null,
      SPJSItestDrLink: null,
      SPJSItestDrPerson: null,
      // A string is a lookup value, like `{ lookupValue }`.
      SPJSItestDrLookup: c,
      SPJSItestDrLookups: []
    })
    // The lookups and the user show what they point at.
    const rowA = byTitle(overwrite, 'SPJS Itest Row A')
    assert.deepEqual(
      [
        rowA.text.SPJSItestDrLookup,
        rowA.text.SPJSItestDrLookups,
        rowA.text.SPJSItestDrPerson
      ],
      [
        'SPJS Itest Source A',
        'SPJS Itest Source B; SPJS Itest Source C',
        userTitle
      ]
    )

    const skip = afterFirst[SKIP]
    assert.deepEqual(skip.types, {
      Title: 'Text',
      SPJSItestDrNumber: 'Number',
      SPJSItestDrText: 'Text'
    })
    assert.deepEqual(skip.items.map(valuesOf), [
      {
        Title: 'SPJS Itest Step 10',
        SPJSItestDrNumber: 10,
        SPJSItestDrText: 'First step'
      },
      {
        Title: 'SPJS Itest Step 20',
        SPJSItestDrNumber: 20,
        SPJSItestDrText: 'Second step'
      },
      {
        Title: 'SPJS Itest Step 30',
        SPJSItestDrNumber: 30,
        SPJSItestDrText: 'Third step'
      }
    ])
  })

  it('restores an edited row with Overwrite, keeps it with Skip, and adds a removed row again', async (t) => {
    if (!user) {
      t.skip('app-only without INTEGRATION_USER: no user for the User column')
      return
    }
    const { conn } = scratch()
    const prior = await read(conn)
    const rowA = byTitle(prior[OVERWRITE], 'SPJS Itest Row A')
    const step20 = byTitle(prior[SKIP], 'SPJS Itest Step 20')
    const step30 = byTitle(prior[SKIP], 'SPJS Itest Step 30')
    await conn.web.lists
      .getByTitle(OVERWRITE)
      .items.getById(rowA.Id)
      .update({
        SPJSItestDrText: 'Edited by hand',
        SPJSItestDrChoice: 'Alfa',
        SPJSItestDrLookupId: byTitle(prior[SOURCE], 'SPJS Itest Source C').Id
      })
    await conn.web.lists
      .getByTitle(SKIP)
      .items.getById(step20.Id)
      .update({ SPJSItestDrText: 'Edited by hand' })
    await conn.web.lists.getByTitle(SKIP).items.getById(step30.Id).recycle()

    assertClean(await apply(conn, template(), { parameters }), 'the third run')
    const after = await read(conn)

    // No list gained or lost a row; the edit is undone by Overwrite, kept by Skip.
    assert.deepEqual(after[SOURCE], prior[SOURCE])
    assert.deepEqual(after[OVERWRITE], prior[OVERWRITE])
    assert.deepEqual(
      after[SKIP].items.map((item) => [
        item.values.Title,
        item.values.SPJSItestDrText
      ]),
      [
        ['SPJS Itest Step 10', 'First step'],
        ['SPJS Itest Step 20', 'Edited by hand'],
        ['SPJS Itest Step 30', 'Third step']
      ]
    )
    // The removed row is a new item; the others kept their ids.
    assert.equal(byTitle(after[SKIP], 'SPJS Itest Step 20').Id, step20.Id)
    assert.notEqual(byTitle(after[SKIP], 'SPJS Itest Step 30').Id, step30.Id)
  })

  describe('defects (lists-datarows.b.json)', () => {
    const RECREATED = 'SPJS Itest DataRows Recreated'
    const TOKENISED = 'SPJS Itest DataRows Tokenised'
    const LINK = 'SPJS Itest DataRows Link'
    const UNRESOLVED = 'SPJS Itest DataRows Unresolved'
    const readDefects = (conn: ISiteConnection) =>
      readDataRows(conn, {
        [RECREATED]: ['SPJSItestDrNote'],
        [TOKENISED]: [],
        [LINK]: ['SPJSItestDrLinkOnly'],
        [UNRESOLVED]: ['SPJSItestDrMissing', 'SPJSItestDrLookups']
      })
    let runs:
      | Promise<{
          first: IApplyResult
          second: IApplyResult
          afterFirst: Record<string, IDataRowsList>
          afterSecond: Record<string, IDataRowsList>
        }>
      | undefined
    /** Applies the defect fixture twice, once for all the pins below. */
    const twice = () =>
      (runs =
        runs ||
        (async () => {
          const { conn } = scratch()
          const first = await apply(conn, loadFixture('lists-datarows.b'), {
            parameters
          })
          const afterFirst = await readDefects(conn)
          const second = await apply(conn, loadFixture('lists-datarows.b'), {
            parameters
          })
          const afterSecond = await readDefects(conn)
          return { first, second, afterFirst, afterSecond }
        })())

    it(
      'keeps a Skip row value in a column from the list Fields on a second apply',
      {
        todo:
          'processField deletes and recreates every list Fields column on each run, so a Skip ' +
          'row loses the value (D4 Lists spec)'
      },
      async () => {
        const { afterFirst, afterSecond } = await twice()
        assert.deepEqual(
          afterFirst[RECREATED].items.map(valuesOf),
          [{ Title: 'SPJS Itest Kept', SPJSItestDrNote: 'Written once' }],
          'the first run wrote the row'
        )
        assert.deepEqual(
          afterSecond[RECREATED].items.map(valuesOf),
          afterFirst[RECREATED].items.map(valuesOf)
        )
      }
    )

    it(
      'matches a row whose KeyColumn value holds a token',
      {
        todo:
          'the KeyColumn filter uses the raw value ({sitename} row) while the item holds the ' +
          'resolved one, so every run adds the row again (D4 Lists spec)'
      },
      async () => {
        const { afterSecond } = await twice()
        const { conn } = scratch()
        assert.deepEqual(
          afterSecond[TOKENISED].items.map((item) => item.values.Title),
          [`${conn.webInfo.Title} row`]
        )
      }
    )

    it(
      'describes a URL without a Description by its resolved Url',
      {
        todo:
          'a { Url } value without Description gets Description = the raw Url, tokens unresolved'
      },
      async () => {
        const { afterFirst } = await twice()
        const url = `https://contoso.com/spjs-itest/${parameters.SPJSItestSlug}`
        assert.deepEqual(
          valuesOf(afterFirst[LINK].items[0]).SPJSItestDrLinkOnly,
          { Url: url, Description: url }
        )
      }
    )

    /**
     * The messages of the errors and warnings a run logged.
     *
     * @param run - The run
     */
    const logged = (run: IApplyResult) =>
      [...run.errors, ...run.warnings].map((entry) => entry.message)

    it(
      'logs a Lookup value that matches no item',
      {
        todo:
          'an unresolved Lookup value is dropped from the payload without a log line: the ' +
          'column stays empty (D4 Lists spec)'
      },
      async () => {
        const { first, afterFirst } = await twice()
        assert.deepEqual(
          valuesOf(
            byTitle(afterFirst[UNRESOLVED], 'SPJS Itest Unresolved single')
          ).SPJSItestDrMissing,
          null,
          'the row is written without the value'
        )
        assert.ok(
          logged(first).some((message) =>
            message.includes('SPJS Itest missing single')
          ),
          `nothing logged names the unresolved value: ${JSON.stringify(
            logged(first)
          )}`
        )
      }
    )

    it(
      'writes the resolved values of a LookupMulti and logs the one that matches no item',
      {
        todo:
          'an unresolved LookupMulti value is sent as null: SharePoint refuses the whole row ' +
          '(400 "A null value was found with the expected type \'Edm.Int32[Nullable=False]\'"), ' +
          'and the logged error does not name the value (D4 Lists spec)'
      },
      async () => {
        const { first, afterFirst } = await twice()
        const sourceA = (
          await scratch()
            .conn.web.lists.getByTitle(SOURCE)
            .items.filter("Title eq 'SPJS Itest Source A'")
            .select('Id')<Array<{ Id: number }>>()
        )[0].Id
        const multi = afterFirst[UNRESOLVED].items.find(
          (item) => item.values.Title === 'SPJS Itest Unresolved multi'
        )
        assert.ok(multi, 'the row is written')
        assert.deepEqual(valuesOf(multi).SPJSItestDrLookups, [sourceA])
        assert.ok(
          logged(first).some((message) =>
            message.includes('SPJS Itest missing multi')
          ),
          `nothing logged names the unresolved value: ${JSON.stringify(
            logged(first)
          )}`
        )
      }
    )
  })
})
