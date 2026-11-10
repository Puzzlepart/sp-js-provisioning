import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  expectedOf,
  PACKAGE_CASES,
  unlisted
} from '../integration/packages.expected'

/**
 * `tests/integration/packages.test.ts` holds each hosting package's run to an exact list of what
 * the Node run cannot do (`packages.expected.ts`). These tests check the list's two functions
 * offline, with lines as the engine and `verify` write them.
 */
describe('packages: what the Node run cannot do', () => {
  const testCase = {
    pkg: 'pp-example',
    mode: 'hub' as const,
    lists: {
      'My List (1)': {
        bindings: ['0x0100AB'],
        fieldRefs: ['GtA'],
        views: { 'All items': ['GtA', 'GtC'] },
        rows: true
      }
    }
  }
  const logged = [
    '[Lists] (Lists): (processContentTypeBinding): Failed to add Content Type 0x0100AB to list My List (1): Error making HttpClient request in queryable [500] Internal Server Error',
    "[Lists] (Lists): (processFieldRef): Field ref '0e82c395-9ed9-43f9-871a-208215b18558' (GtA) for list My List (1): no column with that id on the list or the web, so it is not added.",
    '[Lists] (Lists): (processViewFields): Failed to add field GtA to view All items: Error making HttpClient request in queryable [400] Bad Request',
    '[Lists] (Lists): (processViewFields): Failed to add field GtC to view All items: Error making HttpClient request in queryable [400] Bad Request',
    '[Lists] (Lists): (processListDataRows): Failed to provision row 3 (Title=x) in list My List (1): Error making HttpClient request in queryable [400] Bad Request'
  ]
  const found = [
    '[Lists] My List (1)/ContentTypeBindings/0X0100AB: expected "bound", found "not bound"',
    '[Lists] My List (1)/FieldRefs/GtA: expected "a field", found "none"',
    '[Lists] My List (1)/views/All items/fields: expected ["LinkTitle","GtA","GtC"], found ["LinkTitle"]',
    '[Lists] My List (1)/DataRows/GtA: expected "a column of the list", found "no such column"',
    '[Lists] My List (1)/DataRows/0 (Title=x): expected "an item", found "none"'
  ]

  it('derives the errors a run logs and the findings verify reports from what a list lacks', () => {
    const { errors, findings } = expectedOf(testCase)
    assert.deepEqual(unlisted(logged, errors), { unexpected: [], missing: [] })
    assert.deepEqual(unlisted(found, findings), {
      unexpected: [],
      missing: []
    })
    assert.equal(errors.length, 4)
    assert.equal(findings.length, 4)
  })

  it('shows a new failure and a fixed one', () => {
    const { errors, findings } = expectedOf(testCase)
    const more = unlisted(
      [
        ...logged,
        '[Lists] (Lists): (processViewFields): Failed to add field GtNew to view All items: Error',
        '[Security] (Security): (processRoleAssignment): Failed'
      ],
      errors
    )
    assert.deepEqual(more.unexpected, [
      '[Lists] (Lists): (processViewFields): Failed to add field GtNew to view All items: Error',
      '[Security] (Security): (processRoleAssignment): Failed'
    ])
    assert.deepEqual(more.missing, [])
    // The binding works one day (the JSOM exit): its entry has no line and must be removed.
    const fewer = unlisted(found.slice(1), findings)
    assert.deepEqual(fewer.unexpected, [])
    assert.equal(fewer.missing.length, 1)
    assert.match(fewer.missing[0], /ContentTypeBindings/)
    assert.match(fewer.missing[0], /cannot bind 0x0100AB/)
    // Another list with the same view title does not match.
    assert.deepEqual(
      unlisted(
        ['[Lists] Other/views/All items/fields: expected [], found []'],
        findings
      ).unexpected.length,
      1
    )
  })

  it('adds the other expected errors and findings of a case', () => {
    const extra = {
      handler: 'Lists',
      pattern: /^\[Lists\] L\/views\/V\/JSLink: /,
      why: 'kept by SharePoint'
    }
    const { errors, findings } = expectedOf({
      pkg: 'p',
      mode: 'hub',
      lists: {},
      findings: [extra]
    })
    assert.deepEqual(errors, [])
    assert.deepEqual(findings, [extra])
  })

  it('lists, for every case, what each list lacks', () => {
    assert.ok(PACKAGE_CASES.length >= 9)
    for (const listed of PACKAGE_CASES) {
      assert.ok(listed.lists, `${listed.pkg} ${listed.mode}: no lists`)
      for (const title of Object.keys(listed.lists))
        assert.ok(
          Object.keys(listed.lists[title]).length > 0,
          `${listed.pkg} ${listed.mode}: ${title} lacks nothing`
        )
    }
    // pp-testprosjekt is tested in every mode and both locales.
    assert.deepEqual(
      PACKAGE_CASES.filter((listed) => listed.pkg === 'pp-testprosjekt').map(
        (listed) => `${listed.mode} ${listed.lang || 'nb-NO'}`
      ),
      [
        'hub nb-NO',
        'hub en-US',
        'cloud nb-NO',
        'cloud en-US',
        'project nb-NO',
        'project en-US'
      ]
    )
  })
})
