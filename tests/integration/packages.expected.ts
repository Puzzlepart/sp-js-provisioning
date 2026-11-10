import { Mode } from '../../debug/lib/cli'

/**
 * What the Node run of a hosting package cannot do, per case of `packages.test.ts`
 * (docs/specs/integration-testing.md, Fixtures). Each list says what it lacks afterwards; the suite
 * derives from it the errors the run must log and the differences `verify` must find, both exactly:
 * a new failure and a fixed one both fail the case. Generated from the runs of 2026-11-10 and kept
 * by hand since; remove an entry when what causes it is fixed (the JSOM exit, mostly).
 */

/**
 * What a list lacks after the run.
 */
export interface IListGaps {
  /**
   * Content type ids the list could not bind: the content type is not on the web. ContentTypes
   * runs through JSOM, skipped in Node until the JSOM exit; other ids are PP365's own content
   * types, which a PP365 site has and a scratch web does not.
   */
  bindings?: string[]
  /**
   * Field refs to columns that are neither on the list nor on the web: they come with a content
   * type the list could not bind, or are PP365's own site columns.
   */
  fieldRefs?: string[]
  /** The columns each view could not get (the list lacks them), by view title. */
  views?: Record<string, string[]>
  /** The rows name columns the list lacks: none of them is added. */
  rows?: boolean
}

/**
 * An expected error or finding that is not a list's gap: the handler, a pattern for the line and
 * why it is expected.
 */
export interface IExpected {
  handler: string
  pattern: RegExp
  why: string
}

/**
 * A package, a mode, a locale and the extensions, with what the Node run cannot do. Without
 * `lists` (`INTEGRATION_PACKAGES=all`) the errors and findings are reported, not asserted: the run
 * must only not stop.
 */
export interface IPackageCase {
  pkg: string
  mode: Mode
  lang?: string
  /** `--mode project`: the extensions (default: the ones the setup selects to begin with). */
  extensions?: string[]
  lists?: Record<string, IListGaps>
  /** Expected errors beside the lists' gaps. */
  errors?: IExpected[]
  /** Expected `verify` findings beside the lists' gaps. */
  findings?: IExpected[]
}

/**
 * `pp-testprosjekt`, which uses every feature of the package format, as a hub import, a cloud
 * template and a project setup, in both locales; the project template with pages and the extension
 * that sets a welcome page (pp-enkel-prosjektmal); the extension with a menu two levels deep
 * (pp-forskningsprosjekt); and the hub template with list permissions (pp-veiprosjekt).
 */
export const PACKAGE_CASES: IPackageCase[] = [
  {
    pkg: 'pp-testprosjekt',
    mode: 'hub',
    lists: {
      Prosjekter: { bindings: ['0x0100CBEAE584066D4D87970DD5EDFDFE7F06'] },
      Testsjekkliste: {
        bindings: ['0x0100FFE226AE6A8C4F38A2DAB5E451711F4E'],
        views: {
          'Alle elementer': ['GtSortOrder', 'GtProjectPhase', 'GtTestText']
        },
        rows: true
      }
    }
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'hub',
    lang: 'en-US',
    lists: {
      Projects: { bindings: ['0x0100CBEAE584066D4D87970DD5EDFDFE7F06'] },
      'Test Checklist': {
        bindings: ['0x0100FFE226AE6A8C4F38A2DAB5E451711F4E'],
        views: { 'All items': ['GtSortOrder', 'GtProjectPhase', 'GtTestText'] },
        rows: true
      }
    }
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'cloud',
    lists: {
      Prosjekter: { bindings: ['0x0100CBEAE584066D4D87970DD5EDFDFE7F06'] }
    }
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'cloud',
    lang: 'en-US',
    lists: {
      Projects: { bindings: ['0x0100CBEAE584066D4D87970DD5EDFDFE7F06'] }
    }
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'project',
    lists: {
      Testprosjektliste: {
        bindings: ['0x01009A8EC06C79424B1B99D0FC4E61331392'],
        views: { 'Alle elementer': ['GtTestChoice', 'GtTestProjectNote'] },
        rows: true
      }
    }
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'project',
    lang: 'en-US',
    lists: {
      'Test Project List': {
        bindings: ['0x01009A8EC06C79424B1B99D0FC4E61331392'],
        views: { 'All items': ['GtTestChoice', 'GtTestProjectNote'] },
        rows: true
      }
    }
  },
  {
    pkg: 'pp-enkel-prosjektmal',
    mode: 'project',
    extensions: ['enkelt-prosjekt'],
    lists: {
      Fasesjekkliste: {
        bindings: ['0x0100486B1F8AEA24486FBA1C1BA9146C360C'],
        views: {
          'Alle elementer': [
            'GtSortOrder',
            'GtProjectPhase',
            'GtChecklistStatus',
            'GtComment'
          ],
          Arkivert: ['GtSortOrder', 'GtProjectPhase', 'GtComment'],
          'Etter status': ['GtSortOrder', 'GtProjectPhase', 'GtComment']
        }
      },
      Usikkerhet: {
        bindings: [
          '0x0100A87AE71CBF2643A6BC9D0948BD2EE89701',
          '0x0100A87AE71CBF2643A6BC9D0948BD2EE89702'
        ],
        views: {
          'Alle elementer': [
            'GtRiskDescription',
            'GtRiskProbability',
            'GtRiskConsequence',
            'GtRiskFactor',
            'GtRiskStatus',
            'GtRiskStrategy',
            'GtRiskAction',
            'GtRiskProbabilityPostAction',
            'GtRiskConsequencePostAction',
            'GtRiskFactorPostAction'
          ],
          'Aktuelle risikoer': [
            'GtRiskProximity',
            'GtRiskAction',
            'GtRiskStrategy',
            'GtRiskFactor',
            'GtRiskFactorPostAction'
          ],
          'Aktuelle muligheter': [
            'GtRiskProximity',
            'GtRiskAction',
            'GtRiskStrategy',
            'GtRiskFactor',
            'GtRiskFactorPostAction'
          ],
          'Uakseptable risikoer': [
            'GtRiskDescription',
            'GtRiskProbability',
            'GtRiskConsequence',
            'GtRiskFactor',
            'GtRiskStatus',
            'GtRiskStrategy',
            'GtRiskAction'
          ],
          'Avklart risiko': [
            'GtRiskDescription',
            'GtRiskProbability',
            'GtRiskConsequence',
            'GtRiskFactor',
            'GtRiskStatus',
            'GtRiskStrategy',
            'GtRiskAction',
            'GtRiskProbabilityPostAction',
            'GtRiskConsequencePostAction',
            'GtRiskFactorPostAction'
          ],
          'Uten ansvarlig': [
            'GtRiskDescription',
            'GtRiskFactor',
            'GtRiskStatus',
            'GtRiskAction',
            'GtActionResponsible'
          ],
          'Kun risikoer': [
            'GtRiskDescription',
            'GtRiskProbability',
            'GtRiskConsequence',
            'GtRiskFactor',
            'GtRiskStatus',
            'GtRiskStrategy',
            'GtRiskAction',
            'GtRiskProbabilityPostAction',
            'GtRiskConsequencePostAction',
            'GtRiskFactorPostAction'
          ],
          'Kun muligheter': [
            'GtRiskDescription',
            'GtRiskProbability',
            'GtRiskConsequence',
            'GtRiskFactor',
            'GtRiskStatus',
            'GtRiskStrategy',
            'GtRiskAction',
            'GtRiskProbabilityPostAction',
            'GtRiskConsequencePostAction',
            'GtRiskFactorPostAction'
          ]
        }
      }
    }
  },
  {
    pkg: 'pp-forskningsprosjekt',
    mode: 'project',
    extensions: ['forskningsmal'],
    lists: {
      Fasesjekkliste: {
        bindings: ['0x0100486B1F8AEA24486FBA1C1BA9146C360C'],
        fieldRefs: ['GtcPhaseReference'],
        views: {
          'Alle elementer': [
            'GtSortOrder',
            'GtProjectPhase',
            'GtChecklistStatus',
            'GtComment',
            'GtcPhaseReference'
          ],
          Arkivert: [
            'GtSortOrder',
            'GtProjectPhase',
            'GtComment',
            'GtcPhaseReference'
          ],
          'Etter status': [
            'GtSortOrder',
            'GtProjectPhase',
            'GtComment',
            'GtcPhaseReference'
          ]
        }
      },
      Prosjektleveranser: {
        bindings: ['0x0100D7B74DE815F946D3B0F99D19F9B36B68'],
        fieldRefs: ['GtcDeliveryNumber'],
        views: {
          'Alle elementer': [
            'GtDeliveryDescription',
            'GtDeliveryStartTime',
            'GtDeliveryEndTime',
            'GtDeliveryStatus',
            'GtDeliveryStatusComment'
          ]
        }
      },
      Publiseringer: {
        bindings: ['0x01004B2B9889EA2C462FBE9B54B0D462E96B'],
        views: {
          'Alle elementer': [
            'GtcPubResultId',
            'GtcPubChannel',
            'GtcPubRole',
            'GtcPubIsPublished',
            'GtcPubDate',
            'GtcPubLink'
          ]
        }
      }
    }
  },
  {
    pkg: 'pp-veiprosjekt',
    mode: 'hub',
    lists: {
      Prosjekter: { bindings: ['0x0100805E9E4FEAAB4F0EABAB2600D30DB70C0E'] },
      Prosjektdata: { bindings: ['0x0100805E9E4FEAAB4F0EABAB2600D30DB70C0E'] },
      'Fasesjekkliste Vei': {
        bindings: ['0x0100486B1F8AEA24486FBA1C1BA9146C360C'],
        fieldRefs: ['GtChecklistStatus', 'GtComment'],
        views: { 'Alle elementer': ['GtSortOrder'] },
        rows: true
      },
      'Planneroppgaver Vei': {
        bindings: ['0x01000F43B5C9E86B4D7D9808100EC27E844A'],
        fieldRefs: ['GtSortOrder'],
        views: {
          'Alle elementer': [
            'GtSortOrder',
            'GtDescription',
            'GtCategory',
            'GtChecklist',
            'GtPlannerTags',
            'GtAttachments',
            'GtPlannerPreviewType'
          ]
        },
        rows: true
      },
      Prosjektkolonner: { rows: true },
      Porteføljevisninger: { rows: true }
    },
    findings: [
      {
        handler: 'Lists',
        pattern: /^\[Lists\] Fasesjekkliste Vei\/views\/Alle elementer\/JSLink: /,
        why:
          "SharePoint answers the view update with success and keeps the view's default JSLink (2026-11-10; the reserved site blocks custom script)"
      }
    ]
  }
]

/** A text as a regular expression that matches it literally. */
const literal = (text: string) => text.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')

/**
 * The errors a run must log and the findings `verify` must report for what a case's lists lack
 * (`packages.expected.ts`), plus the case's other expected errors and findings. An error line is
 * `[Lists] (Lists): (<scope>): <message>`, a finding line `[Lists] <path>: expected ..., found ...`.
 *
 * @param testCase - The case
 */
export function expectedOf(
  testCase: IPackageCase
): { errors: IExpected[]; findings: IExpected[] } {
  const errors: IExpected[] = [...(testCase.errors || [])]
  const findings: IExpected[] = [...(testCase.findings || [])]
  const error = (pattern: string, why: string) =>
    errors.push({ handler: 'Lists', pattern: new RegExp(pattern), why })
  const finding = (path: string, why: string) =>
    findings.push({
      handler: 'Lists',
      pattern: new RegExp(`^\\[Lists\\] ${path}`),
      why
    })
  for (const title of Object.keys(testCase.lists || {})) {
    const gaps: IListGaps = testCase.lists[title]
    const list = literal(title)
    for (const id of gaps.bindings || []) {
      const why = `${title} cannot bind ${id}: the content type is not on the web`
      error(
        `processContentTypeBinding\\): Failed to add Content Type ${id} to list ${list}: `,
        why
      )
      finding(`${list}/ContentTypeBindings/${id.toUpperCase()}: `, why)
    }
    if (gaps.fieldRefs) {
      const columns = gaps.fieldRefs.join('|')
      const why = `${title}: field refs to columns that are neither on the list nor the web`
      error(
        `processFieldRef\\): Field ref '[^']+' \\((${columns})\\) for list ${list}: no column`,
        why
      )
      finding(`${list}/FieldRefs/(${columns}): `, why)
    }
    for (const view of Object.keys(gaps.views || {})) {
      const why = `${title}, view ${view}: the list lacks the columns`
      error(
        `processViewFields\\): Failed to add field (${gaps.views[view].join(
          '|'
        )}) to view ${literal(view)}: `,
        why
      )
      finding(`${list}/views/${literal(view)}/fields: `, why)
    }
    if (gaps.rows) {
      const why = `${title}: the rows name columns the list lacks`
      error(
        `processListDataRows\\): Failed to provision row \\d+ \\([^)]*\\) in list ${list}: `,
        why
      )
      finding(`${list}/DataRows/`, why)
    }
  }
  return { errors, findings }
}

/**
 * What differs from an exact list: the actual lines no entry matches, and the entries no line
 * matches. The list of what the Node run cannot do stays exact, so a new failure, and a fixed one,
 * both show.
 *
 * @param actual - The lines, `[Handler] text`
 * @param expected - The expected entries
 */
export function unlisted(
  actual: string[],
  expected: IExpected[]
): { unexpected: string[]; missing: string[] } {
  const matches = (line: string, entry: IExpected) =>
    line.startsWith(`[${entry.handler}] `) && entry.pattern.test(line)
  return {
    unexpected: actual.filter(
      (line) => !expected.some((entry) => matches(line, entry))
    ),
    missing: expected
      .filter((entry) => !actual.some((line) => matches(line, entry)))
      .map((entry) => `[${entry.handler}] ${entry.pattern} (${entry.why})`)
  }
}
