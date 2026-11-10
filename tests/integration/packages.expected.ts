import { Mode } from '../../debug/lib/cli'

/**
 * What a hosting package's run on a scratch web of the reserved site cannot do, per case of
 * `packages.test.ts` (docs/specs/integration-testing.md, Fixtures): what a scratch web lacks that a
 * PP365 hub or project site has (PP365's own site columns and content types), and what the reserved
 * site refuses (a property bag write, which needs custom script). Each list says what it lacks
 * afterwards; the suite derives from it the errors the run must log and the differences `verify`
 * must find, both exactly: a new failure and a fixed one both fail the case. Generated from the
 * runs of 2026-11-11 after the JSOM exit (`INTEGRATION_PACKAGES_REPORT=true` prints every line);
 * remove an entry when what causes it is fixed.
 */

/**
 * What a list lacks after the run.
 */
export interface IListGaps {
  /**
   * Content type ids the list could not bind: PP365's own content types, which a PP365 site has and
   * a scratch web does not.
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
 * A content type whose field refs name columns a scratch web lacks (PP365's own site columns): they
 * are left out, each with an error.
 */
export interface IContentTypeGaps {
  /** The content type's name, as the error names it. */
  name: string
  /** The field refs left out, by name. */
  fieldRefs: string[]
}

/**
 * A package, a mode, a locale and the extensions, with what the run cannot do. Without
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
  /** Content types whose field refs name columns a scratch web lacks, by content type id. */
  contentTypes?: Record<string, IContentTypeGaps>
  /**
   * Property bag keys a site that blocks custom script refuses to write (the reserved site does):
   * expected only there.
   */
  propertyBag?: string[]
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
      Testsjekkliste: {
        rows: true,
        views: {
          'Alle elementer': ['GtSortOrder', 'GtProjectPhase']
        }
      }
    },
    contentTypes: {
      '0x0100CBEAE584066D4D87970DD5EDFDFE7F06': {
        name: 'Prosjekt (Test)',
        fieldRefs: [
          'GtProjectNumber',
          'GtProjectPhase',
          'GtProjectManager',
          'GtProjectOwner',
          'GtStartDate',
          'GtEndDate',
          'GtProjectGoals'
        ]
      },
      '0x0100FFE226AE6A8C4F38A2DAB5E451711F4E': {
        name: 'Testsjekkpunkt',
        fieldRefs: ['GtSortOrder', 'GtProjectPhase']
      }
    },
    propertyBag: ['GtTestPackage']
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'hub',
    lang: 'en-US',
    lists: {
      'Test Checklist': {
        rows: true,
        views: {
          'All items': ['GtSortOrder', 'GtProjectPhase']
        }
      }
    },
    contentTypes: {
      '0x0100CBEAE584066D4D87970DD5EDFDFE7F06': {
        name: 'Project (Test)',
        fieldRefs: [
          'GtProjectNumber',
          'GtProjectPhase',
          'GtProjectManager',
          'GtProjectOwner',
          'GtStartDate',
          'GtEndDate',
          'GtProjectGoals'
        ]
      },
      '0x0100FFE226AE6A8C4F38A2DAB5E451711F4E': {
        name: 'Test Checkpoint',
        fieldRefs: ['GtSortOrder', 'GtProjectPhase']
      }
    },
    propertyBag: ['GtTestPackage']
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'cloud',
    lists: {},
    contentTypes: {
      '0x0100CBEAE584066D4D87970DD5EDFDFE7F06': {
        name: 'Prosjekt (Test)',
        fieldRefs: [
          'GtProjectNumber',
          'GtProjectPhase',
          'GtProjectManager',
          'GtProjectOwner',
          'GtStartDate',
          'GtEndDate',
          'GtProjectGoals'
        ]
      },
      '0x0100FFE226AE6A8C4F38A2DAB5E451711F4E': {
        name: 'Testsjekkpunkt',
        fieldRefs: ['GtSortOrder', 'GtProjectPhase']
      }
    }
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'cloud',
    lang: 'en-US',
    lists: {},
    contentTypes: {
      '0x0100CBEAE584066D4D87970DD5EDFDFE7F06': {
        name: 'Project (Test)',
        fieldRefs: [
          'GtProjectNumber',
          'GtProjectPhase',
          'GtProjectManager',
          'GtProjectOwner',
          'GtStartDate',
          'GtEndDate',
          'GtProjectGoals'
        ]
      },
      '0x0100FFE226AE6A8C4F38A2DAB5E451711F4E': {
        name: 'Test Checkpoint',
        fieldRefs: ['GtSortOrder', 'GtProjectPhase']
      }
    }
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'project',
    lists: {
      Testprosjektliste: {
        rows: true,
        views: {
          'Alle elementer': ['GtTestChoice']
        }
      }
    },
    contentTypes: {
      '0x01009A8EC06C79424B1B99D0FC4E61331392': {
        name: 'Testelement',
        fieldRefs: ['GtTestChoice']
      }
    },
    propertyBag: ['GtTestProjectProvisioned']
  },
  {
    pkg: 'pp-testprosjekt',
    mode: 'project',
    lang: 'en-US',
    lists: {
      'Test Project List': {
        rows: true,
        views: {
          'All items': ['GtTestChoice']
        }
      }
    },
    contentTypes: {
      '0x01009A8EC06C79424B1B99D0FC4E61331392': {
        name: 'Test Item',
        fieldRefs: ['GtTestChoice']
      }
    },
    propertyBag: ['GtTestProjectProvisioned']
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
    },
    contentTypes: {
      '0x0100486B1F8AEA24486FBA1C1BA9146C360C': {
        name: 'Sjekkpunkt',
        fieldRefs: [
          'GtProjectPhase',
          'GtChecklistStatus',
          'GtComment',
          'GtSortOrder',
          'GtcPhaseReference'
        ]
      },
      '0x01004B2B9889EA2C462FBE9B54B0D462E96B': {
        name: 'Publiseringelement',
        fieldRefs: [
          'GtcPubResultId',
          'GtcPubChannel',
          'GtcPubRole',
          'GtcPubIsPublished',
          'GtcPubDate',
          'GtcPubLink'
        ]
      },
      '0x0100D7B74DE815F946D3B0F99D19F9B36B68': {
        name: 'Prosjektleveranse',
        fieldRefs: [
          'GtDeliveryDescription',
          'GtDeliveryStartTime',
          'GtDeliveryEndTime',
          'GtDeliveryQualityExpectations',
          'GtDeliveryAcceptanceCriteria',
          'GtDeliveryQualityTolerance',
          'GtDeliveryQualityResponsible',
          'GtDeliveryAcceptanceMethod',
          'GtDeliveryAcceptanceResponsible',
          'GtDeliveryAcceptanceDate',
          'GtDeliveryStatus',
          'GtDeliveryStatusComment',
          'GtTag',
          'GtcDeliveryNumber'
        ]
      }
    }
  },
  {
    pkg: 'pp-veiprosjekt',
    mode: 'hub',
    lists: {
      'Fasesjekkliste Vei': {
        bindings: ['0x0100486B1F8AEA24486FBA1C1BA9146C360C'],
        fieldRefs: ['GtChecklistStatus', 'GtComment'],
        rows: true,
        views: {
          'Alle elementer': ['GtSortOrder']
        }
      },
      'Planneroppgaver Vei': {
        bindings: ['0x01000F43B5C9E86B4D7D9808100EC27E844A'],
        fieldRefs: ['GtSortOrder'],
        rows: true,
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
        }
      },
      Prosjektkolonner: {
        rows: true
      },
      Porteføljevisninger: {
        rows: true
      }
    },
    contentTypes: {
      '0x0100805E9E4FEAAB4F0EABAB2600D30DB70C0E': {
        name: 'Prosjekt (Vei)',
        fieldRefs: [
          'GtProjectFinanceName',
          'GtProjectNumber',
          'GtArchiveReference',
          'GtProjectServiceArea',
          'GtProjectType',
          'GtUNSustDevGoals',
          'GtProjectPhase',
          'GtProjectManager',
          'GtProjectOwner',
          'GtProjectSupport',
          'GtStartDate',
          'GtEndDate',
          'GtProjectGoals',
          'GtBudgetTotal',
          'GtCostsTotal',
          'GtProjectForecast',
          'GtBudgetLastReportDate',
          'GtGroupId',
          'GtSiteId',
          'GtSiteUrl',
          'GtLastSyncTime',
          'GtProjectLifecycleStatus',
          'GtProjectPhaseText',
          'GtProjectServiceAreaText',
          'GtProjectTypeText',
          'GtUNSustDevGoalsText',
          'GtProjectAdminRoles',
          'GtGainsResponsible',
          'GtProjectTemplate',
          'GtParentProjects',
          'GtChildProjects',
          'GtIsParentProject',
          'GtIsProgram'
        ]
      }
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

/**
 * The parent of a content type id: the id without its last `00` and GUID (`0x0100<32 hex>` is a
 * child of `0x01`), or without its last two digits (`<id>01` is a child of `<id>`).
 *
 * @param id - The content type id
 */
export function parentContentTypeId(id: string): string | undefined {
  const bare = id.replace(/^0x/i, '')
  if (bare.length <= 2) return undefined
  const parent =
    bare.length >= 36 && /^00[\da-f]{32}$/i.test(bare.slice(-34))
      ? bare.slice(0, -34)
      : bare.slice(0, -2)
  return `0x${parent}`
}

/** A text as a regular expression that matches it literally. */
const literal = (text: string) => text.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')

/**
 * The errors a run must log and the findings `verify` must report for what a case's lists and
 * content types lack and what the site refuses (`packages.expected.ts`), plus the case's other
 * expected errors and findings. An error line is `[<Handler>] (<Handler>): (<scope>): <message>`, a
 * finding line `[<Handler>] <path>: expected ..., found ...`.
 *
 * @param testCase - The case
 * @param site - Whether the site blocks custom script (the property bag is refused); default yes
 */
export function expectedOf(
  testCase: IPackageCase,
  site: { customScriptBlocked?: boolean } = {}
): { errors: IExpected[]; findings: IExpected[] } {
  const errors: IExpected[] = [...(testCase.errors || [])]
  const findings: IExpected[] = [...(testCase.findings || [])]
  for (const id of Object.keys(testCase.contentTypes || {})) {
    const { name, fieldRefs } = testCase.contentTypes[id]
    const columns = fieldRefs.map(literal).join('|')
    const why = `${name}: field refs to PP365's own site columns, which a scratch web lacks`
    errors.push({
      handler: 'ContentTypes',
      pattern: new RegExp(
        `processContentTypeFieldRefs\\): Field ref (${columns}) of content type \\[${literal(
          name
        )}\\] \\(${literal(id)}\\) is left out: the web has no site column`
      ),
      why
    })
    findings.push({
      handler: 'ContentTypes',
      pattern: new RegExp(
        `^\\[ContentTypes\\] ${literal(name)} \\(${literal(
          id.toUpperCase()
        )}\\)/FieldRefs/(${columns}): `
      ),
      why
    })
  }
  if (
    testCase.propertyBag &&
    testCase.propertyBag.length > 0 &&
    site.customScriptBlocked !== false
  ) {
    const keys = testCase.propertyBag.map(literal)
    const why =
      'the site blocks custom script, so it refuses the property bag write'
    errors.push({
      handler: 'PropertyBagEntries',
      pattern: new RegExp(
        `ProvisionObjects\\): Property bag entries ${keys.join(
          ', '
        )}(, vti_indexedpropertykeys)? were not written: SharePoint refused the write, as it does on a site that blocks custom script`
      ),
      why
    })
    findings.push({
      handler: 'PropertyBagEntries',
      pattern: new RegExp(
        `^\\[PropertyBagEntries\\] (${keys.join('|')})(/Indexed)?: `
      ),
      why
    })
  }
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
