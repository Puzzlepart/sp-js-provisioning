import { ISiteConnection } from '../connect'

/**
 * What the Features handler changes on a web: the web-scoped features that are active there.
 */
export interface IFeaturesState {
  /** `DefinitionId` of every active web feature, lower case, sorted. */
  active: string[]
}

/**
 * Reads back the web's active features. `_api/web/features` lists only active ones; an inactive
 * feature is absent from it (its `getById` answers `odata.null`, not a 404), so membership is the
 * state.
 *
 * @param conn - The web
 */
export async function readFeatures(
  conn: ISiteConnection
): Promise<IFeaturesState> {
  const features = await conn.web.features.select('DefinitionId')<
    Array<{ DefinitionId: string }>
  >()
  return {
    active: features.map((feature) => feature.DefinitionId.toLowerCase()).sort()
  }
}
