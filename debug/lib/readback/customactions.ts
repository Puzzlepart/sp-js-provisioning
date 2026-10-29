import { ISiteConnection } from '../connect'

/**
 * A web-scoped user custom action as SharePoint stores it. Empty strings stand for members
 * SharePoint returns as `null`; `ClientSideComponentId` is lower case.
 */
export interface ICustomActionState {
  /** Stable across runs unless the action is deleted and added again. */
  Id: string
  Title: string
  Name: string
  Description: string
  Location: string
  Sequence: number
  ClientSideComponentId: string
  ClientSideComponentProperties: string
  Url: string
  RegistrationId: string
  RegistrationType: number
  ScriptBlock: string
  ScriptSrc: string
}

/** The members {@link readCustomActions} reads. */
export const CUSTOM_ACTION_MEMBERS: Array<keyof ICustomActionState> = [
  'Id',
  'Title',
  'Name',
  'Description',
  'Location',
  'Sequence',
  'ClientSideComponentId',
  'ClientSideComponentProperties',
  'Url',
  'RegistrationId',
  'RegistrationType',
  'ScriptBlock',
  'ScriptSrc'
]

/**
 * Reads back the web's user custom actions whose title starts with a prefix, sorted by title (then
 * id, so duplicates show next to each other). The handler matches existing actions by `Title`,
 * so a duplicate title in the result means a run added an action twice.
 *
 * @param conn - The web
 * @param prefix - The title prefix (the fixtures' `SPJS Itest`)
 */
export async function readCustomActions(
  conn: ISiteConnection,
  prefix = 'SPJS Itest'
): Promise<ICustomActionState[]> {
  const actions = await conn.web.userCustomActions
    .filter(`startswith(Title,'${prefix.replace(/'/g, "''")}')`)
    .select(...CUSTOM_ACTION_MEMBERS)<Array<Record<string, any>>>()
  const text = (value: any): string =>
    value === null || value === undefined ? '' : String(value)
  return actions
    .map((action) => ({
      Id: text(action.Id).toLowerCase(),
      Title: text(action.Title),
      Name: text(action.Name),
      Description: text(action.Description),
      Location: text(action.Location),
      Sequence: Number(action.Sequence),
      ClientSideComponentId: text(action.ClientSideComponentId).toLowerCase(),
      ClientSideComponentProperties: text(action.ClientSideComponentProperties),
      Url: text(action.Url),
      RegistrationId: text(action.RegistrationId),
      RegistrationType: Number(action.RegistrationType),
      ScriptBlock: text(action.ScriptBlock),
      ScriptSrc: text(action.ScriptSrc)
    }))
    .filter((action) => action.Title.startsWith(prefix))
    .sort(
      (a, b) =>
        a.Title.localeCompare(b.Title, 'en') || a.Id.localeCompare(b.Id, 'en')
    )
}
