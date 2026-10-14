import { ISiteGroup, IWeb } from '@pnp/sp/presets/all'

/**
 * Role type kind as used by `roleDefinitions.getByType` (mirrors pnpjs'
 * `RoleTypeKind`, which the package does not re-export).
 */
type RoleTypeKind = 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7

/**
 * Well-known English role definition names mapped to `RoleTypeKind`
 * (Reader=2, Contributor=3, WebDesigner=4, Administrator=5, Editor=6). Other
 * role definitions, including `View Only` and every localized name, must be
 * named exactly as they exist on the web.
 */
const wellKnownRoleTypes: { [name: string]: RoleTypeKind } = {
  'full control': 5,
  administrator: 5,
  design: 4,
  edit: 6,
  contribute: 3,
  read: 2
}

/**
 * Resolves a role assignment principal to a principal id: associated group
 * tokens (`{associatedownergroupid}`, `{associatedmembergroupid}`,
 * `{associatedvisitorgroupid}`), a numeric id, a site group name, or a user
 * login name (ensured). Other tokens must be replaced by the caller.
 *
 * @param web - The web
 * @param principal - The principal
 * @param log - Called with a message when a name is not a site group and is
 * resolved as a user instead
 */
export async function resolvePrincipalId(
  web: IWeb,
  principal: string | number,
  log?: (message: string) => void
): Promise<number> {
  if (typeof principal === 'number') return principal
  const value = String(principal).trim()
  if (/^\d+$/.test(value)) return Number(value)
  const token = value.replace(/[{}]/g, '').toLowerCase()
  const associatedGroups: { [key: string]: () => ISiteGroup } = {
    associatedownergroupid: () => web.associatedOwnerGroup,
    associatedmembergroupid: () => web.associatedMemberGroup,
    associatedvisitorgroupid: () => web.associatedVisitorGroup
  }
  if (associatedGroups[token]) {
    const group = await associatedGroups[token]().select('Id')<{
      Id: number
    }>()
    return group.Id
  }
  try {
    const group = await web.siteGroups
      .getByName(value.replace(/'/g, "''"))
      .select('Id')<{ Id: number }>()
    return group.Id
  } catch (groupError) {
    if (log)
      log(
        `No site group named '${value}', resolving it as a user: ${
          (groupError && groupError.message) || groupError
        }`
      )
    try {
      const user = await web.ensureUser(value)
      return user.Id
    } catch (userError) {
      throw new Error(
        `Could not resolve principal '${value}' as a site group (${
          (groupError && groupError.message) || groupError
        }) or as a user (${(userError && userError.message) || userError})`
      )
    }
  }
}

/**
 * Resolves a role definition by (localized) name, falling back to the
 * well-known role type for the English names (Full Control, Design, Edit,
 * Contribute, Read).
 *
 * @param web - The web
 * @param roleDefinition - The role definition name
 */
export async function resolveRoleDefinitionId(
  web: IWeb,
  roleDefinition: string
): Promise<number> {
  try {
    const byName = await web.roleDefinitions
      .getByName(roleDefinition)
      .select('Id')<{ Id: number }>()
    return byName.Id
  } catch (error) {
    const roleTypeKind =
      wellKnownRoleTypes[String(roleDefinition).trim().toLowerCase()]
    if (roleTypeKind === undefined) throw error
    const byType = await web.roleDefinitions
      .getByType(roleTypeKind)
      .select('Id')<{ Id: number }>()
    return byType.Id
  }
}
