import { ISiteConnection } from '../connect'

/**
 * A site group as SharePoint stores it, without ids.
 */
export interface IGroupState {
  Title: string
  Description: string
  Owner: string
  AllowMembersEditMembership: boolean
  AllowRequestToJoinLeave: boolean
  AutoAcceptRequestToJoinLeave: boolean
  OnlyAllowMembersViewMembership: boolean
  RequestToJoinLeaveEmailSetting: string
  /** Login names, lower case, sorted. */
  members: string[]
}

/**
 * What the Security handler provisions on a web: the named site groups, whether the web has
 * unique permissions, and its role assignments (member title -> role names, sorted).
 */
export interface ISecurityState {
  groups: IGroupState[]
  hasUniqueRoleAssignments: boolean
  roleAssignments: Record<string, string[]>
}

/**
 * Reads back the site groups with the given titles and the web's permissions.
 *
 * @param conn - The web
 * @param groupTitles - The groups to read, by title
 */
export async function readSecurity(
  conn: ISiteConnection,
  groupTitles: string[]
): Promise<ISecurityState> {
  const { web } = conn
  const groups: IGroupState[] = []
  for (const title of groupTitles) {
    const group = web.siteGroups.getByName(title.replace(/'/g, "''"))
    const info: any = await group
      .expand('Owner')
      .select(
        'Title',
        'Description',
        'Owner/Title',
        'AllowMembersEditMembership',
        'AllowRequestToJoinLeave',
        'AutoAcceptRequestToJoinLeave',
        'OnlyAllowMembersViewMembership',
        'RequestToJoinLeaveEmailSetting'
      )()
    const users = await group.users.select('LoginName')<
      Array<{ LoginName: string }>
    >()
    groups.push({
      Title: info.Title,
      Description: info.Description ?? '',
      Owner: info.Owner && info.Owner.Title,
      AllowMembersEditMembership: info.AllowMembersEditMembership,
      AllowRequestToJoinLeave: info.AllowRequestToJoinLeave,
      AutoAcceptRequestToJoinLeave: info.AutoAcceptRequestToJoinLeave,
      OnlyAllowMembersViewMembership: info.OnlyAllowMembersViewMembership,
      RequestToJoinLeaveEmailSetting: info.RequestToJoinLeaveEmailSetting ?? '',
      members: users.map((user) => user.LoginName.toLowerCase()).sort()
    })
  }
  const { HasUniqueRoleAssignments } = await web.select(
    'HasUniqueRoleAssignments'
  )<{ HasUniqueRoleAssignments: boolean }>()
  const assignments = await web.roleAssignments.expand(
    'Member',
    'RoleDefinitionBindings'
  )<
    Array<{
      Member: { Title: string }
      RoleDefinitionBindings: Array<{ Name: string }>
    }>
  >()
  const roleAssignments: Record<string, string[]> = {}
  for (const assignment of assignments)
    roleAssignments[assignment.Member.Title] = assignment.RoleDefinitionBindings.map(
      (binding) => binding.Name
    ).sort()
  return {
    groups,
    hasUniqueRoleAssignments: HasUniqueRoleAssignments,
    roleAssignments
  }
}
