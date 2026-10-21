import { ISiteGroupInfo, IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'
import { IRoleAssignment, ISiteGroup, IWebSecurity } from '../schema'
import { processQuery, setGroupOwnerToGroupRequest } from '../util/csom'
import {
  resolvePrincipal,
  resolvePrincipalId,
  resolveRoleDefinitionId
} from '../util/security'
import { TokenHelper } from '../util/tokenhelper'
import { HandlerBase } from './handlerbase'

/**
 * Site group properties copied from the template to the group when they are
 * set. `Title` identifies the group and is handled separately.
 */
const siteGroupProperties: Array<keyof ISiteGroup & keyof ISiteGroupInfo> = [
  'Description',
  'AllowMembersEditMembership',
  'AllowRequestToJoinLeave',
  'AutoAcceptRequestToJoinLeave',
  'OnlyAllowMembersViewMembership',
  'RequestToJoinLeaveEmailSetting'
]

/**
 * Describes the Security Object Handler
 *
 * Creates or updates site groups (title, properties, members, owner) and
 * applies role inheritance and role assignments on the web. Runs before
 * `Lists`, so list `Security` can grant permissions to the groups created
 * here. Failures are logged per group, member and role assignment and never
 * abort the template.
 */
export class Security extends HandlerBase {
  public tokenHelper: TokenHelper
  /** Posts a CSOM request (a seam for unit tests, which have no real PnPjs web). */
  public processQuery: (
    web: IWeb,
    request: string
  ) => Promise<any[]> = processQuery

  /**
   * Creates a new instance of the Security class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('Security', config)
  }

  /**
   * Provisioning security
   *
   * @param web - The web
   * @param security - The security configuration
   * @param context - Provisioning context
   */
  public async ProvisionObjects(
    web: IWeb,
    security: IWebSecurity,
    context: ProvisioningContext
  ): Promise<void> {
    this.tokenHelper = new TokenHelper(context, this.config)
    super.scope_started()
    try {
      for (const siteGroup of security.SiteGroups || []) {
        await this.processSiteGroup(web, siteGroup)
      }
      await this.processWebSecurity(web, security)
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * Creates the site group, or updates it when a group with the same title
   * exists, then adds its members and sets its owner. The owner is set last, so
   * the provisioning user still owns a new group while its members are added.
   *
   * @param web - The web
   * @param gc - The site group configuration
   */
  private async processSiteGroup(web: IWeb, gc: ISiteGroup): Promise<void> {
    const title = this.tokenHelper.replaceTokens(gc.Title)
    super.log_info('processSiteGroup', `Processing site group ${title}.`)
    try {
      const properties = this.getSiteGroupProperties(gc)
      let groupId = await this.getSiteGroupId(web, title)
      if (groupId === undefined) {
        const created = await web.siteGroups.add({
          Title: title,
          ...properties
        })
        groupId = created.Id
        super.log_info(
          'processSiteGroup',
          `Site group ${title} created (${groupId}).`
        )
      } else if (Object.keys(properties).length > 0) {
        await web.siteGroups.getById(groupId).update(properties)
        super.log_info(
          'processSiteGroup',
          `Site group ${title} (${groupId}) exists, properties updated.`
        )
      }
      await this.processSiteGroupMembers(web, groupId, title, gc)
      if (gc.Owner) await this.processSiteGroupOwner(web, groupId, title, gc)
    } catch (error) {
      super.log_error(
        'processSiteGroup',
        `Failed to provision site group ${title}: ${
          (error && error.message) || error
        }`
      )
    }
  }

  /**
   * Returns the id of the site group with the title, or `undefined` when the
   * site collection has no such group.
   *
   * @param web - The web
   * @param title - The group title
   */
  private async getSiteGroupId(
    web: IWeb,
    title: string
  ): Promise<number | undefined> {
    try {
      const group = await web.siteGroups
        .getByName(title.replace(/'/g, "''"))
        .select('Id')<{ Id: number }>()
      return group.Id
    } catch (error) {
      if (error && error.status === 404) return
      throw error
    }
  }

  /**
   * Picks the group properties that are set in the template, with tokens
   * replaced in string values.
   *
   * @param gc - The site group configuration
   */
  private getSiteGroupProperties(gc: ISiteGroup): Partial<ISiteGroupInfo> {
    const properties: Partial<ISiteGroupInfo> = {}
    for (const key of siteGroupProperties) {
      const value = gc[key]
      if (value === undefined) continue
      ;(properties as any)[key] =
        typeof value === 'string'
          ? this.tokenHelper.replaceTokens(value)
          : value
    }
    return properties
  }

  /**
   * Adds the members to the group. With `ClearExistingMembers`, members that
   * are not in the template are removed afterwards - unless a member could not
   * be added, since its login name is then unknown and an existing member could
   * be removed by mistake.
   *
   * @param web - The web
   * @param groupId - The group id
   * @param title - The group title (for logging)
   * @param gc - The site group configuration
   */
  private async processSiteGroupMembers(
    web: IWeb,
    groupId: number,
    title: string,
    gc: ISiteGroup
  ): Promise<void> {
    const group = web.siteGroups.getById(groupId)
    const loginNames = new Set<string>()
    let failed = false
    for (const member of gc.Members || []) {
      const loginName = this.tokenHelper.replaceTokens(member)
      try {
        const user = await web.ensureUser(loginName)
        await group.users.add(user.LoginName)
        loginNames.add(user.LoginName.toLowerCase())
        super.log_info(
          'processSiteGroupMembers',
          `Member ${loginName} added to site group ${title}.`
        )
      } catch (error) {
        failed = true
        super.log_error(
          'processSiteGroupMembers',
          `Failed to add member ${loginName} to site group ${title}: ${
            (error && error.message) || error
          }`
        )
      }
    }
    if (!gc.ClearExistingMembers) return
    if (failed) {
      super.log_warn(
        'processSiteGroupMembers',
        `Existing members of site group ${title} are kept, since not all members could be added.`
      )
      return
    }
    const users = await group.users.select('Id', 'LoginName')<
      Array<{ Id: number; LoginName: string }>
    >()
    for (const user of users) {
      if (loginNames.has(user.LoginName.toLowerCase())) continue
      try {
        await group.users.removeById(user.Id)
        super.log_info(
          'processSiteGroupMembers',
          `Member ${user.LoginName} removed from site group ${title}.`
        )
      } catch (error) {
        super.log_error(
          'processSiteGroupMembers',
          `Failed to remove member ${
            user.LoginName
          } from site group ${title}: ${(error && error.message) || error}`
        )
      }
    }
  }

  /**
   * Sets the group owner and reads it back. The owner is resolved like a role assignment
   * principal, so it can be an associated group token, a site group title (including the group's
   * own) or a user login name. A group becomes the owner through CSOM (`Group.Owner`), since REST's
   * `SetUserAsOwner` answers success for a group id and changes nothing; a user through
   * `SetUserAsOwner`. An owner that did not change is logged as an error.
   *
   * @param web - The web
   * @param groupId - The group id
   * @param title - The group title (for logging)
   * @param gc - The site group configuration
   */
  private async processSiteGroupOwner(
    web: IWeb,
    groupId: number,
    title: string,
    gc: ISiteGroup
  ): Promise<void> {
    const owner = this.tokenHelper.replaceTokens(gc.Owner)
    try {
      const principal = await resolvePrincipal(web, owner, (message) =>
        this.log_info('resolvePrincipalId', message)
      )
      const isGroup =
        principal.isGroup ?? (await this.isSiteGroup(web, principal.id))
      await (isGroup
        ? this.processQuery(
            web,
            setGroupOwnerToGroupRequest(groupId, principal.id)
          )
        : web.siteGroups.getById(groupId).setUserAsOwner(principal.id))
      const actual = await this.getOwnerId(web, groupId)
      if (actual !== principal.id) {
        super.log_error(
          'processSiteGroupOwner',
          `Owner of site group ${title} is ${actual} after setting it to ${owner} (${principal.id}): SharePoint did not apply the change.`
        )
        return
      }
      super.log_info(
        'processSiteGroupOwner',
        `Owner of site group ${title} set to ${owner} (${principal.id}).`
      )
    } catch (error) {
      super.log_error(
        'processSiteGroupOwner',
        `Failed to set owner ${owner} on site group ${title}: ${
          (error && error.message) || error
        }`
      )
    }
  }

  /**
   * Whether a principal id is a site group of the site collection (else it is taken as a user).
   *
   * @param web - The web
   * @param id - The principal id
   */
  private async isSiteGroup(web: IWeb, id: number): Promise<boolean> {
    try {
      await web.siteGroups.getById(id).select('Id')()
      return true
    } catch {
      return false
    }
  }

  /**
   * The id of a site group's owner.
   *
   * @param web - The web
   * @param groupId - The group id
   */
  private async getOwnerId(
    web: IWeb,
    groupId: number
  ): Promise<number | undefined> {
    const group = await web.siteGroups
      .getById(groupId)
      .expand('Owner')
      .select('Owner/Id')<{ Owner?: { Id: number } }>()
    return group.Owner && group.Owner.Id
  }

  /**
   * Applies role inheritance and role assignments on the web. Inheritance is
   * only broken when the web still inherits, so re-applying a template with
   * `ClearSubscopes` does not clear unique permissions set since.
   *
   * @param web - The web
   * @param security - The security configuration
   */
  private async processWebSecurity(
    web: IWeb,
    security: IWebSecurity
  ): Promise<void> {
    if (security.BreakRoleInheritance) {
      try {
        const { HasUniqueRoleAssignments } = await web.select(
          'HasUniqueRoleAssignments'
        )<{ HasUniqueRoleAssignments: boolean }>()
        if (HasUniqueRoleAssignments) {
          super.log_info(
            'processWebSecurity',
            'The web already has unique permissions, role inheritance is left as is.'
          )
        } else {
          await web.breakRoleInheritance(
            !!security.CopyRoleAssignments,
            !!security.ClearSubscopes
          )
          super.log_info(
            'processWebSecurity',
            `Role inheritance broken for the web (copyRoleAssignments=${!!security.CopyRoleAssignments}, clearSubscopes=${!!security.ClearSubscopes}).`
          )
        }
      } catch (error) {
        super.log_error(
          'processWebSecurity',
          `Failed to break role inheritance for the web: ${
            (error && error.message) || error
          }`
        )
      }
    }
    for (const assignment of security.RoleAssignments || []) {
      await this.processWebRoleAssignment(web, assignment)
    }
  }

  /**
   * Adds a role assignment on the web.
   *
   * @param web - The web
   * @param assignment - The role assignment
   */
  private async processWebRoleAssignment(
    web: IWeb,
    assignment: IRoleAssignment
  ): Promise<void> {
    const principal =
      typeof assignment.Principal === 'string'
        ? this.tokenHelper.replaceTokens(assignment.Principal)
        : assignment.Principal
    try {
      const principalId = await resolvePrincipalId(web, principal, (message) =>
        this.log_info('resolvePrincipalId', message)
      )
      const roleDefinitionId = await resolveRoleDefinitionId(
        web,
        assignment.RoleDefinition
      )
      await web.roleAssignments.add(principalId, roleDefinitionId)
      super.log_info(
        'processWebSecurity',
        `Role assignment added on the web: ${principal} (${principalId}) -> ${assignment.RoleDefinition} (${roleDefinitionId}).`
      )
    } catch (error) {
      super.log_error(
        'processWebSecurity',
        `Failed to add role assignment ${principal} -> ${
          assignment.RoleDefinition
        } on the web: ${(error && error.message) || error}`
      )
    }
  }
}
