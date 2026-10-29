import { extractWebUrl, IWeb, SPInstance, spPost } from '@pnp/sp/presets/all'

/**
 * SharePoint's client object model (CSOM) over HTTP: an XML request posted to
 * `_vti_bin/client.svc/ProcessQuery`, for what the REST API cannot do. It goes through the web's
 * own PnPjs pipeline, so it works signed in however the web is (SPFx in the browser, a bearer token
 * or cookies in Node), without JSOM.
 */

/** `SP.RequestContext`, whose static `Current` property is the root of every object path. */
const REQUEST_CONTEXT = '{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}'

/**
 * Checks that a principal id is a positive integer, since it is written into the request XML.
 *
 * @param id - The id
 * @param name - What it is, for the error
 */
function principalId(id: number, name: string): number {
  if (!Number.isInteger(id) || id <= 0)
    throw new Error(`${name} must be a positive integer, got ${id}`)
  return id
}

/**
 * The request that makes a site group the owner of another site group: `Group.Owner = group`, then
 * `Group.Update()`. REST's `SetUserAsOwner` answers success for a group id and changes nothing.
 *
 * @param groupId - The group to change
 * @param ownerGroupId - The group that becomes its owner (may be the group itself)
 */
export function setGroupOwnerToGroupRequest(
  groupId: number,
  ownerGroupId: number
): string {
  const group = principalId(groupId, 'groupId')
  const owner = principalId(ownerGroupId, 'ownerGroupId')
  return [
    '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0"',
    ' ApplicationName="sp-js-provisioning" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">',
    '<Actions>',
    '<SetProperty Id="1" ObjectPathId="5" Name="Owner"><Parameter ObjectPathId="6" /></SetProperty>',
    '<Method Name="Update" Id="2" ObjectPathId="5" />',
    '</Actions>',
    '<ObjectPaths>',
    `<StaticProperty Id="0" TypeId="${REQUEST_CONTEXT}" Name="Current" />`,
    '<Property Id="3" ParentId="0" Name="Web" />',
    '<Property Id="4" ParentId="3" Name="SiteGroups" />',
    `<Method Id="5" ParentId="4" Name="GetById"><Parameters><Parameter Type="Int32">${group}</Parameter></Parameters></Method>`,
    `<Method Id="6" ParentId="4" Name="GetById"><Parameters><Parameter Type="Int32">${owner}</Parameter></Parameters></Method>`,
    '</ObjectPaths>',
    '</Request>'
  ].join('')
}

/**
 * Posts a CSOM request to the web's `_vti_bin/client.svc/ProcessQuery` and throws when SharePoint
 * reports an error in the response (CSOM answers 200 with the error in the body).
 *
 * @param web - The web
 * @param request - The request XML
 */
export async function processQuery(web: IWeb, request: string): Promise<any[]> {
  const webUrl = extractWebUrl(web.toUrl())
  const response = await spPost<any[]>(
    SPInstance([web, webUrl], '_vti_bin/client.svc/ProcessQuery'),
    { body: request, headers: { 'Content-Type': 'text/xml' } }
  )
  const errorInfo =
    Array.isArray(response) && response[0] && response[0].ErrorInfo
  if (errorInfo)
    throw new Error(
      `CSOM request failed: ${errorInfo.ErrorMessage} (${errorInfo.ErrorTypeName})`
    )
  return response
}
