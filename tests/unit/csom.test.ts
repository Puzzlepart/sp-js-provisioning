import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
// The real `@pnp/sp`: the request is built by PnPjs's own pipeline, which a
// hand-rolled fake would not prove.
import { Web } from '@pnp/sp/presets/all'
import { RejectOnError, ResolveOnData } from '@pnp/queryable'
import { processQuery, setGroupOwnerToGroupRequest } from '../../src/util/csom'

/**
 * A web whose requests are recorded and answered with `response` in the `pre`
 * moment, so no network call is made.
 *
 * @param seen - Array each request is pushed to
 * @param response - The CSOM response body
 */
function recordingWeb(seen: any[], response: any): any {
  const recording = (instance: any) => {
    instance.on.pre(async (url: string, init: any) => {
      seen.push({ url, init })
      return [url, init, response]
    })
    return instance
  }
  return Web('https://contoso.sharepoint.com/sites/hub/sub').using(
    ResolveOnData(),
    RejectOnError(),
    recording
  )
}

describe('CSOM: the group owner request', () => {
  it('sets a group as the owner of another group and updates it', () => {
    const request = setGroupOwnerToGroupRequest(12, 3)

    assert.match(
      request,
      /<SetProperty Id="1" ObjectPathId="5" Name="Owner"><Parameter ObjectPathId="6" \/><\/SetProperty>/
    )
    assert.match(request, /<Method Name="Update" Id="2" ObjectPathId="5" \/>/)
    assert.match(
      request,
      /<Method Id="5" ParentId="4" Name="GetById"><Parameters><Parameter Type="Int32">12<\/Parameter>/
    )
    assert.match(
      request,
      /<Method Id="6" ParentId="4" Name="GetById"><Parameters><Parameter Type="Int32">3<\/Parameter>/
    )
    assert.match(request, /<Property Id="4" ParentId="3" Name="SiteGroups" \/>/)
  })

  it('refuses ids that are not positive integers', () => {
    // The ids are written into the XML.
    for (const id of [0, -1, 1.5, Number.NaN, ('3"/>' as unknown) as number])
      assert.throws(
        () => setGroupOwnerToGroupRequest(id, 3),
        /groupId must be a positive integer/
      )
    assert.throws(
      () => setGroupOwnerToGroupRequest(3, 0),
      /ownerGroupId must be a positive integer/
    )
  })
})

describe('CSOM: processQuery', () => {
  it("posts the XML to the web's ProcessQuery endpoint", async () => {
    const seen: any[] = []
    const web = recordingWeb(seen, [
      { SchemaVersion: '15.0.0.0', ErrorInfo: null }
    ])

    await processQuery(web, '<Request />')

    assert.equal(seen.length, 1)
    assert.equal(
      seen[0].url,
      'https://contoso.sharepoint.com/sites/hub/sub/_vti_bin/client.svc/ProcessQuery'
    )
    assert.equal(seen[0].init.method, 'POST')
    assert.equal(seen[0].init.body, '<Request />')
    assert.equal(
      new Headers(seen[0].init.headers).get('Content-Type'),
      'text/xml'
    )
  })

  it('throws the error SharePoint reports in a 200 response', async () => {
    const web = recordingWeb(
      [],
      [
        {
          SchemaVersion: '15.0.0.0',
          ErrorInfo: {
            ErrorMessage: 'Access denied.',
            ErrorTypeName: 'System.UnauthorizedAccessException'
          }
        }
      ]
    )

    await assert.rejects(
      () => processQuery(web, '<Request />'),
      /CSOM request failed: Access denied\. \(System\.UnauthorizedAccessException\)/
    )
  })
})
