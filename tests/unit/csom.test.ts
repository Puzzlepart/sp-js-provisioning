import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
// The real `@pnp/sp`: the request is built by PnPjs's own pipeline, which a
// hand-rolled fake would not prove.
import { Web } from '@pnp/sp/presets/all'
import { RejectOnError, ResolveOnData } from '@pnp/queryable'
import {
  CsomError,
  CsomRequest,
  executeQuery,
  guid,
  guidOf,
  isNull,
  processQuery,
  reference,
  setGroupOwnerToGroupRequest,
  valueObject
} from '../../src/util/csom'
import {
  CsomFailure,
  CSOM_PATH,
  describeCsom,
  FakeCsom,
  parseCsom
} from '../helpers/csom'
import { createWireWeb } from '../helpers/transport'

const CT_INFO = '{168f3091-4554-4f14-8866-b20d48e45b54}'
const TAXONOMY_SESSION = '{981cbc68-9edc-4f8d-872f-71146fcbb84f}'

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

describe('CSOM: the request builder', () => {
  it('writes object paths and actions with ids that are unique in the request', () => {
    const request = new CsomRequest()
    const contentTypes = request.property(request.web(), 'ContentTypes')
    const contentType = request.method(contentTypes, 'GetById', '0x0100AB')
    const found = request.instantiate(contentType)
    const read = request.query(contentType, ['StringId', 'Name'])
    const set = request.set(contentType, 'Name', 'Renamed')
    const update = request.call(contentType, 'Update', true)

    assert.equal(
      request.toString(),
      [
        '<Request AddExpandoFieldTypeSuffix="true" SchemaVersion="15.0.0.0" LibraryVersion="16.0.0.0"',
        ' ApplicationName="sp-js-provisioning" xmlns="http://schemas.microsoft.com/sharepoint/clientquery/2009">',
        '<Actions>',
        '<ObjectPath Id="4" ObjectPathId="3" />',
        '<Query Id="5" ObjectPathId="3"><Query SelectAllProperties="false"><Properties>',
        '<Property Name="StringId" ScalarProperty="true" /><Property Name="Name" ScalarProperty="true" />',
        '</Properties></Query></Query>',
        '<SetProperty Id="6" ObjectPathId="3" Name="Name"><Parameter Type="String">Renamed</Parameter></SetProperty>',
        '<Method Name="Update" Id="7" ObjectPathId="3"><Parameters><Parameter Type="Boolean">true</Parameter></Parameters></Method>',
        '</Actions>',
        '<ObjectPaths>',
        '<StaticProperty Id="0" TypeId="{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}" Name="Current" />',
        '<Property Id="1" ParentId="0" Name="Web" />',
        '<Property Id="2" ParentId="1" Name="ContentTypes" />',
        '<Method Id="3" ParentId="2" Name="GetById"><Parameters><Parameter Type="String">0x0100AB</Parameter></Parameters></Method>',
        '</ObjectPaths>',
        '</Request>'
      ].join('')
    )
    assert.deepEqual([found, read, set, update], [4, 5, 6, 7])
    assert.equal(request.actionCount, 4)
  })

  it('names the context and the web once, however often they are asked for', () => {
    const request = new CsomRequest()
    assert.equal(request.web(), request.web())
    assert.equal(request.current(), request.current())
    request.instantiate(request.property(request.web(), 'Fields'))

    assert.equal((request.toString().match(/Name="Current"/g) || []).length, 1)
    assert.equal((request.toString().match(/Name="Web"/g) || []).length, 1)
    assert.deepEqual(describeCsom(request.toString()), [
      'instantiate Web.Fields'
    ])
  })

  it('writes every kind of parameter', () => {
    const request = new CsomRequest()
    const web = request.web()
    const field = request.method(
      request.property(web, 'AvailableFields'),
      'GetById',
      guid('{1B2C3D4E-0000-4000-8000-00000000000A}')
    )
    request.call(
      web,
      'Everything',
      'text',
      true,
      false,
      42,
      -7,
      null,
      undefined,
      ['a', 'b'],
      guid('1b2c3d4e-0000-4000-8000-00000000000b'),
      reference(field),
      valueObject(CT_INFO, {
        Name: 'A content type',
        Id: '0x0100AB',
        ParentContentType: null,
        Sealed: false,
        Field: reference(field)
      })
    )

    const xml = request.toString()
    for (const part of [
      '<Parameter Type="String">text</Parameter>',
      '<Parameter Type="Boolean">true</Parameter>',
      '<Parameter Type="Boolean">false</Parameter>',
      '<Parameter Type="Int32">42</Parameter>',
      '<Parameter Type="Int32">-7</Parameter>',
      '<Parameter Type="Null" /><Parameter Type="Null" />',
      '<Parameter Type="Array"><Object Type="String">a</Object><Object Type="String">b</Object></Parameter>',
      '<Parameter Type="Guid">{1b2c3d4e-0000-4000-8000-00000000000b}</Parameter>',
      `<Parameter ObjectPathId="${field}" />`,
      // A value object's properties, in name order.
      `<Parameter TypeId="${CT_INFO}">` +
        `<Property Name="Field" ObjectPathId="${field}" />` +
        '<Property Name="Id" Type="String">0x0100AB</Property>' +
        '<Property Name="Name" Type="String">A content type</Property>' +
        '<Property Name="ParentContentType" Type="Null" />' +
        '<Property Name="Sealed" Type="Boolean">false</Property>' +
        '</Parameter>'
    ])
      assert.ok(xml.includes(part), `the request lacks ${part}`)
    assert.deepEqual(describeCsom(xml), [
      'call Web.Everything("text", true, false, 42, -7, null, null, ["a", "b"], ' +
        '{1b2c3d4e-0000-4000-8000-00000000000b}, ' +
        '@Web.AvailableFields.GetById({1b2c3d4e-0000-4000-8000-00000000000a}), ' +
        'ContentTypeCreationInformation{ ' +
        'Field: @Web.AvailableFields.GetById({1b2c3d4e-0000-4000-8000-00000000000a}), ' +
        'Id: "0x0100AB", Name: "A content type", ParentContentType: null, Sealed: false })'
    ])
  })

  it('writes a static method, a method without parameters and a query of the items of a collection', () => {
    const request = new CsomRequest()
    const store = request.method(
      request.staticMethod(TAXONOMY_SESSION, 'GetTaxonomySession'),
      'GetDefaultSiteCollectionTermStore'
    )
    const links = request.property(
      request.method(
        request.property(request.web(), 'ContentTypes'),
        'GetById',
        '0x01'
      ),
      'FieldLinks'
    )
    request.query(store, ['DefaultLanguage'])
    request.query(links, [], ['Id', 'Name'])
    request.call(store, 'CommitAll')

    const xml = request.toString()
    assert.ok(
      xml.includes(
        `<StaticMethod Id="0" Name="GetTaxonomySession" TypeId="${TAXONOMY_SESSION}" />` +
          '<Method Id="1" ParentId="0" Name="GetDefaultSiteCollectionTermStore" />'
      ),
      xml
    )
    assert.ok(
      xml.includes(
        '<Query SelectAllProperties="false"><Properties></Properties></Query>' +
          '<ChildItemQuery SelectAllProperties="false"><Properties>' +
          '<Property Name="Id" ScalarProperty="true" /><Property Name="Name" ScalarProperty="true" />' +
          '</Properties></ChildItemQuery>'
      ),
      xml
    )
    assert.ok(xml.includes('<Method Name="CommitAll" Id="'), xml)
    assert.deepEqual(describeCsom(xml), [
      'query TaxonomySession.GetTaxonomySession().GetDefaultSiteCollectionTermStore() [DefaultLanguage]',
      'query Web.ContentTypes.GetById("0x01").FieldLinks [] items [Id, Name]',
      'call TaxonomySession.GetTaxonomySession().GetDefaultSiteCollectionTermStore().CommitAll()'
    ])
  })

  it('escapes text, and what it holds comes back unchanged when the XML is read', () => {
    const text =
      'Risiko & <mulighet> "sitat" \'apostrof\' æøå — 你好\r\nlinje 2\tkolonne'
    const request = new CsomRequest()
    request.set(request.web(), 'Description', text)
    request.call(request.web(), 'Tag', valueObject(CT_INFO, { 'A&B': text }), [
      text
    ])

    const xml = request.toString()
    assert.ok(
      xml.includes(
        "Risiko &amp; &lt;mulighet&gt; &quot;sitat&quot; 'apostrof' æøå — 你好&#xD;\nlinje 2\tkolonne"
      ),
      xml
    )
    assert.ok(xml.includes('<Property Name="A&amp;B" Type="String">'), xml)
    const [set, call] = parseCsom(xml)
    assert.deepEqual(set.args, [text])
    assert.deepEqual(call.args, [
      { type: 'ContentTypeCreationInformation', properties: { 'A&B': text } },
      [text]
    ])
  })

  it('leaves out characters XML cannot hold', () => {
    const request = new CsomRequest()
    request.set(request.web(), 'Description', 'a\u0000b\u000Bc\u001Fd￾e')

    assert.ok(request.toString().includes('<Parameter Type="String">abcde<'))
  })

  it('refuses what cannot be written into a request', () => {
    const request = new CsomRequest()
    const web = request.web()
    assert.throws(() => guid('not-a-guid'), /must be a GUID, got not-a-guid/)
    assert.throws(
      () => guid('1b2c3d4e-0000-4000-8000-00000000000b"/><x'),
      /must be a GUID/
    )
    assert.throws(() => valueObject('"><x', {}), /A type id must be a GUID/)
    assert.throws(
      () => request.staticMethod('nope', 'Get'),
      /A type id must be a GUID/
    )
    assert.throws(
      () => request.call(web, 'Add', 1.5),
      /32-bit integer, got 1.5/
    )
    assert.throws(
      () => request.call(web, 'Add', Number.NaN),
      /32-bit integer, got NaN/
    )
    assert.throws(
      () => request.call(web, 'Add', 2147483648),
      /32-bit integer, got 2147483648/
    )
    assert.throws(
      () => request.call(web, 'Add', ({ some: 'object' } as unknown) as string),
      /Not a CSOM value: {"some":"object"}/
    )
    assert.throws(
      () => request.property(('1" /><x' as unknown) as number, 'Web'),
      /Not an object path of this request/
    )
    assert.throws(
      () => request.call(web, 'Add', reference(-1)),
      /Not an object path of this request: -1/
    )
  })

  it('reads a GUID out of a CSOM answer, a REST answer and a template', () => {
    const bare = '1b2c3d4e-0000-4000-8000-00000000000b'
    assert.equal(guidOf(`/Guid(${bare})/`), bare)
    assert.equal(guidOf(`{${bare.toUpperCase()}}`), bare)
    assert.equal(guidOf(bare), bare)
    assert.equal(guidOf('no guid'), undefined)
    assert.equal(guidOf(undefined), undefined)
  })

  it('tells an object that is not there from one that is', () => {
    assert.equal(isNull({ IsNull: true }), true)
    assert.equal(isNull({ IsNull: false }), false)
    assert.equal(isNull(undefined), true)
    assert.equal(isNull(null), true)
  })
})

describe('CSOM: the group owner request', () => {
  it('sets a group as the owner of another group and updates it', () => {
    const request = setGroupOwnerToGroupRequest(12, 3)

    assert.deepEqual(describeCsom(request), [
      'set Web.SiteGroups.GetById(12).Owner = @Web.SiteGroups.GetById(3)',
      'call Web.SiteGroups.GetById(12).Update()'
    ])
    assert.match(
      request,
      /<Method Id="\d+" ParentId="\d+" Name="GetById"><Parameters><Parameter Type="Int32">12<\/Parameter>/
    )
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
            ErrorCode: -2147024891,
            ErrorTypeName: 'System.UnauthorizedAccessException'
          }
        }
      ]
    )

    await assert.rejects(
      () => processQuery(web, '<Request />'),
      (error: any) => {
        assert.ok(error instanceof CsomError)
        assert.ok(error instanceof Error)
        assert.equal(
          error.message,
          'CSOM request failed: Access denied. (System.UnauthorizedAccessException)'
        )
        assert.equal(error.name, 'CsomError')
        assert.equal(error.serverMessage, 'Access denied.')
        assert.equal(error.typeName, 'System.UnauthorizedAccessException')
        assert.equal(error.code, -2147024891)
        assert.equal(error.accessDenied, true)
        return true
      }
    )
  })
})

describe('CSOM: executeQuery', () => {
  it('returns the results by action id', async () => {
    const csom = new FakeCsom((action) => {
      if (action.kind === 'instantiate')
        return { IsNull: action.path.includes('0x0100FF') }
      if (action.kind === 'query')
        return { _ObjectType_: 'SP.ContentType', StringId: '0x0100AB' }
    })
    const { web, requests } = createWireWeb((request) =>
      csom.respond(request.body)
    )
    const request = new CsomRequest()
    const contentTypes = request.property(request.web(), 'ContentTypes')
    const there = request.method(contentTypes, 'GetById', '0x0100AB')
    const missing = request.method(contentTypes, 'GetById', '0x0100FF')
    const found = request.instantiate(there)
    const read = request.query(there, ['StringId'])
    const notFound = request.instantiate(missing)
    const updated = request.call(there, 'Update', false)

    const results = await executeQuery(web, request)

    assert.deepEqual(
      requests.map((sent) => [sent.method, sent.path]),
      [['POST', CSOM_PATH]]
    )
    assert.equal(isNull(results[found]), false)
    assert.equal(isNull(results[notFound]), true)
    assert.equal(results[read].StringId, '0x0100AB')
    assert.equal(results[updated], undefined)
    assert.deepEqual(csom.ran, [
      'instantiate Web.ContentTypes.GetById("0x0100AB")',
      'query Web.ContentTypes.GetById("0x0100AB") [StringId]',
      'instantiate Web.ContentTypes.GetById("0x0100FF")',
      'call Web.ContentTypes.GetById("0x0100AB").Update(false)'
    ])
    assert.deepEqual(csom.writes, [
      'call Web.ContentTypes.GetById("0x0100AB").Update(false)'
    ])
  })

  it('throws what SharePoint reports, and the actions after the failing one do not run', async () => {
    const csom = new FakeCsom((action) => {
      if (action.kind === 'set')
        throw new CsomFailure(
          'Det finnes allerede en innholdstype med dette navnet.',
          'Microsoft.SharePoint.SPDuplicateValuesFoundException',
          -2130575342
        )
    })
    const { web } = createWireWeb((request) => csom.respond(request.body))
    const request = new CsomRequest()
    const contentType = request.method(
      request.property(request.web(), 'ContentTypes'),
      'GetById',
      '0x0100AB'
    )
    request.instantiate(contentType)
    request.set(contentType, 'Name', 'Taken')
    request.call(contentType, 'Update', true)

    await assert.rejects(
      () => executeQuery(web, request),
      (error: any) => {
        assert.ok(error instanceof CsomError)
        assert.equal(
          error.serverMessage,
          'Det finnes allerede en innholdstype med dette navnet.'
        )
        assert.equal(error.accessDenied, false)
        return true
      }
    )
    assert.deepEqual(csom.ran, [
      'instantiate Web.ContentTypes.GetById("0x0100AB")',
      'FAILED set Web.ContentTypes.GetById("0x0100AB").Name = "Taken"'
    ])
  })
})
