import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { ProvisioningContext } from '../../src/provisioningcontext'
import { TokenHelper } from '../../src/util/tokenhelper'

describe('TokenHelper: {sitename}', () => {
  it('resolves to the title of the web being provisioned', () => {
    const context = new ProvisioningContext()
    context.web = { Id: 'web-guid', Title: 'Prosjekt A' }

    assert.equal(
      new TokenHelper(context, {}).replaceTokens('{sitename} Styringsgruppe'),
      'Prosjekt A Styringsgruppe'
    )
  })

  it('is left untouched when the web is not loaded', () => {
    const helper = new TokenHelper(new ProvisioningContext(), {})

    assert.equal(helper.replaceTokens('{sitename}'), '{sitename}')
  })
})

describe('TokenHelper: unresolved tokens', () => {
  const helper = () => {
    const context = new ProvisioningContext()
    context.web = { Id: 'web-guid', Title: 'Prosjekt A' }
    context.lists['Kilder'] = 'list-guid'
    return new TokenHelper(context, { parameters: { Fase: 'Konsept' } })
  }

  it('names each known token it cannot resolve, once, as written', () => {
    assert.deepEqual(
      helper().unresolvedTokens(
        '<Field List="{listid:Mangler}" WebId="{webid}" Description="{parameter:Ukjent} {listid:Mangler} {sitecollectiontermstoreid}" />'
      ),
      ['{listid:Mangler}', '{parameter:Ukjent}', '{sitecollectiontermstoreid}']
    )
  })

  it('leaves out tokens it resolves, and text that only looks like a token', () => {
    // `{Title}` (upper case) and JSON are not tokens; `{{listid:X}}` holds one.
    assert.deepEqual(
      helper().unresolvedTokens(
        '{listid:Kilder} {parameter:Fase} {sitename} {Title} {"a":1} {ukjent} {{listid:Kilder}}'
      ),
      []
    )
    assert.deepEqual(helper().unresolvedTokens('{{listid:Borte}}'), [
      '{listid:Borte}'
    ])
  })

  it('finds what replaceTokens left in place', () => {
    const tokens = helper()
    const replaced = tokens.replaceTokens(
      'List="{listid:Kilder}" X="{listid:Borte}"'
    )

    assert.equal(replaced, 'List="list-guid" X="{listid:Borte}"')
    assert.deepEqual(tokens.unresolvedTokens(replaced), ['{listid:Borte}'])
    assert.deepEqual(tokens.unresolvedTokens(undefined), [])
  })
})
