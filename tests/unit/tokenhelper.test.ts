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
