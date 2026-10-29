import assert from 'node:assert/strict'
import { describe, it } from 'node:test'

// The CI scan in .tasks/ is plain JavaScript run by node.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const { findingsIn, hasTenantUrl } = require('../../.tasks/check-secrets.js')

describe('check-secrets', () => {
  it('allows the example hosts, with or without -admin and -my', () => {
    for (const host of [
      'contoso',
      'fabrikam',
      'contoso-admin',
      'contoso-my',
      'Fabrikam-My',
      '<tenant>',
      '<tenant>-admin'
    ])
      assert.equal(
        hasTenantUrl(`https://${host}.sharepoint.com/sites/x`),
        false,
        host
      )
  })

  it('allows a host built at run time', () => {
    // A test writes `https://${tenant}.sharepoint.com`: the source names no host.
    assert.equal(
      hasTenantUrl('`https://${tenant}-admin.sharepoint.com`'),
      false
    )
  })

  it('flags any other SharePoint host', () => {
    for (const host of [
      'puzzlepart',
      'puzzlepart-admin',
      'puzzlepart-my',
      'contosoltd'
    ])
      assert.equal(
        hasTenantUrl(`https://${host}.sharepoint.com/sites/x`),
        true,
        host
      )
  })

  it('flags long cookie values, bearer tokens and certificates, not test stand-ins', () => {
    // Built at run time, so this file does not trip the scan itself.
    const long = 'A'.repeat(60)
    assert.deepEqual(findingsIn(`FedAuth=session; rtFa=b`), [])
    assert.deepEqual(findingsIn(`Cookie: FedAuth=${long}`), [
      '1: a session cookie'
    ])
    assert.deepEqual(
      findingsIn(`Authorization: Bearer ${'ey' + 'J'.repeat(20)}`),
      ['1: a bearer token']
    )
    assert.deepEqual(findingsIn(`\npfx: MII${long}`), ['2: a certificate'])
  })
})
