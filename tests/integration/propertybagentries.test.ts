import { describe, it } from 'node:test'

/**
 * PropertyBagEntries runs through JSOM (spfx-jsom: AllProperties through SP.ClientContext; it also needs custom script (NoScript blocks it)), which cannot run in Node. Its suite
 * lands with the JSOM exit (docs/specs/integration-testing.md, D3; docs/specs/jsom-exit.md).
 */
describe(
  'PropertyBagEntries',
  { skip: 'JSOM handler: browser-only until the JSOM exit (D3)' },
  () => {
    it('applies, reads back and changes nothing on a second apply', () =>
      undefined)
  }
)
