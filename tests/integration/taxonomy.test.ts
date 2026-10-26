import { describe, it } from 'node:test'

/**
 * Taxonomy runs through JSOM (spfx-jsom: term groups, term sets and terms through SP.Taxonomy), which cannot run in Node. Its suite
 * lands with the JSOM exit (docs/specs/integration-testing.md, D3; docs/specs/jsom-exit.md).
 */
describe(
  'Taxonomy',
  { skip: 'JSOM handler: browser-only until the JSOM exit (D3)' },
  () => {
    it('applies, reads back and changes nothing on a second apply', () =>
      undefined)
  }
)
