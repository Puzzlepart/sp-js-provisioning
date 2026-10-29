import { describe, it } from 'node:test'

/**
 * ContentTypes runs through JSOM (spfx-jsom: content types with explicit ids and field links through SP.ClientContext), which cannot run in Node. Its suite
 * lands with the JSOM exit (docs/specs/integration-testing.md, D3; docs/specs/jsom-exit.md).
 */
describe(
  'ContentTypes',
  { skip: 'JSOM handler: browser-only until the JSOM exit (D3)' },
  () => {
    it('applies, reads back and changes nothing on a second apply', () =>
      undefined)
  }
)
