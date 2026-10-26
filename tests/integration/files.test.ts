import { describe, it } from 'node:test'

/**
 * Files runs through JSOM (spfx-jsom: web parts through LimitedWebPartManager, and it refuses whenever an SPFx context is set), which cannot run in Node. Its suite
 * lands with the JSOM exit (docs/specs/integration-testing.md, D3; docs/specs/jsom-exit.md).
 */
describe(
  'Files',
  { skip: 'JSOM handler: browser-only until the JSOM exit (D3)' },
  () => {
    it('applies, reads back and changes nothing on a second apply', () =>
      undefined)
  }
)
