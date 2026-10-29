import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  hookCall,
  IHookRequest,
  readHooks
} from '../../debug/lib/readback/hooks'
import { loadFixture } from '../fixtures/fixtures'
import { applyTwice, skipReason, withScratchWeb } from './helpers'
import { withHookServer } from './hookserver'

/** The handler's wait before it polls a 202's `Location` (`Hooks.pollInterval`). */
const POLL_INTERVAL = 5000

/**
 * The Hooks handler against a local listener. Hooks run in the calling process, so a listener on
 * 127.0.0.1 is reachable; the scratch web is the run's web, whose property bag the handler merges
 * into each POST body. Hooks are called again on every run, by design.
 */
describe('Hooks', { skip: skipReason() }, () => {
  const scratch = withScratchWeb('hooks')
  const server = withHookServer()

  it(
    'calls GET and POST hooks with their headers and bodies, polls a 202 until 200, and calls them again on a second apply',
    { timeout: 120_000 },
    async (t) => {
      const { conn } = scratch()
      const hooks = server()
      const runs: IHookRequest[][] = []
      const fixture = () => loadFixture('hooks', { hookUrl: hooks.url })
      const { first, second, afterFirst } = await applyTwice(
        conn,
        fixture,
        (c) => {
          const run = hooks.take()
          runs.push(run)
          return readHooks(c, run)
        }
      )

      // Each run finishes (D4: it hung when the 202 poll answered 200).
      assert.deepEqual(first.order, ['Hooks'])
      assert.deepEqual(second.order, ['Hooks'])

      const declared = fixture().Hooks
      const [, post, slow] = declared
      const { properties } = afterFirst
      assert.ok(
        Object.keys(properties).length > 0,
        'the web has a property bag to merge'
      )
      t.diagnostic(
        `each POST body carried ${
          Object.keys(properties).length
        } property bag entries`
      )
      // One call per GET and POST hook and one poll of the slow hook's Location; the PUT hook is
      // not supported and not called.
      assert.deepEqual(afterFirst.calls, [
        {
          method: 'GET',
          url: '/get?source=spjs-itest',
          headers: { accept: 'application/json', 'x-spjs-itest': 'get' }
        },
        {
          method: 'POST',
          url: '/post',
          headers: {
            // `fetch` adds `Accept: */*` when the hook sets none.
            accept: '*/*',
            'content-type': 'application/json',
            'x-spjs-itest': 'post',
            'x-spjs-itest-run': 'hooks'
          },
          // The handler merges the property bag over the template's Body.
          body: { ...post.Body, ...properties }
        },
        {
          method: 'POST',
          url: '/slow',
          headers: {
            accept: '*/*',
            'content-type': 'application/json',
            'x-spjs-itest': 'slow'
          },
          body: { ...slow.Body, ...properties }
        },
        {
          // The poll is a GET with the hook's headers and no body.
          method: 'GET',
          url: '/slow/status',
          headers: {
            accept: '*/*',
            'content-type': 'application/json',
            'x-spjs-itest': 'slow'
          }
        }
      ])
      assert.ok(
        !hooks.requests.some((request) => request.url === '/put'),
        'the PUT hook was called'
      )
      // The template's Body values arrive as declared, not overwritten by a property.
      const posted = hookCall(
        runs[0].find((request) => request.url === '/post')
      ).body
      for (const key of Object.keys(post.Body))
        assert.deepEqual(posted[key], post.Body[key], key)

      // The poll waits before it asks.
      for (const run of runs) {
        const accepted = run.find((request) => request.url === '/slow')
        const poll = run.find((request) => request.url === '/slow/status')
        assert.ok(
          poll.receivedAt - accepted.receivedAt >= POLL_INTERVAL - 100,
          `polled after ${poll.receivedAt - accepted.receivedAt} ms`
        )
      }

      // Called again by design: the second run sent the same calls, so the count doubled.
      assert.equal(runs[0].length, 4)
      assert.equal(runs[1].length, runs[0].length)
      assert.equal(hooks.requests.length, 2 * runs[0].length)
    }
  )
})
