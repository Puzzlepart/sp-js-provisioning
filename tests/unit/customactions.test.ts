import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { CustomActions } from '../../src/handlers/customactions'

/**
 * A `CustomActions` handler whose batching seam is replaced with a hand-rolled
 * `[batch, execute]` pair, so the batch lifecycle can be driven from a test
 * without a real PnPjs queryable.
 */
class FakeBatchCustomActions extends CustomActions {
  /** Titles passed to `userCustomActions.add` through the batch, in order */
  public added: string[] = []

  /**
   * Creates a handler with a fake batch.
   *
   * @param executeResult - What the fake `execute()` should resolve/reject with
   * @param requestResult - Factory for the promise each batched `add` returns
   */
  constructor(
    private executeResult: () => Promise<void>,
    private requestResult: (title: string) => Promise<any>
  ) {
    super({})
  }

  protected createBatch(): [(instance: any) => any, () => Promise<void>] {
    const batch = (instance: any) => instance
    return [batch, () => this.executeResult()]
  }

  /**
   * Builds the fake web the handler provisions against. `userCustomActions` is
   * invokable (to read the existing actions) and `using` returns a collection
   * whose `add` resolves with the configured per-request promise.
   */
  public createWeb(): any {
    const userCustomActions: any = async () => []
    userCustomActions.select = () => userCustomActions
    userCustomActions.using = () => ({
      add: (action: any) => {
        this.added.push(action.Title)
        return this.requestResult(action.Title)
      }
    })
    return { userCustomActions }
  }
}

/**
 * Runs `run` with every `unhandledRejection` listener detached (node:test
 * installs one and would attribute a stray rejection to an unrelated test) and
 * returns the rejections Node reported as unobserved.
 *
 * @param run - The work to observe
 */
async function collectUnhandledRejections(
  run: () => Promise<void>
): Promise<any[]> {
  const unhandled: any[] = []
  const existing = process.listeners('unhandledRejection')
  for (const listener of existing) {
    process.removeListener('unhandledRejection', listener as any)
  }
  const capture = (reason: any) => unhandled.push(reason)
  process.on('unhandledRejection', capture)
  try {
    await run()
    // Node reports unobserved rejections on the next turn of the event loop,
    // so the check has to outlive the awaited work.
    await new Promise((resolve) => setImmediate(resolve))
    await new Promise((resolve) => setImmediate(resolve))
  } finally {
    process.removeListener('unhandledRejection', capture)
    for (const listener of existing) {
      process.on('unhandledRejection', listener as any)
    }
  }
  return unhandled
}

describe('CustomActions: v4 batching', () => {
  it('leaves no batched request unobserved when execute rejects', async () => {
    // `execute()` and the per-request promises have to be awaited together:
    // awaiting `execute()` first abandons the N request promises, which then
    // reject unobserved and take the process down under Node's default
    // unhandled rejection policy.
    let thrown: Error
    const handler = new FakeBatchCustomActions(
      async () => {
        throw new Error('batch failed')
      },
      async (title) => {
        throw new Error(`request failed: ${title}`)
      }
    )

    const unhandled = await collectUnhandledRejections(async () => {
      await assert
        .rejects(
          handler.ProvisionObjects(handler.createWeb(), [
            { Title: 'CustomActionOne' },
            { Title: 'CustomActionTwo' }
          ] as any),
          /batch failed|request failed/
        )
        .then(
          () => undefined,
          (error) => {
            thrown = error
          }
        )
    })

    assert.deepEqual(handler.added, ['CustomActionOne', 'CustomActionTwo'])
    assert.equal(thrown, undefined)
    assert.deepEqual(
      unhandled.map((error) => error.message),
      []
    )
  })

  it('throws when a batched custom action fails but the batch itself succeeds', async () => {
    const handler = new FakeBatchCustomActions(
      async () => undefined,
      async (title) => {
        if (title === 'CustomActionTwo') {
          throw new Error('request failed')
        }
        return { Title: title }
      }
    )

    const unhandled = await collectUnhandledRejections(async () => {
      await assert.rejects(
        handler.ProvisionObjects(handler.createWeb(), [
          { Title: 'CustomActionOne' },
          { Title: 'CustomActionTwo' }
        ] as any),
        /request failed/
      )
    })

    assert.deepEqual(
      unhandled.map((error) => error.message),
      []
    )
  })
})
