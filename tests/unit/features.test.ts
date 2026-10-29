import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { Features } from '../../src/handlers/features'

const followingContent = 'a7a2793e-67cd-4dc1-9fd0-43f61581207a'

/**
 * A fake web whose features are active as given, recording each add and remove.
 *
 * @param active - The active feature ids
 * @param calls - Where each call is recorded
 */
function webWith(active: string[], calls: string[]): any {
  const collection: any = async () => active.map((id) => ({ DefinitionId: id }))
  collection.select = () => collection
  collection.add = async (id: string, force: boolean) => {
    calls.push(`add ${id} ${force}`)
  }
  collection.remove = async (id: string, force: boolean) => {
    calls.push(`remove ${id} ${force}`)
  }
  return { features: collection }
}

describe('Features', () => {
  it('leaves a feature already in the requested state alone', async () => {
    // SharePoint answers 500 to activating an active feature and 400 to
    // deactivating an inactive one: re-applying a template stopped here.
    const calls: string[] = []
    await new Features({}).ProvisionObjects(
      webWith([followingContent.toUpperCase()], calls),
      [
        { id: `{${followingContent}}`, deactivate: false, force: false },
        {
          id: 'b0b2a0a5-0000-0000-0000-000000000001',
          deactivate: true,
          force: false
        }
      ]
    )

    assert.deepEqual(calls, [])
  })

  it('activates an inactive feature and deactivates an active one', async () => {
    const calls: string[] = []
    await new Features({}).ProvisionObjects(
      webWith([followingContent], calls),
      [
        {
          id: 'b0b2a0a5-0000-0000-0000-000000000001',
          deactivate: false,
          force: false
        },
        { id: followingContent, deactivate: true, force: true }
      ]
    )

    assert.deepEqual(calls, [
      'add b0b2a0a5-0000-0000-0000-000000000001 false',
      `remove ${followingContent} true`
    ])
  })

  it('re-activates an active feature with force', async () => {
    const calls: string[] = []
    await new Features({}).ProvisionObjects(
      webWith([followingContent], calls),
      [{ id: followingContent, deactivate: false, force: true }]
    )

    assert.deepEqual(calls, [`add ${followingContent} true`])
  })
})

describe('Features: the same feature twice in a template', () => {
  /**
   * A fake web that keeps the features' state and answers like SharePoint: 500 to activating an
   * active feature, 400 to deactivating an inactive one.
   *
   * @param active - The active feature ids
   * @param calls - Where each call is recorded
   */
  function statefulWeb(active: Set<string>, calls: string[]): any {
    const collection: any = async () =>
      [...active].map((id) => ({ DefinitionId: id }))
    collection.select = () => collection
    collection.add = async (id: string, force: boolean) => {
      calls.push(`add ${id}`)
      if (active.has(id) && !force)
        throw Object.assign(new Error('already activated'), { status: 500 })
      active.add(id)
    }
    collection.remove = async (id: string) => {
      calls.push(`remove ${id}`)
      if (!active.has(id))
        throw Object.assign(new Error('not activated'), { status: 400 })
      active.delete(id)
    }
    return { features: collection }
  }

  it('follows each entry in order: deactivate then activate ends active', async () => {
    const active = new Set([followingContent])
    const calls: string[] = []
    await new Features({}).ProvisionObjects(statefulWeb(active, calls), [
      { id: followingContent, deactivate: true, force: false },
      { id: followingContent, deactivate: false, force: false }
    ])

    assert.deepEqual(calls, [
      `remove ${followingContent}`,
      `add ${followingContent}`
    ])
    assert.equal(active.has(followingContent), true)
  })

  it('activates a feature listed twice once', async () => {
    const active = new Set<string>()
    const calls: string[] = []
    await new Features({}).ProvisionObjects(statefulWeb(active, calls), [
      { id: followingContent, deactivate: false, force: false },
      { id: followingContent, deactivate: false, force: false }
    ])

    assert.deepEqual(calls, [`add ${followingContent}`])
  })
})
