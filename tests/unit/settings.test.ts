import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { differingSettings } from '../../src/util/settings'

describe('differingSettings', () => {
  it('keeps only the settings whose value differs from what is stored', () => {
    assert.deepEqual(
      differingSettings(
        {
          Description: '',
          EnableVersioning: true,
          RowLimit: 30,
          Scope: 1,
          MajorVersionLimit: 50
        },
        { Description: '', EnableVersioning: true, RowLimit: 30, Scope: 0 }
      ),
      // A setting the object lacks is written, as before.
      { Scope: 1, MajorVersionLimit: 50 }
    )
  })

  it('compares the same text as the same, an object as JSON, and a ViewQuery as SharePoint stores it', () => {
    assert.deepEqual(
      differingSettings(
        {
          RowLimit: '30',
          Formatter: { a: 1 },
          ViewQuery: "<OrderBy><FieldRef Name='ID' /></OrderBy>"
        },
        {
          RowLimit: 30,
          Formatter: { a: 1 },
          ViewQuery: '<OrderBy><FieldRef Name="ID" /></OrderBy>'
        }
      ),
      {}
    )
    assert.deepEqual(
      differingSettings({ Formatter: { a: 2 } }, { Formatter: { a: 1 } }),
      { Formatter: { a: 2 } }
    )
  })
})
