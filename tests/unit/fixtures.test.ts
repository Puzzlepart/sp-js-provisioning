import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, it } from 'node:test'
import {
  ALL_PARTS,
  composeAll,
  loadFixture,
  TEMPLATES
} from '../fixtures/fixtures'

/**
 * `all.json` is made from the handler fixtures (`npm run build:fixtures`), so `everything.test.ts`
 * and `provision --package tests/fixtures/templates/all.json` apply what the suites apply.
 */
describe('fixtures: all.json', () => {
  it('is the composition of the handler fixtures (npm run build:fixtures)', () => {
    const file = JSON.parse(readFileSync(join(TEMPLATES, 'all.json'), 'utf8'))
    assert.deepEqual(
      file,
      composeAll(),
      'tests/fixtures/templates/all.json is out of date: run npm run build:fixtures'
    )
  })

  it('holds every list, site column and section of its parts, once', () => {
    const all = composeAll()
    const titles = all.Lists.map((list: any) => list.Title)
    assert.equal(new Set(titles).size, titles.length, 'a list title twice')
    const names = all.SiteFields.map(
      (xml: string) => /Name="([^"]+)"/.exec(xml)[1]
    )
    assert.equal(new Set(names).size, names.length, 'a site column twice')
    for (const part of ALL_PARTS) {
      const fixture = loadFixture(part, { user: 'user@contoso.com' })
      for (const section of Object.keys(fixture))
        assert.ok(all[section] !== undefined, `${part}: ${section} is missing`)
    }
  })

  it('fills its placeholders with loadFixture, and needs no hook listener', () => {
    const text = JSON.stringify(
      loadFixture('all', { user: 'user@contoso.com' })
    )
    assert.doesNotMatch(text, /@@/)
    assert.throws(() => loadFixture('all'), /needs a value for @@user@@/)
  })
})
