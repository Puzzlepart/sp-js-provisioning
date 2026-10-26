import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  base64EncodeString,
  makeUrlRelative,
  mapStrings,
  replaceUrlTokens
} from '../../src/util'

const spfxContext = {
  pageContext: {
    site: {
      absoluteUrl: 'https://contoso.sharepoint.com/sites/hub',
      serverRelativeUrl: '/sites/hub'
    },
    web: { absoluteUrl: 'https://contoso.sharepoint.com/sites/hub/sub' }
  }
}

/**
 * Runs `fn` with a fake browser `window` (and `document`), removed afterwards.
 *
 * @param fake - The globals to set
 * @param fn - The code
 */
function withBrowser<T>(fake: Record<string, any>, fn: () => T): T {
  const g = globalThis as any
  const saved = { window: g.window, document: g.document }
  Object.assign(g, fake)
  try {
    return fn()
  } finally {
    g.window = saved.window
    g.document = saved.document
    if (saved.window === undefined) delete g.window
    if (saved.document === undefined) delete g.document
  }
}

describe('util: URL tokens without a browser', () => {
  it('resolves every URL token from the SPFx context, {hosturl} as the origin', () => {
    // Node has no `window`: this used to throw a ReferenceError.
    assert.equal(
      replaceUrlTokens(
        '{site}|{sitecollection}|{wpgallery}|{themegallery}|{hosturl}',
        { spfxContext }
      ),
      [
        '/sites/hub',
        'https://contoso.sharepoint.com/sites/hub',
        'https://contoso.sharepoint.com/sites/hub/_catalogs/wp',
        'https://contoso.sharepoint.com/sites/hub/_catalogs/theme/15',
        // No trailing ':' (it used to be `${protocol}//${host}:${port}`).
        'https://contoso.sharepoint.com'
      ].join('|')
    )
  })

  it('leaves {hosturl} and writes null for {site} without a context or a browser', () => {
    // Pinned: `{site}` without a site has always become "null".
    assert.equal(replaceUrlTokens('{hosturl} {site}', {}), '{hosturl} null')
  })

  it('makes a URL relative by removing its origin', () => {
    assert.equal(
      makeUrlRelative(
        'https://contoso.sharepoint.com/sites/hub/_catalogs/theme/15/a.spcolor'
      ),
      '/sites/hub/_catalogs/theme/15/a.spcolor'
    )
  })

  it("encodes a string's UTF-16LE bytes as base64 (indexed property bag keys)", () => {
    for (const key of [
      'GtSiteId',
      'Fase\u2013status',
      '\u0424\u0430\u0437\u0430'
    ])
      assert.equal(
        base64EncodeString(key),
        Buffer.from(key, 'utf16le').toString('base64'),
        key
      )
  })
})

describe('util: URL tokens in a browser', () => {
  const location = {
    protocol: 'https:',
    host: 'contoso.sharepoint.com',
    hostname: 'contoso.sharepoint.com',
    port: ''
  }

  it('reads a classic page context and the page location without an SPFx context', () => {
    const value = withBrowser(
      {
        window: {
          location,
          _spPageContextInfo: {
            siteAbsoluteUrl: 'https://contoso.sharepoint.com/sites/classic',
            siteServerRelativeUrl: '/sites/classic'
          }
        }
      },
      () => replaceUrlTokens('{site} {sitecollection} {hosturl}', {})
    )

    assert.equal(
      value,
      '/sites/classic https://contoso.sharepoint.com/sites/classic https://contoso.sharepoint.com'
    )
  })

  it('prefers the SPFx context over the page', () => {
    const value = withBrowser(
      { window: { location: { ...location, host: 'other.example' } } },
      () => replaceUrlTokens('{hosturl}', { spfxContext })
    )

    assert.equal(value, 'https://contoso.sharepoint.com')
  })

  it("removes only the page's origin in a browser", () => {
    const relative = withBrowser(
      {
        document: {
          location: { protocol: 'https:', hostname: 'contoso.sharepoint.com' }
        }
      },
      () => [
        makeUrlRelative('https://contoso.sharepoint.com/sites/hub/a'),
        makeUrlRelative('https://cdn.example/a')
      ]
    )

    assert.deepEqual(relative, ['/sites/hub/a', 'https://cdn.example/a'])
  })
})

describe('util: mapStrings', () => {
  it('replaces in every string, keys included, and copies the rest as is', () => {
    const value = {
      title: '{x}',
      '{x}Key': 1,
      list: ['{x}', 2, null, { deep: '{x}{x}' }],
      flag: false
    }
    const mapped = mapStrings(value, (text) => text.replace(/{x}/g, 'a "b"'))

    assert.deepEqual(mapped, {
      title: 'a "b"',
      'a "b"Key': 1,
      list: ['a "b"', 2, null, { deep: 'a "b"a "b"' }],
      flag: false
    })
    // A copy: the template is left as it was.
    assert.equal(value.title, '{x}')
    assert.notEqual(mapped.list, value.list)
  })
})
