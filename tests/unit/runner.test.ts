import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import * as crypto from 'crypto'
import * as fs from 'fs'
import { rmSync } from 'node:fs'
import * as os from 'os'
import * as path from 'path'
import * as forge from 'node-forge'
import { spfi } from '@pnp/sp'
import '@pnp/sp/presets/all'
import { readCertificate, sessionCookies, signedIn } from '../../debug/lib/auth'
import {
  applyLocale,
  cloudHubSubset,
  handlerOrder,
  NODE_UNSUPPORTED,
  parseArgs,
  planHandlers,
  resolveTemplate,
  UsageError
} from '../../debug/lib/cli'
import { authMode, configured, readEnv, ROOT } from '../../debug/lib/env'
import { sourceOf, startRun } from '../../debug/lib/logs'
import { FunctionListener, Logger, LogLevel } from '@pnp/logging'
import { DefaultHandlerSort } from '../../src/handlers/exports'

const handlers = handlerOrder(DefaultHandlerSort)

/**
 * Asserts that a call fails with a usage error (exit code 2) whose message matches.
 *
 * @param call - The call
 * @param message - What the message must contain
 */
function assertUsageError(call: () => unknown, message: RegExp): void {
  try {
    call()
  } catch (error) {
    assert.ok(error instanceof UsageError, `not a UsageError: ${error}`)
    assert.equal((error as UsageError).exitCode, 2)
    assert.match((error as Error).message, message)
    return
  }
  assert.fail('expected a UsageError')
}

/**
 * A temporary folder, removed by the caller.
 */
function tempDir(): string {
  return fs.mkdtempSync(path.join(os.tmpdir(), 'spjs-runner-test-'))
}

describe('runner: arguments', () => {
  it('parses flags, handlers and parameters', () => {
    const args = parseArgs([
      '--package',
      'template.json',
      '--handlers',
      'Security, Lists',
      '--parameters',
      'stamp=1,url=https://example.org/?a=b',
      '--strict',
      '--twice'
    ])

    assert.deepEqual(args.handlers, ['Security', 'Lists'])
    // A value may itself contain `=`.
    assert.deepEqual(args.parameters, {
      stamp: '1',
      url: 'https://example.org/?a=b'
    })
    assert.equal(args.mode, 'hub')
    assert.ok(args.strict && args.twice && !args.dryRun)
    assert.equal(args.site, undefined)
  })

  it('rejects what it does not understand with exit code 2', () => {
    assertUsageError(() => parseArgs(['--bogus']), /Unknown argument: --bogus/)
    assertUsageError(
      () => parseArgs(['--package', 'x.json', '--mode', 'project']),
      /Invalid --mode "project"/
    )
    assertUsageError(
      () => parseArgs(['--strict']),
      /Missing required --package/
    )
    assertUsageError(
      () => parseArgs(['--package', '--strict']),
      /Missing value for --package/
    )
    assertUsageError(
      () => parseArgs(['--package', 'x.json', '--parameters', 'novalue']),
      /Invalid --parameters entry "novalue"/
    )
  })
})

describe('runner: handler plan', () => {
  it('runs the handlers in the engine order, Security and WebSettings included', () => {
    // The runner's hand-kept list once dropped Security and WebSettings, so a
    // template's sections for them were silently not applied.
    assert.equal(handlers[0], 'Taxonomy')
    assert.equal(handlers[handlers.length - 1], 'Hooks')
    assert.ok(handlers.indexOf('Security') < handlers.indexOf('Lists'))
    assert.ok(handlers.includes('WebSettings'))
    assert.deepEqual(
      [...handlers].sort(),
      Object.keys(DefaultHandlerSort).sort()
    )
  })

  it('plans the template sections Node can run and names the skipped ones', () => {
    const plan = planHandlers(
      {
        $schema: '../schema.json',
        Parameters: { a: 'b' },
        WebSettings: {},
        Files: [],
        Lists: [],
        Security: {},
        Taxonomy: {}
      },
      handlers
    )

    assert.deepEqual(plan.run, ['Security', 'Lists', 'WebSettings'])
    assert.deepEqual(plan.skipped, ['Taxonomy', 'Files'])
  })

  it('validates --handlers against the engine', () => {
    assert.deepEqual(planHandlers({}, handlers, ['ContentTypes']).run, [
      'ContentTypes'
    ])
    assertUsageError(
      () => planHandlers({}, handlers, ['Lists', 'Bogus']),
      /Unknown handler\(s\): Bogus\. Valid handlers: Taxonomy, SiteFields/
    )
  })

  it('keeps the smoke template runnable from Node', () => {
    // tests/fixtures/templates/smoke.json is the first live run; it must not
    // depend on a handler the runner skips in Node.
    const smoke = JSON.parse(
      fs.readFileSync(
        path.join(ROOT, 'tests', 'fixtures', 'templates', 'smoke.json'),
        'utf8'
      )
    )
    const plan = planHandlers(smoke, handlers)

    assert.deepEqual(plan.run, ['SiteFields', 'Security', 'Lists'])
    assert.deepEqual(plan.skipped, [])
    for (const name of NODE_UNSUPPORTED) assert.equal(smoke[name], undefined)
  })
})

describe('runner: cloud subset', () => {
  // Mirrors PP365 PackageInstaller.provisionCloudTemplateHubDependencies
  // (releases/1.15).
  const schema = {
    SiteFields: ['<Field />'],
    ContentTypes: [{ ID: '0x0100AA' }],
    Taxonomy: { TermGroup: { Id: 'g' }, TermSets: [] },
    Security: { SiteGroups: [] },
    Files: [{}],
    Lists: [
      { Title: 'Bindings', ContentTypeBindings: [{ ContentTypeID: '0x01' }] },
      {
        Title: 'Bindings and rows',
        ContentTypeBindings: [{ ContentTypeID: '0x01' }],
        DataRows: { Rows: [] }
      },
      { Title: 'Hub rows', DataRows: { Rows: [] } },
      { Title: 'Project library' },
      { Title: 'Content source', DataRows: { Rows: [] } }
    ]
  }
  const manifest = {
    provisioning: { listContent: [{ sourceList: 'Content source' }] }
  }

  it('keeps fields, content types, binding and hub-row lists and taxonomy', () => {
    const subset = cloudHubSubset(schema, manifest)

    assert.deepEqual(Object.keys(subset), [
      'SiteFields',
      'ContentTypes',
      'Lists',
      'Taxonomy'
    ])
    assert.deepEqual(
      subset.Lists.map((list: any) => list.Title),
      ['Bindings', 'Bindings and rows', 'Hub rows']
    )
  })

  it('drops taxonomy when asked', () => {
    assert.equal(cloudHubSubset(schema, manifest, true).Taxonomy, undefined)
  })
})

describe('runner: packages', () => {
  it('resolves a package folder to its localized hub template', () => {
    const dir = tempDir()
    try {
      fs.writeFileSync(
        path.join(dir, 'manifest.json'),
        JSON.stringify({
          name: 'pp-test',
          provisioning: {
            hubTemplate: 'hub.json',
            localized: { 'en-US': { hubTemplate: 'hub.en-US.json' } }
          }
        })
      )
      fs.writeFileSync(path.join(dir, 'hub.json'), '{"Lists":[]}')
      fs.writeFileSync(path.join(dir, 'hub.en-US.json'), '{"SiteFields":[]}')

      assert.deepEqual(resolveTemplate({ package: dir }).schema, { Lists: [] })
      const english = resolveTemplate({ package: dir, lang: 'en-US' })
      assert.deepEqual(english.schema, { SiteFields: [] })
      assert.equal(english.label, 'pp-test -> hub.en-US.json')
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('merges a locale the way PackageInstaller does', () => {
    const manifest = {
      provisioning: {
        hubTemplate: 'hub.json',
        listContent: [{ sourceList: 'Usikkerhet' }],
        localized: {
          'en-US': { listContent: [{ sourceList: 'Uncertainties' }] }
        }
      }
    }

    const english = applyLocale(manifest, 'en-US')
    assert.equal(english.provisioning.hubTemplate, 'hub.json')
    assert.deepEqual(english.provisioning.listContent, [
      { sourceList: 'Uncertainties' }
    ])
    assert.equal(applyLocale(manifest, 'sv-SE'), manifest)
  })
})

describe('runner: configuration', () => {
  it('treats empty values and .env.example placeholders as unset', () => {
    assert.equal(configured('A', { A: '' }), undefined)
    assert.equal(configured('A', { A: '  ' }), undefined)
    assert.equal(
      configured('A', { A: 'https://<tenant>.sharepoint.com/sites/x' }),
      undefined
    )
    assert.equal(configured('A', { A: ' value ' }), 'value')
  })

  it('reads the site, the sign-in and the cookie file', () => {
    const configuration = readEnv({
      INTEGRATION_SITE_URL: 'https://contoso.sharepoint.com/sites/spjs-itest/',
      CI_TENANT: 'contoso.onmicrosoft.com',
      CI_CLIENT_ID: 'id'
    })

    assert.equal(
      configuration.siteUrl,
      'https://contoso.sharepoint.com/sites/spjs-itest'
    )
    // The certificate needs all three variables.
    assert.equal(configuration.certificate, undefined)
    assert.equal(
      configuration.cookiesFile,
      path.resolve(
        ROOT,
        '..',
        'prosjektportalen365',
        'e2e',
        '.auth',
        'user.json'
      )
    )
  })

  it('picks the forced mode, else the certificate, else existing cookies', () => {
    const certificate = readEnv({
      CI_TENANT: 't',
      CI_CLIENT_ID: 'c',
      CI_CERT_BASE64: 'p'
    })
    const none = readEnv({})

    assert.equal(
      authMode(certificate, () => true),
      'certificate'
    )
    assert.equal(
      authMode(
        readEnv({
          CI_TENANT: 't',
          CI_CLIENT_ID: 'c',
          CI_CERT_BASE64: 'p',
          INTEGRATION_AUTH: 'cookies'
        }),
        () => false
      ),
      'cookies'
    )
    assert.equal(
      authMode(none, () => true),
      'cookies'
    )
    assert.equal(
      authMode(none, () => false),
      undefined
    )
  })
})

describe('runner: cookie sign-in', () => {
  const now = Date.UTC(2026, 9, 21)

  /**
   * Writes a Playwright storage state with the given cookies.
   *
   * @param cookies - The cookies
   */
  function storageState(dir: string, cookies: any[]): string {
    const file = path.join(
      dir,
      `${cookies[0].domain}-${cookies[0].expires}.json`
    )
    fs.writeFileSync(file, JSON.stringify({ cookies, origins: [] }))
    return file
  }

  it('sends the session cookies for the host only', () => {
    const dir = tempDir()
    try {
      const file = storageState(dir, [
        {
          name: 'FedAuth',
          value: 'a',
          domain: 'contoso.sharepoint.com',
          expires: now / 1000 + 3600
        },
        { name: 'rtFa', value: 'b', domain: '.sharepoint.com', expires: -1 },
        {
          name: 'other',
          value: 'c',
          domain: 'login.microsoftonline.com',
          expires: -1
        }
      ])

      assert.equal(
        sessionCookies(file, 'contoso.sharepoint.com', now),
        'FedAuth=a; rtFa=b'
      )
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })

  it('names the renew command when the session is missing or has expired', () => {
    const dir = tempDir()
    try {
      const expired = storageState(dir, [
        {
          name: 'FedAuth',
          value: 'a',
          domain: 'contoso.sharepoint.com',
          expires: now / 1000 - 1
        }
      ])
      const otherHost = storageState(dir, [
        {
          name: 'FedAuth',
          value: 'a',
          domain: 'fabrikam.sharepoint.com',
          expires: -1
        }
      ])

      for (const file of [expired, otherHost])
        assert.throws(
          () => sessionCookies(file, 'contoso.sharepoint.com', now),
          /npx playwright test --project=setup/
        )
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})

describe('runner: certificate sign-in', () => {
  /**
   * A self-signed certificate and its key in a PFX, base64, the way CI_CERT_BASE64 holds one.
   *
   * @param password - The PFX password: '' or none (null)
   */
  function pfx(password: string | null): { base64: string; der: string } {
    const keys = forge.pki.rsa.generateKeyPair(1024)
    const cert = forge.pki.createCertificate()
    cert.publicKey = keys.publicKey
    cert.serialNumber = '01'
    cert.validity.notBefore = new Date(Date.UTC(2026, 0, 1))
    cert.validity.notAfter = new Date(Date.UTC(2027, 0, 1))
    const subject = [{ name: 'commonName', value: 'spjs-test' }]
    cert.setSubject(subject)
    cert.setIssuer(subject)
    cert.sign(keys.privateKey, forge.md.sha256.create())
    const asn1 = forge.pkcs12.toPkcs12Asn1(
      keys.privateKey,
      [cert],
      (password as unknown) as string,
      { algorithm: '3des' }
    )
    return {
      base64: forge.util.encode64(forge.asn1.toDer(asn1).getBytes()),
      der: forge.asn1.toDer(forge.pki.certificateToAsn1(cert)).getBytes()
    }
  }

  it('reads the key and the SHA-256 thumbprint with an empty or no password', () => {
    for (const password of ['', null]) {
      const { base64, der } = pfx(password)
      const certificate = readCertificate(base64)

      assert.match(certificate.privateKey, /BEGIN RSA PRIVATE KEY/)
      assert.equal(
        certificate.thumbprintSha256,
        crypto
          .createHash('sha256')
          .update(Buffer.from(der, 'binary'))
          .digest('hex')
      )
    }
  })

  it('names the variable, never its value, when the PFX is not valid', () => {
    assert.throws(
      () => readCertificate('bm90IGEgcGZ4'),
      (error: Error) =>
        /CI_CERT_BASE64/.test(error.message) &&
        !error.message.includes('bm90IGEgcGZ4')
    )
  })
})

describe('runner: log lines', () => {
  it('attributes a line to the handler that wrote it', () => {
    assert.equal(
      sourceOf('(provision #2) (Lists): (processList): Processing list X'),
      'Lists'
    )
    assert.equal(
      sourceOf(' (WebProvisioner): (loadContext): Failed to load'),
      'WebProvisioner'
    )
    assert.equal(sourceOf('Code execution scope started'), undefined)
  })
})

describe('runner: sign-in pipeline', () => {
  it('sends the session cookie, a browser user agent and a request digest on writes', async () => {
    // Real PnPjs 4 with the runner's cookie behaviour and a fake transport:
    // what SharePoint would receive, without a tenant.
    const site = 'https://contoso.sharepoint.com/sites/spjs-itest'
    const directory = tempDir()
    try {
      const file = path.join(directory, 'user.json')
      fs.writeFileSync(
        file,
        JSON.stringify({
          cookies: [
            {
              name: 'FedAuth',
              value: 'session',
              domain: 'contoso.sharepoint.com',
              expires: -1
            }
          ]
        })
      )
      const seen: Array<{ url: string; method: string; headers: Headers }> = []
      const sp = spfi(site).using(
        signedIn('cookies', site, readEnv({ INTEGRATION_COOKIES_FILE: file })),
        (instance) => {
          instance.on.send.replace(async (url: URL, init: any) => {
            seen.push({
              url: url.toString(),
              method: init.method,
              headers: new Headers(init.headers)
            })
            const body = url.toString().includes('/_api/contextinfo')
              ? { FormDigestValue: 'digest-1', FormDigestTimeoutSeconds: 1800 }
              : { Title: 'SPJS Itest' }
            return new Response(JSON.stringify(body), {
              status: 200,
              headers: { 'Content-Type': 'application/json' }
            })
          })
          return instance
        }
      )

      const web = await sp.web.select('Title')<{ Title: string }>()
      await sp.web.update({ Description: 'Written by the test' })

      assert.equal(web.Title, 'SPJS Itest')
      const read = seen[0]
      assert.equal(
        decodeURIComponent(read.url),
        `${site}/_api/web?$select=Title`
      )
      assert.equal(read.headers.get('Cookie'), 'FedAuth=session')
      assert.match(read.headers.get('User-Agent'), /^Mozilla\/5\.0/)
      const write = seen.find(
        (request) =>
          request.method === 'POST' && !request.url.includes('contextinfo')
      )
      assert.ok(write, `no write in ${JSON.stringify(seen.map((r) => r.url))}`)
      assert.equal(write.headers.get('X-RequestDigest'), 'digest-1')
      assert.equal(write.headers.get('Cookie'), 'FedAuth=session')
    } finally {
      rmSync(directory, { recursive: true, force: true })
    }
  })
})

describe('runner: log collection per run', () => {
  it('clears the collector for each run and leaves the other listeners alone', () => {
    // The engine subscribes its console listener once per process, so the
    // runner must not clear the logger's listeners between runs.
    const other: string[] = []
    Logger.subscribe(FunctionListener((entry) => other.push(entry.message)))
    try {
      const first = startRun()
      Logger.write(
        '(provision #1) (Lists): (processList): first',
        LogLevel.Error
      )
      const firstErrors = first.errors().map((entry) => entry.message)
      const second = startRun()
      Logger.write(
        '(provision #2) (Lists): (processList): second',
        LogLevel.Error
      )

      assert.equal(second, first)
      assert.deepEqual(firstErrors, [
        '(provision #1) (Lists): (processList): first'
      ])
      assert.deepEqual(
        second.errors().map((entry) => entry.source),
        ['Lists']
      )
      assert.equal(second.errors().length, 1)
      assert.equal(other.length, 2)
    } finally {
      Logger.clearSubscribers()
    }
  })
})
