import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import {
  describeError,
  PLACEHOLDERS,
  redact,
  redactXml,
  unredacted
} from '../../debug/lib/redact'

// Built at run time: nothing here is a real tenant, user or id, and a literal in a test file would
// be one of the repository's own.
const tenant = ['north', 'wind'].join('')
const guid = () =>
  ['3f2b9c1e', 'a7d4', '4e5f', '9b1a', '0c2d3e4f5a6b'].join('-')
const login = `kari.nordmann@${tenant}.onmicrosoft.com`

/** Runs `fn` with environment variables set, then restores them. */
function withEnv(values: Record<string, string | undefined>, fn: () => void) {
  const before: Record<string, string | undefined> = {}
  for (const name of Object.keys(values)) {
    before[name] = process.env[name]
    if (values[name] === undefined) delete process.env[name]
    else process.env[name] = values[name]
  }
  try {
    fn()
  } finally {
    for (const name of Object.keys(before))
      if (before[name] === undefined) delete process.env[name]
      else process.env[name] = before[name]
  }
}

const NONE = {
  INTEGRATION_USER: undefined,
  CI_CLIENT_ID: undefined,
  CI_TENANT: undefined,
  INTEGRATION_SITE_URL: undefined,
  CI_CERT_BASE64: undefined
}

describe('redaction: secrets', () => {
  it('removes cookie values, bearer tokens and certificates', () => {
    withEnv(NONE, () => {
      const cookie = 'A'.repeat(40)
      assert.equal(
        redact(
          `set-cookie: FedAuth=${cookie}; path=/ Cookie: FedAuth=${cookie}; rtFa=${cookie}`
        ),
        'set-cookie: <redacted>'
      )
      assert.equal(
        redact(`Cookie: rtFa=${cookie}; x=1`),
        'Cookie: rtFa=<redacted>; x=1'
      )
      assert.equal(
        redact(`Authorization: Bearer ${'ey' + 'J'.repeat(20)}`),
        'Authorization: Bearer <redacted>'
      )
      const pfx = 'MII' + 'Kc'.repeat(60)
      assert.equal(redact(`pfx ${pfx}== end`), 'pfx <redacted> end')
      // A PEM block: its markers and each line of its body, also when the runner relays a log a
      // line at a time. A marker alone never swallows what follows it.
      assert.equal(
        redact(
          `key:\n-----BEGIN PRIVATE KEY-----\n${'Qw'.repeat(32)}\n${'Zx'.repeat(
            32
          )}\n-----END PRIVATE KEY-----\nafter`
        ),
        'key:\n<redacted>\n<redacted>\n<redacted>\n<redacted>\nafter'
      )
      assert.equal(
        redact(
          `-----BEGIN CERTIFICATE-----${'Qw'.repeat(
            8
          )}-----END CERTIFICATE----- x`
        ),
        '<redacted> x'
      )
      assert.equal(
        redact('expected -----BEGIN CERTIFICATE----- in <testcase name="x"/>'),
        'expected <redacted> in <testcase name="x"/>'
      )
      assert.equal(redact('Qw'.repeat(32)), '<redacted>')
      assert.equal(redact('a short line of text'), 'a short line of text')
    })
  })

  it('removes the configured certificate wherever it is', () => {
    const certificate = 'Zm9v' + 'YmFy'.repeat(30)
    withEnv({ ...NONE, CI_CERT_BASE64: certificate }, () => {
      assert.equal(redact(`cert=${certificate};`), 'cert=<redacted>;')
    })
  })
})

describe('redaction: the tenant, its site and its users', () => {
  it('replaces SharePoint hosts, onmicrosoft domains and the site in a path, not the example ones', () => {
    withEnv(NONE, () => {
      assert.equal(
        redact(
          `https://${tenant}.sharepoint.com/sites/hemmelig/spjs-itest-lists-261110121935/_api/web and https://${tenant}-admin.sharepoint.com`
        ),
        'https://<tenant>.sharepoint.com/sites/<site>/spjs-itest-lists-261110121935/_api/web and https://<tenant>-admin.sharepoint.com'
      )
      assert.equal(
        redact(
          `Kan ikke åpne /sites/hemmelig/spjs-itest-x-1/SitePages/Hjem.aspx: finnes ikke.`
        ),
        'Kan ikke åpne /sites/<site>/spjs-itest-x-1/SitePages/Hjem.aspx: finnes ikke.'
      )
      assert.equal(redact('/teams/Styret (web)'), '/teams/<site> (web)')
      assert.equal(
        redact(`${tenant}.onmicrosoft.com`),
        '<tenant>.onmicrosoft.com'
      )
      // Documentation's hosts keep their paths.
      for (const example of [
        'https://contoso.sharepoint.com/sites/demo/sub',
        'https://fabrikam-my.sharepoint.com/personal/x',
        'contoso.onmicrosoft.com',
        'kari@contoso.com and ola@example.org'
      ])
        assert.equal(redact(example), example)
    })
  })

  it('replaces e-mail addresses and logins, also URL-encoded and in a OneDrive path', () => {
    withEnv(NONE, () => {
      assert.equal(
        redact(`i:0#.f|membership|${login}`),
        'i:0#.f|membership|<user>'
      )
      assert.equal(redact(`Owner: ${login.toUpperCase()}.`), 'Owner: <user>.')
      assert.equal(
        redact(
          `ensureuser('i%3A0%23.f%7Cmembership%7Ckari%40${tenant}.onmicrosoft.com')`
        ),
        "ensureuser('<user>')"
      )
      assert.equal(
        redact(`/personal/kari_nordmann_${tenant}_onmicrosoft_com/Documents`),
        '/personal/<user>/Documents'
      )
      // Not addresses: a package at a version, a scoped package, a path.
      for (const text of [
        '> sp-js-provisioning@1.5.1 test:integration',
        'node_modules/@pnp/queryable/behaviors/parsers.js:95:16',
        'tsx@4.19.2'
      ])
        assert.equal(redact(text), text)
    })
  })

  it('replaces what the run was configured with, in any case', () => {
    const clientId = guid().replace(/^3f2b/, '77aa')
    withEnv(
      {
        ...NONE,
        INTEGRATION_USER: login,
        CI_CLIENT_ID: clientId,
        CI_TENANT: 'firma.example.no',
        INTEGRATION_SITE_URL: `https://${tenant}.sharepoint.com/omrader/Hemmelig/`
      },
      () => {
        assert.equal(
          redact(
            `AADSTS700016: Application with identifier '${clientId.toUpperCase()}' was not found in the directory 'firma.example.no'.`
          ),
          "AADSTS700016: Application with identifier '<client-id>' was not found in the directory '<tenant>'."
        )
        assert.equal(
          redact(`web /omrader/hemmelig/spjs-itest-a-1 (${login})`),
          'web /omrader/<site>/spjs-itest-a-1 (<user>)'
        )
        assert.equal(redact('/omrader/hemmeligere'), '/omrader/hemmeligere')
      }
    )
    // Placeholders of .env.example are not values.
    withEnv(
      {
        ...NONE,
        INTEGRATION_USER: '<user>@<tenant>.onmicrosoft.com',
        CI_TENANT: '<tenant>.onmicrosoft.com'
      },
      () => {
        assert.equal(redact('nothing to hide'), 'nothing to hide')
      }
    )
  })
})

describe('redaction: ids', () => {
  it("replaces a GUID that is not one of the repository's own", () => {
    withEnv(NONE, () => {
      assert.equal(
        redact(`List {${guid()}} of web ${guid().toUpperCase()}`),
        'List {<guid>} of web <guid>'
      )
      // A field id of a fixture (tests/fixtures/templates/sitefields.json) is public: kept.
      const own = 'cf9d3103-f63b-4666-a9ca-da4098eb1e90'
      assert.equal(redact(`Field ref '${own}'`), `Field ref '${own}'`)
      assert.equal(
        redact(`Field ref '{${own.toUpperCase()}}'`),
        `Field ref '{${own.toUpperCase()}}'`
      )
      // A content type id is no GUID.
      assert.equal(
        redact('0x0100CBEAE584066D4D87970DD5EDFDFE7F06'),
        '0x0100CBEAE584066D4D87970DD5EDFDFE7F06'
      )
    })
  })
})

describe('redaction: one rule, applied once or twice', () => {
  const samples = () => [
    `https://${tenant}.sharepoint.com/sites/hemmelig/sub (${login}) {${guid()}}`,
    `set-cookie: FedAuth=${'B'.repeat(50)}`,
    `Cookie: rtFa=${'B'.repeat(50)}; x=1`,
    `Bearer ${'ey' + 'K'.repeat(30)}`,
    `/teams/x and ${tenant}.onmicrosoft.com`,
    'MII' + 'Ab'.repeat(40)
  ]

  it('changes nothing in a text it has redacted', () => {
    withEnv(NONE, () => {
      for (const sample of samples()) {
        const once = redact(sample)
        assert.notEqual(once, sample)
        assert.equal(redact(once), once, once)
        assert.deepEqual(unredacted(once), [])
        // As a report holds it: escaped.
        assert.deepEqual(unredacted(redactXml(sample)), [], redactXml(sample))
      }
    })
  })

  it('says where a text is not redacted and what kind of value it holds, never the value', () => {
    withEnv(NONE, () => {
      const text = [
        'fine',
        `at https://${tenant}.sharepoint.com/sites/hemmelig`,
        `by ${login}`,
        `FedAuth=${'C'.repeat(50)}`,
        `id ${guid()}`
      ].join('\n')
      assert.deepEqual(unredacted(text), [
        { line: 2, kinds: ['tenant', 'site'] },
        { line: 3, kinds: ['user'] },
        { line: 4, kinds: ['redacted'] },
        { line: 5, kinds: ['guid'] }
      ])
      assert.ok(!JSON.stringify(unredacted(text)).includes(tenant))
    })
  })

  it('escapes every placeholder for XML', () => {
    withEnv(NONE, () => {
      assert.equal(
        redactXml(
          `<failure message="at https://${tenant}.sharepoint.com/sites/x by ${login} {${guid()}}">`
        ),
        '<failure message="at https://&lt;tenant&gt;.sharepoint.com/sites/&lt;site&gt; by &lt;user&gt; {&lt;guid&gt;}">'
      )
      for (const name of PLACEHOLDERS)
        assert.equal(redactXml(`<${name}>`), `&lt;${name}&gt;`)
    })
  })

  it('describes an error by its message and status, never its response', () => {
    withEnv(NONE, () => {
      const error = Object.assign(
        new Error(`Forbidden at https://${tenant}.sharepoint.com/sites/x`),
        {
          status: 403,
          response: { headers: { 'set-cookie': `FedAuth=${'A'.repeat(40)}` } }
        }
      )
      assert.equal(
        describeError(error),
        'Forbidden at https://<tenant>.sharepoint.com/sites/<site> (403)'
      )
      assert.equal(describeError('plain'), 'plain')
      assert.equal(describeError(undefined), 'undefined')
    })
  })
})
