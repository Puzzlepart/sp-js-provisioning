import assert from 'node:assert/strict'
import * as fs from 'fs'
import { rmSync } from 'node:fs'
import * as os from 'os'
import * as path from 'path'
import { describe, it } from 'node:test'
import { redactXml } from '../../debug/lib/redact'
import { checkReports } from '../reporters/check-reports'

const tenant = ['north', 'wind'].join('')

describe('check-reports', () => {
  it('passes redacted reports and names the file, line and kind of what is not', () => {
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), 'spjs-reports-'))
    try {
      const failure = `<failure message="at https://${tenant}.sharepoint.com/sites/hemmelig by kari@${tenant}.onmicrosoft.com">`
      fs.mkdirSync(path.join(dir, 'integration'))
      fs.writeFileSync(
        path.join(dir, 'integration.xml'),
        `<testsuites>\n${redactXml(failure)}\n</testsuites>\n`
      )
      fs.writeFileSync(
        path.join(dir, 'integration', 'index.html'),
        '<p>https://&lt;tenant&gt;.sharepoint.com/sites/&lt;site&gt;/spjs-itest-lists-1</p>'
      )
      fs.writeFileSync(
        path.join(dir, 'image.png'),
        `https://${tenant}.sharepoint.com`
      )
      assert.deepEqual(checkReports(dir), { files: 2, findings: [] })

      fs.writeFileSync(
        path.join(dir, 'reachability.xml'),
        `<testsuites>\n${failure}\n</testsuites>\n`
      )
      const { files, findings } = checkReports(dir)
      assert.equal(files, 3)
      assert.deepEqual(findings, [
        'reachability.xml:2: <user>, <tenant>, <site>'
      ])
      assert.ok(!findings.join().includes(tenant))
    } finally {
      rmSync(dir, { recursive: true, force: true })
    }
  })
})
