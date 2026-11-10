import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import htmlReporter, {
  createReportCollector,
  describeFailure,
  IReporterEvent,
  renderReport,
  titleFor
} from '../reporters/html'

// Ported with the reporter from prosjektportalen365's integration/tests/html-reporter.test.ts.
const ROOT = '/work/repo'
const FILE = `${ROOT}/tests/integration/lists.test.ts`
const CONTEXT = {
  title: 'sp-js-provisioning integration tests',
  generatedAt: new Date('2026-10-14T12:00:00Z')
}

/** Node's `ERR_TEST_FAILURE` wrapper around what a test threw, as it crosses from the file's process. */
function failure(failureType: string, cause: unknown, message = 'test failed') {
  return { code: 'ERR_TEST_FAILURE', failureType, message, cause }
}

/** The events of one test file as Node reports them in declaration order (process isolation). */
function fileEvents(): IReporterEvent[] {
  const at = (nesting: number, name: string, line: number) => ({
    nesting,
    name,
    line,
    column: 3,
    file: FILE,
    entryFile: FILE
  })
  return [
    { type: 'test:start', data: at(0, 'Lists', 10) },
    { type: 'test:start', data: at(1, 'creates the lists', 11) },
    {
      type: 'test:pass',
      data: {
        ...at(1, 'creates the lists', 11),
        details: { type: 'test', duration_ms: 812 }
      }
    },
    {
      type: 'test:diagnostic',
      data: { nesting: 1, message: '3 lists', file: FILE, line: 11 }
    },
    { type: 'test:start', data: at(1, 'stores a <b>column</b>', 20) },
    {
      type: 'test:fail',
      data: {
        ...at(1, 'stores a <b>column</b>', 20),
        details: {
          type: 'test',
          duration_ms: 40,
          error: failure('testCodeFailure', {
            name: 'AssertionError',
            message: "Expected values to be strictly equal:\n\n'a' !== 'b'\n",
            code: 'ERR_ASSERTION',
            stack:
              "AssertionError [ERR_ASSERTION]: Expected values to be strictly equal:\n\n'a' !== 'b'\n\n" +
              `    at TestContext.<anonymous> (${FILE}:21:12)\n` +
              '    at Test.runInAsyncScope (node:async_hooks:226:14)'
          })
        }
      }
    },
    { type: 'test:start', data: at(1, 'adds a script link', 30) },
    {
      type: 'test:pass',
      data: {
        ...at(1, 'adds a script link', 30),
        skip: 'custom script is blocked on the site',
        details: { type: 'test', duration_ms: 0 }
      }
    },
    {
      type: 'test:fail',
      data: {
        ...at(0, 'Lists', 10),
        details: {
          type: 'suite',
          duration_ms: 900,
          error: failure(
            'subtestsFailed',
            '1 subtest failed',
            '1 subtest failed'
          )
        }
      }
    },
    {
      type: 'test:stdout',
      data: {
        file: 'tests/integration/lists.test.ts',
        message: `wrote ${FILE}\n`,
        entryFile: FILE
      }
    }
  ]
}

describe('the HTML reporter', () => {
  it('builds the tree of suites and tests in declaration order, with statuses and reasons', () => {
    const collector = createReportCollector(ROOT, ROOT)
    fileEvents().forEach(collector.add)
    const [file] = collector.report().files
    assert.equal(file.path, 'tests/integration/lists.test.ts')
    const [suite] = file.nodes
    assert.equal(suite.kind, 'suite')
    assert.equal(suite.status, 'failed')
    assert.equal(suite.note, '1 subtest failed')
    assert.deepEqual(
      suite.children.map((node) => [node.name, node.status, node.note]),
      [
        ['creates the lists', 'passed', undefined],
        ['stores a <b>column</b>', 'failed', undefined],
        [
          'adds a script link',
          'skipped',
          'custom script is blocked on the site'
        ]
      ]
    )
    assert.deepEqual(suite.children[0].diagnostics, ['3 lists'])
    assert.deepEqual(file.output, [
      { stream: 'stdout', text: 'wrote tests/integration/lists.test.ts\n' }
    ])
  })

  it("shows what the test threw, with the root and Node's own frames removed from the stack", () => {
    const collector = createReportCollector(ROOT, ROOT)
    fileEvents().forEach(collector.add)
    const error = collector.report().files[0].nodes[0].children[1].error
    assert.equal(error?.failureType, 'testCodeFailure')
    assert.match(
      error?.message ?? '',
      /^AssertionError \[ERR_ASSERTION\]: Expected values to be strictly equal:\n\n'a' !== 'b'$/
    )
    assert.equal(
      error?.stack,
      '    at TestContext.<anonymous> (tests/integration/lists.test.ts:21:12)'
    )
  })

  it('names the hook that failed, and keeps a cause chain', () => {
    const hook = describeFailure(
      failure(
        'hookFailed',
        { name: 'Error', message: 'no site', stack: 'Error: no site' },
        'failed running before hook'
      ),
      ROOT
    )
    assert.equal(hook.message, 'failed running before hook\n\nError: no site')
    const chained = describeFailure(
      failure('testCodeFailure', {
        name: 'Error',
        message: 'outer',
        cause: { name: 'HttpError', message: 'inner', status: 403 }
      }),
      ROOT
    )
    assert.equal(
      chained.message,
      'Error: outer\n\nCaused by: HttpError: inner\n(status: 403)'
    )
  })

  it('marks a file that threw while loading, and keeps its error output', () => {
    const collector = createReportCollector(ROOT, ROOT)
    const broken = `${ROOT}/tests/unit/broken.test.ts`
    collector.add({
      type: 'test:stderr',
      data: {
        file: 'tests/unit/broken.test.ts',
        message: 'SyntaxError: x\n',
        entryFile: broken
      }
    })
    collector.add({
      type: 'test:start',
      data: {
        nesting: 0,
        name: 'tests/unit/broken.test.ts',
        line: 1,
        column: 1,
        file: broken
      }
    })
    collector.add({
      type: 'test:fail',
      data: {
        nesting: 0,
        name: 'tests/unit/broken.test.ts',
        file: broken,
        details: {
          type: 'test',
          duration_ms: 40,
          error: failure('testCodeFailure', undefined)
        }
      }
    })
    const [file] = collector.report().files
    assert.equal(file.loadFailed, true)
    assert.equal(
      file.nodes[0].error?.message,
      "The file's process failed before it reported a test; its error output is below."
    )
    assert.equal(file.output[0].text, 'SyntaxError: x\n')
  })

  it('tells a process that exited after its tests from one that never loaded, and shows a test that never finished as not finished', () => {
    const collector = createReportCollector(ROOT, ROOT)
    const events = fileEvents()
    // The file's process died after the first test passed: no pass or fail for the rest.
    events.slice(0, 4).forEach(collector.add)
    collector.add({
      type: 'test:start',
      data: {
        nesting: 0,
        name: 'tests/integration/lists.test.ts',
        line: 1,
        column: 1,
        file: FILE
      }
    })
    collector.add({
      type: 'test:fail',
      data: {
        nesting: 0,
        name: 'tests/integration/lists.test.ts',
        file: FILE,
        details: {
          type: 'test',
          duration_ms: 900,
          error: {
            ...failure('testCodeFailure', undefined),
            exitCode: null,
            signal: 'SIGKILL'
          }
        }
      }
    })
    const [file] = collector.report().files
    assert.equal(file.loadFailed, false)
    const [suite, processNode] = file.nodes
    assert.deepEqual(
      [suite.status, suite.note],
      ['cancelled', "did not finish: the file's process ended first"]
    )
    assert.equal(suite.children[0].status, 'passed')
    assert.equal(suite.children[0].note, undefined)
    assert.equal(
      processNode.error?.message,
      "The file's process was killed by SIGKILL after its tests had run."
    )
  })

  it('takes the verdict from the run summary, which counts no suites, and embeds it for the test report site', () => {
    const collector = createReportCollector(ROOT, ROOT)
    fileEvents().forEach(collector.add)
    collector.add({
      type: 'test:summary',
      data: { file: FILE, success: false, counts: { tests: 3 }, duration_ms: 1 }
    })
    collector.add({
      type: 'test:summary',
      data: {
        success: false,
        counts: {
          tests: 3,
          passed: 1,
          failed: 1,
          cancelled: 0,
          skipped: 1,
          todo: 0
        },
        duration_ms: 950
      }
    })
    const html = renderReport(collector.report(), CONTEXT)
    // What .tasks/publish-test-report.js reads the counts from.
    const embedded = /<script type="application\/json" id="report-summary">(.*?)<\/script>/.exec(
      html
    )?.[1]
    assert.deepEqual(JSON.parse(embedded ?? 'null'), {
      success: false,
      counts: {
        tests: 3,
        passed: 1,
        failed: 1,
        cancelled: 0,
        skipped: 1,
        todo: 0
      },
      durationMs: 950
    })
    assert.match(html, /<div class="verdict failed">Failed<\/div>/)
    assert.match(
      html,
      /<title>sp-js-provisioning integration tests<\/title>[\s\S]*<h1>sp-js-provisioning integration tests<\/h1>/
    )
  })

  it('escapes what tests and SharePoint wrote, so no name or message becomes markup', () => {
    const collector = createReportCollector(ROOT, ROOT)
    fileEvents().forEach(collector.add)
    const html = renderReport(collector.report(), CONTEXT)
    assert.ok(html.includes('stores a &lt;b&gt;column&lt;/b&gt;'))
    assert.ok(!html.includes('<b>column</b>'))
    assert.ok(!html.includes(ROOT))
  })

  it('redacts cookies, tokens, the tenant and its site from every text on the page', () => {
    // Built at run time, so this file does not trip the scan for tenant URLs itself.
    const tenant = ['north', 'wind'].join('')
    const cookie = 'A'.repeat(40)
    const site = `https://${tenant}.sharepoint.com/sites/hemmelig`
    const collector = createReportCollector(ROOT, ROOT)
    const at = { nesting: 0, line: 1, column: 1, file: FILE, entryFile: FILE }
    collector.add({
      type: 'test:start',
      data: { ...at, name: `reads ${site}` }
    })
    collector.add({
      type: 'test:fail',
      data: {
        ...at,
        name: `reads ${site}`,
        details: {
          type: 'test',
          duration_ms: 1,
          error: failure('testCodeFailure', {
            name: 'Error',
            message: `Forbidden at ${site}; Cookie: FedAuth=${cookie}`,
            stack: `Error: Forbidden at ${site}; Cookie: FedAuth=${cookie}\n    at x (${FILE}:2:3)`
          })
        }
      }
    })
    collector.add({
      type: 'test:stdout',
      data: {
        file: FILE,
        message: `Authorization: Bearer ${'ey' + 'J'.repeat(20)}\n`,
        entryFile: FILE
      }
    })
    const html = renderReport(collector.report(), CONTEXT)
    assert.ok(!html.includes(tenant))
    assert.ok(!html.includes('hemmelig'))
    assert.ok(!html.includes(cookie))
    assert.ok(!html.includes('eyJJ'))
    assert.ok(
      html.includes(
        'reads https://&lt;tenant&gt;.sharepoint.com/sites/&lt;site&gt;'
      )
    )
    assert.ok(html.includes('FedAuth=&lt;redacted&gt;'))
    assert.ok(html.includes('Bearer &lt;redacted&gt;'))
  })

  it('titles the page by the layer its test files are in', () => {
    assert.equal(
      titleFor(['tests/unit/lists.test.ts', 'tests/unit/csom.test.ts']),
      'sp-js-provisioning unit tests'
    )
    assert.equal(
      titleFor(['tests/integration/reachability.test.ts']),
      'sp-js-provisioning integration tests'
    )
    assert.equal(
      titleFor(['tests/unit/lists.test.ts', 'tests/integration/lists.test.ts']),
      'sp-js-provisioning tests'
    )
    assert.equal(titleFor([]), 'sp-js-provisioning tests')
  })

  it('writes a page however malformed the stream is, and never throws', async () => {
    async function* events(): AsyncIterableIterator<IReporterEvent> {
      yield { type: 'test:start' }
      yield { type: 'test:fail', data: { nesting: 3, details: null } }
      yield { type: 'test:diagnostic', data: 42 }
      yield { type: 'test:summary', data: { counts: 'none' } }
    }
    let html = ''
    for await (const chunk of htmlReporter(events())) html += chunk
    assert.match(html, /^<!doctype html>/)
    assert.match(html, /<\/html>\n$/)
  })
})
