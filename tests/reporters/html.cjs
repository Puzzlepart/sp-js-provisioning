/**
 * The HTML reporter as `node --test` loads it: `--test-reporter=./tests/reporters/html.cjs`.
 *
 * node --test loads a reporter with Node's own loader, which tsx's hooks do not reach: Node 22
 * strips the types of `html.ts` itself, and then cannot resolve its extensionless imports
 * (`../../debug/lib/redact`). This file loads `html.ts` through tsx instead, as the tests are.
 */
const tsx = require('tsx/cjs/api')

module.exports = tsx.require('./html', __filename).default
