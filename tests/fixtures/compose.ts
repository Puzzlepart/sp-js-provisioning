/* eslint-disable no-console */
/**
 * Writes `tests/fixtures/templates/all.node.json` from the handler fixtures (`composeAllNode`).
 * `npm run build:fixtures`; `tests/unit/fixtures.test.ts` fails when the file is out of date.
 */
import * as fs from 'fs'
import * as path from 'path'
import { composeAllNode, TEMPLATES } from './fixtures'

const file = path.join(TEMPLATES, 'all.node.json')
fs.writeFileSync(file, `${JSON.stringify(composeAllNode(), null, 2)}\n`)
console.log(`Wrote ${path.relative(process.cwd(), file)}`)
