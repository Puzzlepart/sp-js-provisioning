/* eslint-disable no-console */
/**
 * Writes `tests/fixtures/templates/all.json` from the handler fixtures (`composeAll`).
 * `npm run build:fixtures`; `tests/unit/fixtures.test.ts` fails when the file is out of date.
 */
import * as fs from 'fs'
import * as path from 'path'
import { composeAll, TEMPLATES } from './fixtures'

const file = path.join(TEMPLATES, 'all.json')
fs.writeFileSync(file, `${JSON.stringify(composeAll(), null, 2)}\n`)
console.log(`Wrote ${path.relative(process.cwd(), file)}`)
