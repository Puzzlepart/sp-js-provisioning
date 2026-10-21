/**
 * Checks what `npm pack` would publish: the compiled entry point and every handler must be in the
 * tarball. The 1.3.6 package shipped without `lib/`, and consumers' builds failed with TS2307 /
 * MODULE_NOT_FOUND. Run after `npm run build` (`npm pack` runs `prepack`, which builds).
 *
 * Usage: node .tasks/check-pack.js
 */
const { execFileSync } = require('node:child_process')
const { readdirSync } = require('node:fs')
const { join } = require('node:path')

const root = join(__dirname, '..')
const output = execFileSync('npm', ['pack', '--dry-run', '--json', '--ignore-scripts'], {
  cwd: root,
  encoding: 'utf8'
})
const [pack] = JSON.parse(output)
const files = new Set(pack.files.map((file) => file.path))

const handlers = readdirSync(join(root, 'src', 'handlers'))
  .filter((name) => name.endsWith('.ts'))
  .map((name) => `lib/handlers/${name.replace(/\.ts$/, '.js')}`)
const required = ['lib/index.js', 'lib/index.d.ts', 'schema.json', ...handlers]
const missing = required.filter((path) => !files.has(path))

if (missing.length) {
  console.error(
    `check-pack: ${missing.length} file(s) missing from the package: ${missing.join(', ')}. ` +
      'Run `npm run build` first.'
  )
  process.exit(1)
}
console.log(`check-pack: OK, ${pack.entryCount} files (${required.length} required present)`)
