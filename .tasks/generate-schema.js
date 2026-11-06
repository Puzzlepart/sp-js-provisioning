#!/usr/bin/env node
/**
 * Generates schema.json, the template's JSON Schema, from the template contract in src/schema.ts
 * (typescript-json-schema; TSDoc becomes `description`). `--check` writes nothing and exits 1 when
 * schema.json is not what the generator makes, so CI catches a schema.ts change without it.
 *
 * Two things are adjusted after generation:
 * - IHooks' `Headers` and `Body` are typed with DOM types in schema.ts; in a template they are a
 *   plain object of header strings and a JSON object the handler spreads into the request body.
 * - The root accepts any other key, as the engine does (a key without a handler is skipped).
 *
 * Usage: node .tasks/generate-schema.js [--check]
 */
const fs = require('fs')
const path = require('path')
const TJS = require('typescript-json-schema')

const root = path.resolve(__dirname, '..')
const target = path.join(root, 'schema.json')

/**
 * Sorts object keys recursively, so the output does not depend on declaration order.
 *
 * @param {any} value
 */
function sorted(value) {
  if (Array.isArray(value)) return value.map(sorted)
  if (value === null || typeof value !== 'object') return value
  return Object.keys(value)
    .sort()
    .reduce((result, key) => {
      result[key] = sorted(value[key])
      return result
    }, {})
}

/**
 * The schema typescript-json-schema makes for the root `Schema` interface, adjusted as described
 * above.
 */
function generate() {
  const program = TJS.getProgramFromFiles(
    [path.join(root, 'src', 'schema.ts')],
    { strictNullChecks: false, skipLibCheck: true, lib: ['es2016', 'dom'] },
    root
  )
  const schema = TJS.generateSchema(program, 'Schema', {
    ref: true,
    required: false,
    noExtraProps: false,
    titles: false,
    ignoreErrors: false
  })
  if (!schema) throw new Error('typescript-json-schema made no schema for Schema')

  const hooks = schema.definitions.IHooks
  hooks.properties.Headers = {
    additionalProperties: { type: 'string' },
    description: 'Request headers, by name.',
    type: 'object'
  }
  hooks.properties.Body = {
    description:
      'The request body as a JSON object; the handler adds its own properties and sends it as JSON.',
    type: 'object'
  }
  // Drop the definitions nothing refers to any more (the DOM types IHooks no longer uses), until
  // none is left; a $ref names its definition URI-encoded (Record%3Cstring%2Cstring%3E).
  let removed = true
  while (removed) {
    removed = false
    const referenced = new Set()
    const collect = (value) => {
      if (Array.isArray(value)) value.forEach(collect)
      else if (value && typeof value === 'object')
        for (const [key, child] of Object.entries(value))
          if (key === '$ref' && typeof child === 'string')
            referenced.add(decodeURIComponent(child.replace('#/definitions/', '')))
          else collect(child)
    }
    collect({ ...schema, definitions: undefined })
    for (const [name, definition] of Object.entries(schema.definitions))
      if (referenced.has(name)) collect(definition)
    for (const name of Object.keys(schema.definitions))
      if (!referenced.has(name)) {
        delete schema.definitions[name]
        removed = true
      }
  }
  return `${JSON.stringify(sorted(schema), null, 4)}\n`
}

const check = process.argv.includes('--check')
const generated = generate()
if (check) {
  const current = fs.existsSync(target) ? fs.readFileSync(target, 'utf8') : ''
  if (current !== generated) {
    console.error(
      'schema.json is not up to date with src/schema.ts: run `npm run build:schema` and commit it.'
    )
    process.exit(1)
  }
  console.log('schema.json is up to date with src/schema.ts.')
} else {
  fs.writeFileSync(target, generated)
  console.log('schema.json generated from src/schema.ts.')
}
