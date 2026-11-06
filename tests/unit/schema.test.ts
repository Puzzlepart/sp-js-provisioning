import assert from 'node:assert/strict'
import { existsSync, readdirSync, readFileSync, statSync } from 'node:fs'
import { join, relative, resolve } from 'node:path'
import { describe, it } from 'node:test'
// eslint-disable-next-line @typescript-eslint/no-var-requires
const Ajv = require('ajv')

const root = join(__dirname, '..', '..')
const schema = JSON.parse(readFileSync(join(root, 'schema.json'), 'utf8'))
const fixtures = join(root, 'tests', 'fixtures', 'templates')

/**
 * The hosting repository's packages: INTEGRATION_HOSTING_DIR, else a prosjektportalen-hosting
 * checkout next to this one. Undefined when neither exists.
 */
function hostingPackages(): string | undefined {
  const dir = process.env.INTEGRATION_HOSTING_DIR
    ? resolve(process.env.INTEGRATION_HOSTING_DIR)
    : resolve(root, '..', 'prosjektportalen-hosting')
  const packages = join(dir, 'packages')
  return existsSync(packages) ? packages : undefined
}

/**
 * Every .json file below a folder.
 *
 * @param dir - The folder
 */
function jsonFiles(dir: string): string[] {
  return readdirSync(dir).reduce((files: string[], name) => {
    const path = join(dir, name)
    if (statSync(path).isDirectory()) return files.concat(jsonFiles(path))
    return name.endsWith('.json') ? files.concat(path) : files
  }, [])
}

/**
 * The template JSON files of the hosting packages (their `provisioning` folders).
 *
 * @param packages - The hosting repository's packages folder
 */
function hostingTemplates(packages: string): string[] {
  return readdirSync(packages)
    .map((name) => join(packages, name, 'provisioning'))
    .filter((dir) => existsSync(dir))
    .reduce((files: string[], dir) => files.concat(jsonFiles(dir)), [])
}

const read = (path: string): any =>
  JSON.parse(readFileSync(path, 'utf8').replace(/^﻿/, ''))

describe('schema.json', () => {
  const validate = new Ajv({ allErrors: true }).compile(schema)
  const errors = () =>
    (validate.errors || [])
      .map((error: any) => `${error.dataPath} ${error.message}`)
      .join('; ')

  it('accepts every fixture template', () => {
    const files = readdirSync(fixtures).filter((name) => name.endsWith('.json'))
    assert.ok(files.length > 0)
    for (const name of files)
      assert.ok(validate(read(join(fixtures, name))), `${name}: ${errors()}`)
  })

  const packages = hostingPackages()
  it(
    "accepts every hosting package's templates",
    {
      skip: packages
        ? false
        : 'no prosjektportalen-hosting checkout (set INTEGRATION_HOSTING_DIR)'
    },
    () => {
      const files = hostingTemplates(packages)
      assert.ok(files.length > 0)
      for (const path of files)
        assert.ok(
          validate(read(path)),
          `${relative(packages, path)}: ${errors()}`
        )
    }
  )

  it('rejects a template whose sections have the wrong shape', () => {
    assert.equal(validate({ Lists: { Title: 'Not an array' } }), false)
    assert.equal(validate({ SiteFields: [{ Name: 'Not XML' }] }), false)
    assert.equal(
      validate({ Navigation: { QuickLaunch: [{ Title: 1 }] } }),
      false
    )
  })

  it('accepts a section the engine has no handler for, as the engine skips it', () => {
    assert.ok(
      validate({ Version: '1.0', SomethingElse: { any: 'thing' } }),
      errors()
    )
  })

  it('describes the sections whose meaning changed in 1.5.0', () => {
    assert.match(
      schema.properties.SiteFields.description,
      /merged onto its stored schema/
    )
    const list = schema.definitions.IListInstance.properties
    assert.ok(list.RemoveExistingViews.description)
    assert.ok(list.RemoveExistingContentTypes.description)
    assert.ok(
      schema.definitions.INavigationNode.properties.IgnoreExisting.description
    )
  })

  it('types the hooks without DOM types', () => {
    const hooks = schema.definitions.IHooks.properties
    assert.deepEqual(hooks.Headers.additionalProperties, { type: 'string' })
    assert.equal(hooks.Body.type, 'object')
    for (const name of Object.keys(schema.definitions))
      assert.doesNotMatch(name, /^(Body|Headers|ReadableStream)/)
  })
})
