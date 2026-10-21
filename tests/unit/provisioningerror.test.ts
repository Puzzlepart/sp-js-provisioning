import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as ts from 'typescript'
import { ProvisioningError } from '../../src/provisioningerror'

/**
 * The class as `npm run build` emits it (ES5, CommonJS), evaluated: subclasses
 * of `Error` lose their prototype there unless the constructor restores it.
 */
function compiledToEs5(): typeof ProvisioningError {
  const source = readFileSync(
    join(__dirname, '..', '..', 'src', 'provisioningerror.ts'),
    'utf8'
  )
  const { outputText } = ts.transpileModule(source, {
    compilerOptions: {
      target: ts.ScriptTarget.ES5,
      module: ts.ModuleKind.CommonJS
    }
  })
  const exports: any = {}
  // eslint-disable-next-line no-new-func
  new Function('exports', outputText)(exports)
  return exports.ProvisioningError
}

describe('ProvisioningError', () => {
  it('reads the message of an Error and keeps the handler and the cause', () => {
    const cause = Object.assign(new Error('Access denied'), { status: 403 })
    const error = new ProvisioningError('Lists', cause)

    assert.equal(error.message, 'Access denied')
    assert.equal(error.handler, 'Lists')
    assert.equal(error.cause, cause)
    assert.equal(error.name, 'ProvisioningError')
  })

  it('gives a readable message for a string, an object or nothing at all', () => {
    // Files threw a string; PropertyBagEntries rejected with undefined, which
    // made the constructor itself throw a TypeError and lose the handler.
    assert.equal(new ProvisioningError('Files', 'Not here').message, 'Not here')
    assert.equal(
      new ProvisioningError('PropertyBagEntries', undefined).message,
      'The PropertyBagEntries handler failed without an error'
    )
    assert.equal(
      new ProvisioningError('Hooks', null).message,
      'The Hooks handler failed without an error'
    )
    assert.equal(
      new ProvisioningError('Hooks', { message: 'From an object' }).message,
      'From an object'
    )
    assert.equal(new ProvisioningError('Hooks', 42).message, '42')
  })

  it('stays an instance of itself when compiled to ES5', () => {
    const Es5ProvisioningError = compiledToEs5()
    const error = new Es5ProvisioningError('Files', 'Not here')

    assert.ok(error instanceof Es5ProvisioningError)
    assert.ok(error instanceof Error)
    assert.equal(error.handler, 'Files')
    assert.equal(error.message, 'Not here')
  })
})
