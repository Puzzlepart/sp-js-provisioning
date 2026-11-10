import assert from 'node:assert/strict'
import { describe, it } from 'node:test'
import { readFileSync } from 'node:fs'
import { join } from 'node:path'
import * as ts from 'typescript'
import { FunctionListener, Logger } from '@pnp/logging'
import { HandlerBase } from '../../src/handlers/handlerbase'
import {
  ProvisioningError,
  readableErrorMessage
} from '../../src/provisioningerror'

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

  it("reads SharePoint's JSOM and REST errors, and never says [object Object]", () => {
    // ContentTypes rejects with JSOM's { sender, args }, whose sender is cyclic: the message was
    // «[object Object]».
    const sender: any = {}
    sender.self = sender
    const jsom = {
      sender,
      args: { get_message: () => 'Innholdstypen finnes allerede.' }
    }
    assert.equal(
      new ProvisioningError('ContentTypes', jsom).message,
      'Innholdstypen finnes allerede.'
    )
    assert.equal(
      new ProvisioningError('Lists', {
        message: { lang: 'nb-NO', value: 'Ingen tilgang' }
      }).message,
      'Ingen tilgang'
    )
    assert.equal(
      new ProvisioningError('Hooks', { code: 'E1' }).message,
      '{"code":"E1"}'
    )
    assert.equal(readableErrorMessage(undefined), undefined)
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

describe('HandlerBase.scope_ended', () => {
  it('logs what a JSOM request rejected with, not undefined', () => {
    // «Code execution scope ended with error: undefined» was all the setup log said when the
    // ContentTypes handler's JSOM request was refused.
    const messages: string[] = []
    Logger.subscribe(FunctionListener((entry) => messages.push(entry.message)))
    try {
      new HandlerBase('ContentTypes').scope_ended({
        sender: {},
        args: { get_message: () => 'Innholdstypen finnes allerede.' }
      })
      assert.deepEqual(messages, [
        '(ContentTypes): (ProvisionObjects): Code execution scope ended with error: Innholdstypen finnes allerede.'
      ])
    } finally {
      Logger.clearSubscribers()
    }
  })
})
