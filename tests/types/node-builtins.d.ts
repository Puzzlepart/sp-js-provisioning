/**
 * Minimal declarations for the Node builtins the unit tests use.
 *
 * The package is pinned to `@types/node@8`, which predates both `node:test` and
 * `node:assert/strict`. Declaring just the surface the tests touch keeps
 * `npm run typecheck:tests` working without dragging a TypeScript upgrade into
 * the PnPjs v4 migration.
 */
declare module 'node:test' {
  /** The context a test function gets (`it(name, (t) => ...)`). */
  interface TestContext {
    /** Marks the test skipped, with the reason. */
    skip(message?: string): void
    /** Marks the test as a known failure (`# TODO`), with the reason. */
    todo(message?: string): void
    /** Writes a line to the report (`# message`). */
    diagnostic(message: string): void
    readonly name: string
  }
  interface TestOptions {
    /** Skip the test or suite: `true`, or the reason. */
    skip?: boolean | string
    /** Run, but report a failure as a known one: `true`, or the reason. */
    todo?: boolean | string
    timeout?: number
    concurrency?: number | boolean
  }
  type TestFunction = (t: TestContext) => void | Promise<void>
  type SuiteFunction = () => void | Promise<void>
  type HookFunction = () => void | Promise<void>
  export function describe(name: string, fn: SuiteFunction): void
  export function describe(
    name: string,
    options: TestOptions,
    fn: SuiteFunction
  ): void
  export function it(name: string, fn: TestFunction): void
  export function it(name: string, options: TestOptions, fn: TestFunction): void
  export function before(fn: HookFunction, options?: { timeout?: number }): void
  export function after(fn: HookFunction, options?: { timeout?: number }): void
  export function beforeEach(
    fn: HookFunction,
    options?: { timeout?: number }
  ): void
  export function afterEach(
    fn: HookFunction,
    options?: { timeout?: number }
  ): void
}

declare module 'node:assert/strict' {
  interface Assert {
    (value: any, message?: string): void
    ok(value: any, message?: string): void
    equal(actual: any, expected: any, message?: string): void
    deepEqual(actual: any, expected: any, message?: string): void
    match(value: string, pattern: RegExp, message?: string): void
    throws(
      fn: () => any,
      error?: RegExp | Function | object,
      message?: string
    ): void
    fail(message?: string): never
    notEqual(actual: any, expected: any, message?: string): void
    notDeepEqual(actual: any, expected: any, message?: string): void
    deepStrictEqual(actual: any, expected: any, message?: string): void
    doesNotMatch(value: string, pattern: RegExp, message?: string): void
    rejects(
      promise: Promise<any> | (() => Promise<any>),
      error?: RegExp | Function | object,
      message?: string
    ): Promise<void>
  }
  const assert: Assert
  export default assert
}

declare module 'node:fs' {
  export function readFileSync(path: string, encoding: string): string
  export function mkdirSync(
    path: string,
    options?: { recursive?: boolean }
  ): string | undefined
  export function rmSync(
    path: string,
    options?: { recursive?: boolean; force?: boolean }
  ): void
}

declare module 'node:path' {
  export function join(...parts: string[]): string
}
