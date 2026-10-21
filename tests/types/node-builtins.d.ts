/**
 * Minimal declarations for the Node builtins the unit tests use.
 *
 * The package is pinned to `@types/node@8`, which predates both `node:test` and
 * `node:assert/strict`. Declaring just the surface the tests touch keeps
 * `npm run typecheck:tests` working without dragging a TypeScript upgrade into
 * the PnPjs v4 migration.
 */
declare module 'node:test' {
  type TestFunction = () => void | Promise<void>
  export function describe(name: string, fn: TestFunction): void
  export function it(name: string, fn: TestFunction): void
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
  export function rmSync(
    path: string,
    options?: { recursive?: boolean; force?: boolean }
  ): void
}

declare module 'node:path' {
  export function join(...parts: string[]): string
}
