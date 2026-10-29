import { FunctionListener, ILogEntry, Logger, LogLevel } from '@pnp/logging'

/**
 * Runs `fn` with every log line captured, at every level, then removes the listener and restores
 * the log level.
 *
 * @param fn - The code under test
 */
export async function captureLogs<T>(
  fn: () => Promise<T>
): Promise<{ result: T; entries: ILogEntry[] }> {
  const entries: ILogEntry[] = []
  const level = Logger.activeLogLevel
  Logger.activeLogLevel = LogLevel.Verbose
  Logger.subscribe(FunctionListener((entry) => entries.push(entry)))
  try {
    return { result: await fn(), entries }
  } finally {
    Logger.clearSubscribers()
    Logger.activeLogLevel = level
  }
}

/**
 * The level a line matching `pattern` was logged at, or `undefined` when none matched.
 *
 * @param entries - The captured lines
 * @param pattern - What the line must contain
 */
export function levelOf(
  entries: ILogEntry[],
  pattern: RegExp
): LogLevel | undefined {
  const entry = entries.find((candidate) => pattern.test(candidate.message))
  return entry && entry.level
}
