import { FunctionListener, ILogEntry, Logger, LogLevel } from '@pnp/logging'

export { LogLevel }

/**
 * A log line from a provisioning run, attributed to the handler that wrote it.
 */
export interface ILogRecord {
  /** The handler (`Lists`) or `WebProvisioner`, from the line's `(Name): (scope):` prefix. */
  source?: string
  level: LogLevel
  message: string
  data?: any
}

/**
 * The handler a log line comes from. Handlers write `<prefix> (<Handler>): (<scope>): <message>`,
 * the way PP365's PackageInstaller groups its install log.
 *
 * @param message - The log line
 */
export function sourceOf(message: string): string | undefined {
  const match = /\(([A-Za-z]+)\): \([^)]*\):/.exec(message)
  return match ? match[1] : undefined
}

/**
 * The log lines of a run.
 */
export interface ILogCollector {
  entries: ILogRecord[]
  errors(): ILogRecord[]
  warnings(): ILogRecord[]
  clear(): void
}

/**
 * Collects every line the engine logs (at the logger's active level) for assertions: `--strict`
 * fails on an error a handler logged and carried on after, not only on one it threw.
 */
export function collectLogs(): ILogCollector {
  const entries: ILogRecord[] = []
  Logger.subscribe(
    FunctionListener((entry: ILogEntry) => {
      entries.push({
        source: sourceOf(entry.message),
        level: entry.level,
        message: entry.message,
        data: entry.data
      })
    })
  )
  return {
    entries,
    errors: () => entries.filter((entry) => entry.level >= LogLevel.Error),
    warnings: () => entries.filter((entry) => entry.level === LogLevel.Warning),
    clear: () => {
      entries.length = 0
    }
  }
}

/** The collector of this process, subscribed on the first run. */
let collector: ILogCollector | undefined

/**
 * Starts a provisioning run's logging: one collector per process, cleared for each run. Listeners
 * are not removed, since the engine subscribes its console listener once per process and a
 * cleared one would not come back for the next run.
 *
 * Use this module's `Logger`, not one from `await import('@pnp/logging')`: a dynamic import of the
 * package loads its ES module build, a different `Logger` from the one the engine's static imports
 * share.
 */
export function startRun(): ILogCollector {
  collector ??= collectLogs()
  collector.clear()
  return collector
}
