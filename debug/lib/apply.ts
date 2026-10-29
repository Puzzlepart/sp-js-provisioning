import { Logger, LogLevel } from '@pnp/logging'
import { DefaultHandlerSort } from '../../src/handlers/exports'
import { WebProvisioner } from '../../src/webprovisioner'
import { handlerOrder, planHandlers } from './cli'
import { ISiteConnection } from './connect'
import { ILogRecord, startRun } from './logs'

/**
 * What one template run did: the handlers it ran (from the progress callback), the ones it skipped
 * because they cannot run in Node, the errors and warnings the handlers logged, and the
 * `ProvisioningError` it stopped with, if any. `apply` never throws for a failed run, so callers
 * (the runner, the tests) can assert on it.
 */
export interface IApplyResult {
  order: string[]
  skipped: string[]
  /** Every line the run logged, at Info and above. */
  entries: ILogRecord[]
  errors: ILogRecord[]
  warnings: ILogRecord[]
  error?: any
}

export interface IApplyOptions {
  /** Only these handlers (validated against the engine; also forces skipped ones). */
  handlers?: string[]
  /** Values for `{parameter:...}` tokens. */
  parameters?: Record<string, string>
  /** Print the engine's log lines with this prefix; without it the run is quiet. */
  prefix?: string
  /** Called as each handler starts. */
  onHandler?: (handler: string) => void
}

/**
 * Applies a template to a site the way PP365's PackageInstaller does
 * (`WebProvisioner(web).setup({ spfxContext, ... }).applyTemplate(schema)`), with the SPFx
 * context stand-in of the connection.
 *
 * @param conn - The site
 * @param schema - The template
 * @param options - Handlers, parameters, logging
 */
export async function apply(
  conn: ISiteConnection,
  schema: Record<string, any>,
  options: IApplyOptions = {}
): Promise<IApplyResult> {
  const plan = planHandlers(
    schema,
    handlerOrder(DefaultHandlerSort),
    options.handlers
  )
  const logs = startRun()
  // Collect Info and above; without a prefix there is no console listener, so the run is quiet.
  if (!options.prefix) Logger.activeLogLevel = LogLevel.Info
  const provisioner = new WebProvisioner(conn.web).setup({
    spfxContext: conn.context,
    parameters: options.parameters ?? {},
    // Without `logging` the engine subscribes no console listener: a quiet run.
    ...(options.prefix
      ? { logging: { prefix: options.prefix, activeLogLevel: LogLevel.Info } }
      : {})
  })
  const order: string[] = []
  let error: any
  try {
    await provisioner.applyTemplate(schema, plan.run, (handler) => {
      order.push(handler)
      if (options.onHandler) options.onHandler(handler)
    })
  } catch (caught) {
    error = caught
  }
  return {
    order,
    skipped: plan.skipped,
    entries: logs.entries.slice(),
    errors: logs.errors(),
    warnings: logs.warnings(),
    error
  }
}
