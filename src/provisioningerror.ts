/**
 * A readable message for whatever an operation threw or rejected with: a text as it is, an error's
 * `message` (or `message.value`, as in a SharePoint REST error), a JSOM request's
 * `args.get_message()` (`executeQueryAsync` rejects with `{ sender, args }`), else a short JSON of
 * the object, so a message never reads «[object Object]» or «undefined».
 *
 * @param error - What was thrown or rejected with
 * @returns The message, or `undefined` when there is nothing to read
 */
export function readableErrorMessage(error: any): string | undefined {
  if (error === undefined || error === null || error === '') return
  if (typeof error === 'string') return error
  if (typeof error !== 'object') return String(error)
  const message = error.message
  if (typeof message === 'string' && message) return message
  if (message && typeof message.value === 'string' && message.value)
    return message.value
  if (error.args && typeof error.args.get_message === 'function') {
    const jsomMessage = error.args.get_message()
    if (typeof jsomMessage === 'string' && jsomMessage) return jsomMessage
  }
  try {
    const json = JSON.stringify(error)
    if (json && json !== '{}')
      return json.length > 200 ? `${json.slice(0, 197)}...` : json
  } catch {
    // A cyclic object, such as a JSOM request's sender, has no JSON.
  }
  return String(error)
}

/**
 * Describes the Provisioning Error: the handler that failed, and what it threw or rejected with.
 */
export class ProvisioningError extends Error {
  /**
   * Creates a new instance of the ProvisioningError class
   *
   * @param handler - The handler that failed
   * @param cause - What the handler threw or rejected with: an `Error`, but possibly a string or
   * nothing at all
   */
  constructor(public handler: string, public cause?: any) {
    super(ProvisioningError.describe(handler, cause))
    this.name = 'ProvisioningError'
    // Keeps `instanceof ProvisioningError` working when the class is compiled to ES5.
    Object.setPrototypeOf(this, ProvisioningError.prototype)
  }

  /**
   * A readable message for whatever the handler failed with.
   *
   * @param handler - The handler that failed
   * @param cause - What it threw or rejected with
   */
  private static describe(handler: string, cause: any): string {
    return (
      readableErrorMessage(cause) ??
      `The ${handler} handler failed without an error`
    )
  }
}
