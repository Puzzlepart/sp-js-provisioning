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
    if (typeof cause === 'string' && cause) return cause
    if (cause && typeof cause.message === 'string' && cause.message)
      return cause.message
    if (cause === undefined || cause === null || cause === '')
      return `The ${handler} handler failed without an error`
    return String(cause)
  }
}
