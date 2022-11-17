/**
 * Describes the Provisioning Error
 */
export class ProvisioningError extends Error {
  constructor(public handler: string, error: Error) {
    super(error.message)
  }
}
