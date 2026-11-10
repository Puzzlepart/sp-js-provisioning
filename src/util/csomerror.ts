/**
 * An error SharePoint reported in the answer to a CSOM request.
 */
export class CsomError extends Error {
  /**
   * Creates a new instance of the CsomError class
   *
   * @param serverMessage - SharePoint's message, in the language of the web or the user
   * @param typeName - The error's type, such as `System.UnauthorizedAccessException`
   * @param code - SharePoint's error code
   */
  constructor(
    public serverMessage: string,
    public typeName: string,
    public code?: number
  ) {
    super(`CSOM request failed: ${serverMessage} (${typeName})`)
    this.name = 'CsomError'
    // Keeps `instanceof CsomError` working when the class is compiled to ES5.
    Object.setPrototypeOf(this, CsomError.prototype)
  }

  /**
   * Whether SharePoint refused the request for lack of permission. On a site that blocks custom
   * script that is also the answer to a property bag write, whoever asks.
   */
  public get accessDenied(): boolean {
    return /unauthorizedaccessexception|accessdenied/i.test(this.typeName || '')
  }
}
