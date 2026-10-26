/**
 * Text that may reach a terminal, a test report or a CI artifact, without secrets or the tenant:
 * session cookie values (`FedAuth`, `rtFa`, any `Set-Cookie`), bearer tokens, and SharePoint hosts
 * other than the example ones.
 *
 * @param text - The text
 */
export function redact(text: string): string {
  return String(text)
    .replace(/\b(FedAuth|rtFa|SIMI|EdgeAccessCookie)=[^;\s"',]*/gi, '$1=<redacted>')
    .replace(/(set-cookie['"]?\s*[:=]\s*['"]?)[^'"\n]*/gi, '$1<redacted>')
    .replace(/Bearer\s+[A-Za-z0-9._~+/=-]+/g, 'Bearer <redacted>')
    .replace(
      /\b([a-z0-9]+(?:-(?!(?:admin|my)\.sharepoint\.com)[a-z0-9]+)*)((?:-admin|-my)?)\.sharepoint\.com/gi,
      (match, host, suffix) =>
      /^(contoso|fabrikam)$/i.test(host)
        ? match
        : `<tenant>${suffix}.sharepoint.com`
    )
}

/**
 * The message of an error, redacted: never the error object itself, whose `cause.response` holds
 * the response headers (with any `Set-Cookie`).
 *
 * @param error - The error
 */
export function describeError(error: any): string {
  if (error === undefined || error === null) return String(error)
  const status = error.status ?? (error.cause && error.cause.status)
  const message =
    typeof error === 'string' ? error : error.message || String(error)
  return redact(status && !message.includes(`[${status}]`) ? `${message} (${status})` : message)
}
