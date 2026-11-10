import { ISiteConnection } from './connect'
import { env } from './env'

/**
 * The test fixtures' two run-time placeholders (tests/fixtures/README.md), written where the engine
 * replaces no tokens: `@@user@@` (a User value or a group member) and `@@hookUrl@@` (a Hooks `Url`,
 * a listener only the Hooks suite starts). No real template has them.
 */
export const PLACEHOLDERS = ['@@user@@', '@@hookUrl@@']

/**
 * The placeholders a template holds.
 *
 * @param schema - The template
 */
export function placeholdersIn(schema: Record<string, any>): string[] {
  const text = JSON.stringify(schema)
  return PLACEHOLDERS.filter((placeholder) => text.includes(placeholder))
}

/**
 * A copy of a template with the placeholders filled in. A placeholder without a value is an error,
 * so a template never reaches SharePoint half-filled.
 *
 * @param schema - The template
 * @param values - The values, by placeholder name (`user`, `hookUrl`)
 * @param name - The template's name, for the error
 */
export function fillPlaceholders(
  schema: Record<string, any>,
  values: { user?: string; hookUrl?: string },
  name = 'The template'
): Record<string, any> {
  const fill = (value: any): any => {
    if (typeof value === 'string')
      return value.replace(/@@(user|hookUrl)@@/g, (match, key: string) => {
        const filled = (values as any)[key]
        if (filled === undefined)
          throw new Error(`${name} needs a value for ${match}`)
        return filled
      })
    if (Array.isArray(value)) return value.map(fill)
    if (value && typeof value === 'object') {
      const copy: Record<string, any> = {}
      for (const key of Object.keys(value)) copy[key] = fill(value[key])
      return copy
    }
    return value
  }
  return fill(schema)
}

/**
 * The user `@@user@@` stands for, as SharePoint stores its login: the signed-in user with cookies,
 * or `INTEGRATION_USER` as the web resolves it with the certificate (an app has no user).
 * `undefined` when there is none.
 *
 * @param conn - The web
 */
export async function fixtureUser(
  conn: ISiteConnection
): Promise<string | undefined> {
  if (conn.mode === 'cookies')
    return (
      await conn.web.currentUser.select('LoginName')<{ LoginName: string }>()
    ).LoginName
  return env.user && (await conn.web.ensureUser(env.user)).LoginName
}
