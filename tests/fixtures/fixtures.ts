import * as fs from 'fs'
import * as path from 'path'

/** The folder with the fixture templates. */
export const TEMPLATES = path.join(__dirname, 'templates')

/**
 * Values for the two test-only placeholders, written where the engine replaces no tokens: `@@user@@`
 * (a User value or a group member) and `@@hookUrl@@` (a Hooks `Url`).
 */
export interface IFixtureValues {
  user?: string
  hookUrl?: string
}

/**
 * A fresh copy of a fixture template, `tests/fixtures/templates/<name>.json`, with the test-only
 * placeholders filled in. Fresh on every call, because handlers mutate their input (WebSettings
 * writes values back, ContentTypes sorts in place). A placeholder without a value is an error, so a
 * fixture never reaches SharePoint half-filled.
 *
 * @param name - The fixture, e.g. `lists`
 * @param values - The placeholder values
 */
export function loadFixture(
  name: string,
  values: IFixtureValues = {}
): Record<string, any> {
  const fill = (value: any): any => {
    if (typeof value === 'string')
      return value.replace(/@@(user|hookUrl)@@/g, (match, key: string) => {
        const filled = (values as any)[key]
        if (filled === undefined)
          throw new Error(`Fixture ${name} needs a value for ${match}`)
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
  const parsed = JSON.parse(
    fs.readFileSync(path.join(TEMPLATES, `${name}.json`), 'utf8')
  )
  delete parsed.$schema
  return fill(parsed)
}
