import { ProvisioningContext } from '../provisioningcontext'
import { IProvisioningConfig } from '../provisioningconfig'

/** The keys `resolveToken` knows: `{key}` with another key is not a token (`{Title}`, JSON). */
const KNOWN_KEYS = new Set([
  'listid',
  'listviewid',
  'webid',
  'siteid',
  'sitename',
  'sitecollectionid',
  'sitecollectiontermstoreid',
  'termstoreid',
  'parameter'
])

/**
 * Describes the Token Helper
 *
 * Replaces `{key}` and `{key:value}` tokens in strings (field XML, data row
 * values, file sources, web part properties). Unknown keys are left untouched,
 * so JSON-like text such as `{"a":1}` or column placeholders such as `{Title}`
 * (uppercase keys never match) are never damaged.
 *
 * Supported tokens:
 *
 * - `{listid:List Title}` - list id from `context.lists`
 * - `{listviewid:List Title|View Title}` - view id from `context.listViews`
 * - `{webid}` / `{siteid}` - id of the web being provisioned
 * - `{sitename}` - title of the web being provisioned
 * - `{sitecollectionid}` - id of the site collection (falls back to the web id)
 * - `{sitecollectiontermstoreid}` / `{termstoreid}` - id of the default term store
 * - `{parameter:Name}` - value from `config.parameters`
 *
 * `unresolvedTokens` names the known tokens a string still holds, so a handler can say which value
 * is missing instead of writing the token silently.
 */
export class TokenHelper {
  /**
   * Matches `{key}` and `{key:value}`: the key is lowercase letters, the value
   * (optional) is anything but braces, so list titles with spaces, `|` and
   * Norwegian letters are supported.
   */
  private tokenRegex = /{([a-z]+)(?::([^{}]*))?}/g

  /**
   * Creates a new instance of the TokenHelper class
   */
  constructor(
    public context: ProvisioningContext,
    public config: IProvisioningConfig
  ) {}

  /**
   * Replaces all resolvable tokens in the string. Tokens that cannot be
   * resolved (unknown key, or no value in the context/config) are kept as-is.
   *
   * @param string - The string to replace tokens in
   */
  public replaceTokens(string: string): string {
    if (typeof string !== 'string' || !string.includes('{')) return string
    return string.replace(
      this.tokenRegex,
      (match: string, tokenKey: string, tokenValue: string) => {
        const replacement = this.resolveToken(tokenKey, tokenValue)
        return replacement === undefined || replacement === null
          ? match
          : replacement
      }
    )
  }

  /**
   * The tokens in the string that `replaceTokens` knows but cannot resolve now (no list with that
   * title yet, no term store id, no such parameter, ...), each once, as written: `{listid:Kilder}`.
   * They are left in place by `replaceTokens`; a key it does not know is not a token and is not
   * reported.
   *
   * @param string - The string, before or after `replaceTokens` (which leaves only these)
   */
  public unresolvedTokens(string: string): string[] {
    if (typeof string !== 'string' || !string.includes('{')) return []
    const tokens: string[] = []
    const regex = new RegExp(this.tokenRegex.source, 'g')
    let match: RegExpExecArray | null
    while ((match = regex.exec(string)) !== null) {
      const [token, tokenKey, tokenValue] = match
      if (!KNOWN_KEYS.has(tokenKey) || tokens.includes(token)) continue
      const value = this.resolveToken(tokenKey, tokenValue)
      if (value === undefined || value === null) tokens.push(token)
    }
    return tokens
  }

  /**
   * Resolves a single token to its value, or `undefined` when it cannot be
   * resolved.
   *
   * @param tokenKey - Token key (e.g. `listid`)
   * @param tokenValue - Token value (e.g. the list title), if any
   */
  private resolveToken(
    tokenKey: string,
    tokenValue?: string
  ): string | undefined {
    switch (tokenKey) {
      case 'listid':
        return this.context.lists[tokenValue]
      case 'listviewid':
        return this.context.listViews[tokenValue]
      case 'webid':
      case 'siteid':
        return this.context.web && this.context.web.Id
      case 'sitename':
        return this.context.web && this.context.web.Title
      case 'sitecollectionid':
        return this.context.siteId || (this.context.web && this.context.web.Id)
      case 'sitecollectiontermstoreid':
      case 'termstoreid':
        return this.context.termStoreId
      case 'parameter':
        return this.config.parameters
          ? this.config.parameters[tokenValue]
          : undefined
    }
  }
}
