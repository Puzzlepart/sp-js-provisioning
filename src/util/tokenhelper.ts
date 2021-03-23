import { ProvisioningContext } from '../provisioningcontext'
import { IProvisioningConfig } from '../provisioningconfig'

/**
 * Describes the Token Helper
 */
export class TokenHelper {
  private tokenRegex = /{[a-z]*:[ A-z|ÅÆØåæø]*}/g

  /**
   * Creates a new instance of the TokenHelper class
   */
  constructor(
    public context: ProvisioningContext,
    public config: IProvisioningConfig
  ) {}

  public replaceTokens(string: string) {
    let m: RegExpExecArray
    while ((m = this.tokenRegex.exec(string)) !== null) {
      if (m.index === this.tokenRegex.lastIndex) {
        this.tokenRegex.lastIndex++
      }
      for (const match of m) {
        const [tokenKey, tokenValue] = match.replace(/[{}]/g, '').split(':')
        switch (tokenKey) {
          case 'listid':
            {
              if (this.context.lists[tokenValue]) {
                string = string.replace(match, this.context.lists[tokenValue])
              }
            }
            break
          case 'listviewid':
            {
              if (this.context.listViews[tokenValue]) {
                string = string.replace(
                  match,
                  this.context.listViews[tokenValue]
                )
              }
            }
            break
          case 'webid':
            {
              if (this.context.web.Id) {
                string = string.replace(match, this.context.web.Id)
              }
            }
            break
          case 'siteid':
            {
              if (this.context.web.Id) {
                string = string.replace(match, this.context.web.Id)
              }
            }
            break
          case 'sitecollectionid':
            {
              if (this.context.web.Id) {
                string = string.replace(match, this.context.web.Id)
              }
            }
            break
          case 'parameter':
            {
              if (this.config.parameters) {
                const parameterValue = this.config.parameters[tokenValue]
                if (parameterValue) {
                  string = string.replace(match, parameterValue)
                }
              }
            }
            break
        }
      }
    }
    return string
  }
}
