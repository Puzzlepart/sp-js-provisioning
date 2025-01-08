import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { IHooks } from '../schema'
import { HandlerBase } from './handlerbase'

/**
 * Describes the Hooks Object Handler
 */
export class Hooks extends HandlerBase {
  /**
   * Creates a new instance of the Hooks class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('Hooks', config)
  }

  /**
   * Provisioning Hooks
   *
   * @param hooks - The hook(s) to apply
   */
  public async ProvisionObjects(web: IWeb, hooks: IHooks[]): Promise<void> {
    super.scope_started()
    const promises = []
    const properties = await web.allProperties()

    // eslint-disable-next-line unicorn/no-array-for-each
    hooks.forEach((hook, index) => {
      if (hook.Method === 'GET') {
        super.log_info('processHooks', `Starting GET request: '${hook.Title}'.`)

        const getRequest = {
          method: 'GET',
          headers: hook.Headers || {}
        }

        promises.push(
          fetch(hook.Url, getRequest).then(async (response) => {
            const result = await Hooks.getJsonResult(response)

            if (!response.ok) {
              throw new Error(
                `${result ? ` | ${result} \n\n` : ''}- Hook ${index + 1}/${
                  hooks.length
                }: ${hook.Title}`
              )
            }
          })
        )
      } else if (hook.Method === 'POST') {
        super.log_info(
          'processHooks',
          `Starting POST request: '${hook.Title}'.`
        )

        const hookBody = { ...hook.Body, ...properties }

        const postRequest = {
          method: 'POST',
          body: JSON.stringify(hookBody) || '',
          headers: hook.Headers || {}
        }

        promises.push(
          fetch(hook.Url, postRequest).then(async (response) => {
            if (!response.ok) {
              const result = await Hooks.getJsonResult(response)
              throw new Error(
                `${result ? ` | ${result} \n\n` : ''}- Hook ${index + 1}/${
                  hooks.length
                }: ${hook.Title}`
              )
            } else if (response.status === 202) {
              const getPendingRequest = {
                method: 'GET',
                headers: hook.Headers || {}
              }

              const getPendingResult = (url: string): Promise<any> => {
                return new Promise((resolvePending, reject) => {
                  setTimeout(async () => {
                    await fetch(url, getPendingRequest).then(
                      async (response) => {
                        if (!response.ok) {
                          const result = await Hooks.getJsonResult(response)
                          reject(
                            new Error(
                              `${result ? ` | ${result} \n\n` : ''}- Hook ${
                                index + 1
                              }/${hooks.length}: ${hook.Title}`
                            )
                          )
                        } else if (response.status === 202) {
                          resolvePending(getPendingResult(url))
                        }
                      }
                    )
                  }, 5000)
                }).catch((error) => {
                  throw error
                })
              }

              const pendingResultLocation = response.headers.get('location')
              await getPendingResult(pendingResultLocation)
            }
          })
        )
      } else {
        super.log_info(
          'processHooks',
          `Method: '${hook.Method}' not supported.`
        )
      }
    })

    try {
      await Promise.all(promises)
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  public static getJsonResult(response: any): Promise<any> {
    return new Promise(async (resolve) => {
      if (!response.ok) {
        try {
          const jsonResponse = await response.json()
          resolve(
            `${response.status}${
              response.statusText ? ` - ${response.statusText}` : ''
            }${jsonResponse['error'] ? ` | ${jsonResponse['error']}` : ''}`
          )
        } catch {}
      }
      resolve(
        `${response.status}${
          response.statusText ? ` - ${response.statusText}` : ''
        }`
      )
    })
  }
}
