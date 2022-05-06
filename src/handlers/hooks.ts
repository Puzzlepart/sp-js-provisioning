import { HandlerBase } from './handlerbase'
import { IHooks } from '../schema'
import { Web } from '@pnp/sp'
import { IProvisioningConfig } from '../provisioningconfig'

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
    super(Hooks.name, config)
  }

  /**
   * Provisioning Hooks
   *
   * @param hooks - The hook(s) to apply
   */
  public async ProvisionObjects(
    web: Web,
    hooks: IHooks[],
  ): Promise<void> {
    super.scope_started()
    const promises = []


    hooks.forEach(async (hook, index) => {
      if (hook.Method === 'GET') {
        super.log_info(
          'processHooks',
          `Starting GET request: '${hook.Title}'.`
        )

        const getRequest = {
          method: 'GET',
          headers: hook.Headers || {},
        }

        promises.push(fetch(hook.Url, getRequest).then(async (res) => {
          const result = await Hooks.getJsonResult(res)

          if (!res.ok) {
            throw new Error(`${(result ? ` | ${result} \n\n` : '')}- Hook ${index + 1}/${hooks.length}: ${hook.Title}`)
          }
        }))
      } else if (hook.Method === 'POST') {
        super.log_info(
          'processHooks',
          `Starting POST request: '${hook.Title}'.`
        )

        hook.Body['pp_webUrl'] = web['_parentUrl']

        const postRequest = {
          method: 'POST',
          body: JSON.stringify(hook.Body) || '',
          headers: hook.Headers || {},
        }

        promises.push(fetch(hook.Url, postRequest).then(async (res) => {
          if (!res.ok) {
            const result = await Hooks.getJsonResult(res)
            throw new Error(`${(result ? ` | ${result} \n\n` : '')}- Hook ${index + 1}/${hooks.length}: ${hook.Title}`)
          } else if (res.status === 202) {
            const getPendingRequest = {
              method: 'GET',
              headers: hook.Headers || {},
            }

            const getPendingResult = (url: string): Promise<any> => {
              return new Promise((resolvePending, reject) => {
                setTimeout(async () => {
                  await fetch(url, getPendingRequest).then(async (res) => {
                    if (!res.ok) {
                      const result = await Hooks.getJsonResult(res)
                      reject(new Error(`${(result ? ` | ${result} \n\n` : '')}- Hook ${index + 1}/${hooks.length}: ${hook.Title}`))
                    } else if (res.status == 202) {
                      resolvePending(getPendingResult(url))
                    }
                  })
                }, (5000))
              }).catch((error) => {
                throw error
              })
            }

            const pendingResultLocation = res.headers.get('location')
            await getPendingResult(pendingResultLocation)
          }
        }))
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

  public static async getJsonResult(res: any): Promise<any> {
    return new Promise(async (resolve) => {
      if (!res.ok) {
        try {
          const jsonResponse = await res.json()
          resolve(`${res.status}${res.statusText ? ` - ${res.statusText}` : ''}${(jsonResponse['error'] ? ` | ${jsonResponse['error']}` : '')}`)
        } catch {}
      }
      resolve(`${res.status}${res.statusText ? ` - ${res.statusText}` : ''}`)
    })
  }
}
