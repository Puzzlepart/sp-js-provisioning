import { IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { IHooks } from '../schema'
import { HandlerBase } from './handlerbase'

/**
 * Describes the Hooks Object Handler
 */
export class Hooks extends HandlerBase {
  /**
   * How long to wait before each poll of a `202 Accepted` hook's `Location`, in milliseconds
   */
  protected pollInterval = 5000

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
      // A request that cannot be sent (network, CORS) fails the hook with its name, as an error
      // status does: with hooks running together, a bare "Failed to fetch" names none.
      const send = (url: string, init: RequestInit) =>
        fetch(url, init).catch((error) => {
          throw new Error(
            `${error && error.message ? error.message : error} - Hook ${
              index + 1
            }/${hooks.length}: ${hook.Title}`
          )
        })
      if (hook.Method === 'GET') {
        super.log_info('processHooks', `Starting GET request: '${hook.Title}'.`)

        const getRequest = {
          method: 'GET',
          headers: hook.Headers || {}
        }

        promises.push(
          send(hook.Url, getRequest).then(async (response) => {
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
          send(hook.Url, postRequest).then(async (response) => {
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

              // Polls the Location until it answers something other than 202: another 2xx
              // ends the hook, an error status or a failed request fails it.
              const getPendingResult = async (url: string): Promise<void> => {
                await new Promise((resolveWait) =>
                  setTimeout(resolveWait, this.pollInterval)
                )
                const pendingResponse = await send(url, getPendingRequest)
                if (!pendingResponse.ok) {
                  const result = await Hooks.getJsonResult(pendingResponse)
                  throw new Error(
                    `${result ? ` | ${result} \n\n` : ''}- Hook ${index + 1}/${
                      hooks.length
                    }: ${hook.Title}`
                  )
                }
                if (pendingResponse.status === 202) return getPendingResult(url)
              }

              // The Location is resolved against the URL that answered (after any redirect; the
              // hook's own URL where a response has none), as a browser resolves a relative one.
              // Without one there is nothing to poll.
              const location = response.headers.get('location')
              if (!location)
                throw new Error(
                  `202 Accepted without a Location to poll - Hook ${
                    index + 1
                  }/${hooks.length}: ${hook.Title}`
                )
              let pollUrl: string
              try {
                pollUrl = new URL(location, response.url || hook.Url).toString()
              } catch {
                throw new Error(
                  `Invalid Location '${location}' to poll - Hook ${index + 1}/${
                    hooks.length
                  }: ${hook.Title}`
                )
              }
              await getPendingResult(pollUrl)
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
