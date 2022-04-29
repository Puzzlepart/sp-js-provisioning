import { HandlerBase } from './handlerbase'
import { IHooks } from '../schema'
import { Web } from '@pnp/sp'
import { IProvisioningConfig } from '../provisioningconfig'
import { ProvisioningContext } from '../provisioningcontext'

/**
 * Describes the Hooks Object Handler
 */
export class Hooks extends HandlerBase {
  /**
   * Creates a new instance of the Hooks class
   *
   * @param config - Provisioning config
   * @param client - HttpClient to use
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
    try {
      console.log(hooks)

      hooks.forEach(async hook => {
        super.log_info(
          'processHooks',
          `Running hook: ${hook.Title}.`
        )

        console.log("Running hook", hook.Title, hook.Method)


        if (hook.Method === 'GET') {
          super.log_info(
            'processHooks',
            `Starting GET request: '${hook.Title}'.`
          )

          try {
            const getRequest = {
              method: 'GET',
              headers: hook.Headers || {},
            };

            const response: Response | any = await fetch(hook.Url, getRequest)
            const responseBody = await response.json()

            console.log("response.status", response.status)

            alert(JSON.stringify(responseBody))

            super.log_info(
              'processHooks',
              `Ran hook successfully.`
            )
          } catch (error) {
            console.log(error)
            super.log_info(
              'processHooks',
              `Failed to run hook: ${hook.Title}. | ${error.message}`
            )
          }
        } else if (hook.Method === 'POST') {
          super.log_info(
            'processHooks',
            `Starting POST request: '${hook.Title}'.`
          )

          try {
            const postRequest = {
              method: 'POST',
              body: JSON.stringify(hook.Body) || '',
              headers: hook.Headers || {},
            }

            const response: Response | any = await fetch(hook.Url, postRequest)
            console.log("response.status", response.status)

            super.log_info(
              'processHooks',
              `Ran hook, status: ${response.status}.`
            )

          } catch (error) {
            console.log(error)
            super.log_info(
              'processHooks',
              `Failed to run hook: ${hook.Title}. | ${error.message}`
            )
          }
        } else {
          super.log_info(
            'processHooks',
            `Method: '${hook.Method}' not supported.`
          )
        }

        // const request = {
        //   type: hook.method || 'GET',
        //   headers: hook.headers || {},
        // };

        // const response: Response | any = await fetch(hook.url, request);
        // const responseBody = await response.json();


        // super.log_info(
        //   'processHooks',
        //   `Ran hook, result: ${responseBody}.`
        // )
        // console.log(responseBody)

        // this.client.get(hook.url, HttpClient.configurations.v1)
        //   .then(async (res: HttpClientResponse): Promise<any> => {
        //     if (res.status !== 202) {
        //       console.log(res, res.status)
        //       return Hooks.getJsonResult(res);
        //     }
        //     console.log("res.status === 202", res)

        //     // let getPendingResult = (url): Promise<any> => {
        //     //   return new Promise((resolvePending) => {
        //     //     setTimeout(async () => {
        //     //       let result = await this.client.get(url, HttpClient.configurations.v1);
        //     //       if (result.status == 202) resolvePending(getPendingResult(url));
        //     //       else resolvePending(result);
        //     //     }, (5000));
        //     //   });
        //     // };

        //     // let pendingResultLocation = res.headers.get("location");
        //     // return Hooks.getJsonResult(await getPendingResult(pendingResultLocation));
        //   })
        //   .then((result: any): void => {
        //     console.log(result)
        //     super.log_info(
        //       'processHooks',
        //       `Ran hook, result: ${result}.`
        //     )
        //   })
      });

      super.log_info(
        'processHooks',
        `${hooks.length} hook(s) ran successfully.`
      )
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      super.log_info(
        'processHooks',
        `Failed to run hook(s): ${error.message}`
      )
      throw error
    }
  }

  // public static async getJsonResult(res: any): Promise<any> {
  //   return new Promise(async (resolve, reject) => {
  //     try {
  //       const jsonResponse = await res.json();
  //       if (res.ok) resolve(jsonResponse);
  //       else reject(`${res.status}${res.statusText ? ` - ${res.statusText}` : ``}${(jsonResponse["error"] ? ` | ${jsonResponse["error"]}` : ``)}`);
  //     } catch {
  //       if (!res.ok) reject(`${res.status} - ${res.statusText}`);
  //     }
  //   });
  // }
}
