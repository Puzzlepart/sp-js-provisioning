import { IProvisioningConfig } from '../provisioningconfig'
import * as xmljs from 'xml-js'

/** Node's `Buffer`, where there is one (the engine is built without Node typings). */
declare const Buffer: any

/**
 * The browser's `window`, or `undefined` outside a browser (Node).
 */
function browserWindow(): any {
  return typeof window === 'undefined' ? undefined : window
}

/**
 * The origin behind `{hosturl}`: the SPFx context's web, else the page's location; `undefined`
 * when there is neither.
 *
 * @param config - Provisioning config
 */
function hostUrl(config: IProvisioningConfig): string | undefined {
  const absoluteUrl =
    config.spfxContext &&
    config.spfxContext.pageContext &&
    config.spfxContext.pageContext.web &&
    config.spfxContext.pageContext.web.absoluteUrl
  if (absoluteUrl) return new URL(absoluteUrl).origin
  const location = browserWindow() && browserWindow().location
  return location ? `${location.protocol}//${location.host}` : undefined
}

/**
 * Replaces the URL tokens: `{site}` (server-relative) and `{sitecollection}` (absolute) from the
 * SPFx context's site collection, else from `_spPageContextInfo` on a classic page; `{wpgallery}`,
 * `{themegallery}`; `{hosturl}` (the origin). Without a site, `{site}` and `{sitecollection}`
 * become `null`; without an origin, `{hosturl}` is left as is.
 *
 * @param string - The string to replace tokens in
 * @param config - Provisioning config
 */
export function replaceUrlTokens(
  string: string,
  config: IProvisioningConfig
): string {
  let siteAbsoluteUrl = null
  let siteRelativeUrl = null
  const win = browserWindow()
  if (config.spfxContext) {
    siteAbsoluteUrl = config.spfxContext.pageContext.site.absoluteUrl
    siteRelativeUrl = config.spfxContext.pageContext.site.serverRelativeUrl
  } else if (win && win.hasOwnProperty('_spPageContextInfo')) {
    siteAbsoluteUrl = win._spPageContextInfo.siteAbsoluteUrl
    siteRelativeUrl = win._spPageContextInfo.siteServerRelativeUrl
  }
  const host = hostUrl(config)
  const replaced = string
    .replace(/{site}/g, siteRelativeUrl)
    .replace(/{sitecollection}/g, siteAbsoluteUrl)
    .replace(/{wpgallery}/g, `${siteAbsoluteUrl}/_catalogs/wp`)
    .replace(/{themegallery}/g, `${siteAbsoluteUrl}/_catalogs/theme/15`)
  return host === undefined ? replaced : replaced.replace(/{hosturl}/g, host)
}

/**
 * A deep copy of a JSON value with `replace` applied to every string in it, keys included. For
 * token replacement in structured values: replacing in their JSON text broke on a value with a
 * quote or a backslash (`{sitename}` of a web titled `Rehab "Gamle skole"`).
 *
 * @param value - The value
 * @param replace - What to do with each string
 */
export function mapStrings<T>(value: T, replace: (text: string) => string): T {
  if (typeof value === 'string') return replace(value) as any
  if (Array.isArray(value))
    return value.map((item) => mapStrings(item, replace)) as any
  if (value && typeof value === 'object') {
    const copy: Record<string, any> = {}
    for (const key of Object.keys(value))
      copy[replace(key)] = mapStrings((value as any)[key], replace)
    return copy as any
  }
  return value
}

/**
 * Removes the origin from an absolute URL: the page's in the browser, any origin in Node.
 *
 * @param absUrl - The absolute URL
 */
export function makeUrlRelative(absUrl: string): string {
  if (typeof document !== 'undefined' && document.location)
    return absUrl.replace(
      `${document.location.protocol}//${document.location.hostname}`,
      ''
    )
  return absUrl.replace(/^https?:\/\/[^/]+/i, '')
}

/**
 * Base64 of a string's UTF-16LE bytes (how SharePoint encodes indexed property bag keys).
 *
 * @param string - The string
 */
export function base64EncodeString(string: string): string {
  const bytes = []
  for (let index = 0; index < string.length; ++index) {
    const code = string.charCodeAt(index)
    // Low byte, then high byte: a character above U+00FF used to go to btoa whole and throw.
    bytes.push(code & 255, code >> 8)
  }
  const binary = String.fromCharCode.apply(null, bytes)
  return typeof btoa === 'function'
    ? btoa(binary)
    : Buffer.from(binary, 'binary').toString('base64')
}

export function isNode(): boolean {
  return typeof window === 'undefined'
}

export function addFieldAttributes(
  schemaXml: string,
  attributes: Record<string, any>
) {
  const fieldXmlJson = JSON.parse(xmljs.xml2json(schemaXml))
  fieldXmlJson.elements[0].attributes = {
    ...fieldXmlJson.elements[0].attributes,
    ...attributes
  }
  schemaXml = xmljs.json2xml(fieldXmlJson)
  return schemaXml
}
