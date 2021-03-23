import { IProvisioningConfig } from '../provisioningconfig'
import * as xmljs from 'xml-js'
import { TypedHash } from '@pnp/common'

export function replaceUrlTokens(
  string: string,
  config: IProvisioningConfig
): string {
  let siteAbsoluteUrl = null
  let siteRelativeUrl = null
  if (config.spfxContext) {
    siteAbsoluteUrl = config.spfxContext.pageContext.site.absoluteUrl
    siteRelativeUrl = config.spfxContext.pageContext.site.serverRelativeUrl
  } else if (window.hasOwnProperty('_spPageContextInfo')) {
    siteAbsoluteUrl = _spPageContextInfo.siteAbsoluteUrl
    siteRelativeUrl = _spPageContextInfo.siteServerRelativeUrl
  }
  return string
    .replace(/{site}/g, siteRelativeUrl)
    .replace(/{sitecollection}/g, siteAbsoluteUrl)
    .replace(/{wpgallery}/g, `${siteAbsoluteUrl}/_catalogs/wp`)
    .replace(
      /{hosturl}/g,
      `${window.location.protocol}//${window.location.host}:${window.location.port}`
    )
    .replace(/{themegallery}/g, `${siteAbsoluteUrl}/_catalogs/theme/15`)
}

export function makeUrlRelative(absUrl: string): string {
  return absUrl.replace(
    `${document.location.protocol}//${document.location.hostname}`,
    ''
  )
}

export function base64EncodeString(string: string): string {
  const bytes = []
  for (let index = 0; index < string.length; ++index) {
    bytes.push(string.charCodeAt(index), 0)
  }
  const b64encoded = window.btoa(String.fromCharCode.apply(null, bytes))
  return b64encoded
}

export function isNode(): boolean {
  return typeof window === 'undefined'
}

export function addFieldAttributes(
  schemaXml: string,
  attributes: TypedHash<any>
) {
  const fieldXmlJson = JSON.parse(xmljs.xml2json(schemaXml))
  fieldXmlJson.elements[0].attributes = {
    ...fieldXmlJson.elements[0].attributes,
    ...attributes
  }
  schemaXml = xmljs.json2xml(fieldXmlJson)
  return schemaXml
}
