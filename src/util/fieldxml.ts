import * as xmljs from 'xml-js'

/**
 * Field schema XML (`<Field ...>`) for updating a list field or a site column in place: what a
 * template says merged onto what SharePoint stores, compared without the noise SharePoint adds, and
 * the type changes SharePoint can make in place without losing values (docs/specs/lists-re-apply.md,
 * docs/specs/sitefields-re-apply.md).
 */

interface IXmlNode {
  type: string
  name?: string
  attributes?: Record<string, string>
  elements?: IXmlNode[]
  text?: string | number | boolean
  cdata?: string
}

/** Attributes of a list field's identity and storage: SharePoint's value is kept. */
const STORED_ATTRIBUTES = [
  'ID',
  'Name',
  'StaticName',
  'ColName',
  'RowOrdinal',
  'SourceID'
]

/**
 * Type changes SharePoint applies to a list field in place, keeping its values (probe 2026-10-26:
 * Text to Note, Lookup to LookupMulti). Any other change, such as multi to single, loses values.
 */
const WIDENING: Record<string, string> = {
  Text: 'Note',
  Lookup: 'LookupMulti',
  User: 'UserMulti'
}

const GUID = /^{?([\da-f]{8}(?:-[\da-f]{4}){3}-[\da-f]{12})}?$/i

/**
 * A GUID in lower case without braces, or `undefined` when the value is not one (an unresolved
 * token, a list URL, `UserInfo`).
 *
 * @param value - The value
 */
export function guidOf(value?: string): string | undefined {
  const guid = GUID.exec(String(value ?? '').trim())
  return guid ? guid[1].toLowerCase() : undefined
}

/**
 * The `<Field>` element of a schema XML.
 *
 * @param xml - The schema XML
 */
function parseField(xml: string): IXmlNode {
  const root = xmljs.xml2js(xml, { compact: false }) as IXmlNode
  const field = (root.elements || []).find((node) => node.type === 'element')
  if (!field) throw new Error(`Not a field schema XML: ${xml.slice(0, 80)}`)
  return field
}

/**
 * An attribute value escaped for XML: the parser decodes `&amp;`, `&lt;` and character references,
 * and xml-js writes back only `"` escaped. A stored `&amp;&amp;` in a column formatter, or a line
 * break in a description, would otherwise come back raw.
 *
 * @param value - The decoded value
 */
function escapeAttribute(value: string): string {
  return String(value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/\r/g, '&#xD;')
    .replace(/\n/g, '&#xA;')
    .replace(/\t/g, '&#x9;')
}

/**
 * A node with its attribute values, and its children's, escaped for `js2xml`.
 *
 * @param node - The node
 */
function escaped(node: IXmlNode): IXmlNode {
  const attributes: Record<string, string> = {}
  for (const name of Object.keys(node.attributes || {}))
    attributes[name] = escapeAttribute(node.attributes[name])
  return {
    ...node,
    ...(node.attributes ? { attributes } : {}),
    ...(node.elements
      ? { elements: node.elements.map((child) => escaped(child)) }
      : {})
  }
}

const serialize = (field: IXmlNode): string =>
  xmljs.js2xml({ elements: [escaped(field)] } as any)

const childElements = (node: IXmlNode): IXmlNode[] =>
  (node.elements || []).filter((child) => child.type === 'element')

/**
 * The attributes of a field schema XML.
 *
 * @param xml - The schema XML
 */
export function fieldAttributes(xml: string): Record<string, string> {
  return { ...(parseField(xml).attributes || {}) }
}

/**
 * The type a field schema XML declares, with `Mult="TRUE"` on a Lookup, User or taxonomy field
 * read as the multi-value type SharePoint reports (`TypeAsString`).
 *
 * @param attributes - The field's attributes
 */
export function declaredType(attributes: Record<string, string>): string {
  const type = attributes.Type
  const multi = /^true$/i.test(attributes.Mult || '')
  if (multi && (type === 'Lookup' || type === 'User')) return `${type}Multi`
  if (multi && type === 'TaxonomyFieldType') return 'TaxonomyFieldTypeMulti'
  return type
}

/**
 * How a template's field type relates to the stored one: the same, a widening SharePoint applies in
 * place, or another change that would lose values.
 *
 * @param stored - The stored type (`TypeAsString`)
 * @param template - The template's type (`declaredType`)
 */
export function typeChange(
  stored: string,
  template: string
): 'same' | 'widening' | 'other' {
  if (!template || stored === template) return 'same'
  return WIDENING[stored] === template ? 'widening' : 'other'
}

/**
 * An attribute value as SharePoint compares it: GUIDs without braces in lower case, booleans in
 * lower case.
 *
 * @param value - The value
 */
function normaliseValue(value: unknown): string {
  const text = String(value)
  const guid = GUID.exec(text)
  if (guid) return guid[1].toLowerCase()
  return /^(true|false)$/i.test(text) ? text.toLowerCase() : text
}

/**
 * The value of a `Customization` property (a taxonomy field's `TermSetId`, `SspId`, `TextField`).
 *
 * @param field - The field element
 * @param name - The property's name
 */
function propertyValue(field: IXmlNode, name: string): string | undefined {
  const customization = childElements(field).find(
    (child) => child.name === 'Customization'
  )
  const array =
    customization &&
    childElements(customization).find(
      (child) => child.name === 'ArrayOfProperty'
    )
  const property =
    array && childElements(array).find((child) => propertyName(child) === name)
  const value =
    property && childElements(property).find((child) => child.name === 'Value')
  const text = value && (value.elements || []).find((n) => n.type === 'text')
  return text ? String(text.text) : undefined
}

/**
 * A target attribute or `Customization` property of a field.
 *
 * @param field - The field element
 * @param name - The attribute or property
 */
function targetValue(field: IXmlNode, name: string): string | undefined {
  return (field.attributes || {})[name] ?? propertyValue(field, name)
}

/**
 * What a field points at, by which its stored values are read: a lookup's `List` and `WebId`, a
 * taxonomy field's `TermSetId` and `SspId`.
 */
export const TARGET = ['List', 'WebId', 'TermSetId', 'SspId']

/**
 * The value of one target (`TARGET`) of a field schema XML: the attribute, or the `Customization`
 * property; `undefined` when the field has none.
 *
 * @param xml - The schema XML
 * @param name - The target
 */
export function fieldTarget(xml: string, name: string): string | undefined {
  return targetValue(parseField(xml), name)
}

/**
 * Whether the template points a field at another target (a lookup at another list, a taxonomy
 * field at another term set) than the stored one. SharePoint accepts the change and keeps the
 * stored values, which then name items or terms of the old target.
 *
 * @param stored - The stored schema XML
 * @param template - The template's schema XML, tokens replaced
 */
export function targetChanged(stored: string, template: string): boolean {
  const s = parseField(stored)
  const t = parseField(template)
  return TARGET.some((name) => {
    const [before, after] = [targetValue(s, name), targetValue(t, name)]
    return (
      before !== undefined &&
      after !== undefined &&
      normaliseValue(before) !== normaliseValue(after)
    )
  })
}

/**
 * The `Name` of a `<Property>` in a `Customization`.
 *
 * @param property - The property element
 */
function propertyName(property: IXmlNode): string | undefined {
  const name = childElements(property).find((child) => child.name === 'Name')
  const text =
    name && (name.elements || []).find((node) => node.type === 'text')
  return text ? String(text.text) : undefined
}

/**
 * A `Customization` with the template's properties replacing the stored ones of the same name, and
 * the stored ones the template does not mention kept (a taxonomy field's `TextField`, the hidden
 * note field SharePoint made for it).
 *
 * @param stored - The stored element
 * @param template - The template's element
 */
function mergeCustomization(
  stored: IXmlNode,
  template: IXmlNode,
  keep: string[]
): IXmlNode {
  const array = (node: IXmlNode) =>
    childElements(node).find((child) => child.name === 'ArrayOfProperty')
  const storedArray = array(stored)
  const templateArray = array(template)
  if (!storedArray || !templateArray) return template
  const properties = childElements(storedArray)
  for (const property of childElements(templateArray)) {
    const index = properties.findIndex(
      (existing) => propertyName(existing) === propertyName(property)
    )
    if (index < 0) properties.push(property)
    else if (!keep.includes(propertyName(property)))
      properties[index] = property
  }
  return {
    ...template,
    elements: (template.elements || []).map((node) =>
      node === templateArray ? { ...templateArray, elements: properties } : node
    )
  }
}

/**
 * The template's field merged onto the stored one, to update a list field or a site column in place.
 * The template's attributes win, except the stored field's identity and storage (`ID`, `Name`,
 * `StaticName`, `ColName`, `RowOrdinal`, `SourceID`) and those in `keep`; `Version` is dropped. The
 * type is written as SharePoint stores it (`declaredType`): `Type="Lookup" Mult="TRUE"` merged as
 * written would set Type="Lookup" on a LookupMulti column. A child element the template declares
 * (`CHOICES`, `Default`, `Formula`, `Validation`, ...) replaces the stored one, unless `keep` names
 * it; `Customization` properties are merged by name (those in `keep` stay stored).
 *
 * @param stored - The stored schema XML
 * @param template - The template's schema XML, tokens replaced
 * @param keep - More attributes, `Customization` properties or child elements whose stored value is
 * kept (an attribute is left out when the stored field has none)
 */
export function mergeFieldXml(
  stored: string,
  template: string,
  keep: string[] = []
): string {
  const s = parseField(stored)
  const t = parseField(template)
  const storedAttributes = s.attributes || {}
  const attributes: Record<string, string> = {
    ...storedAttributes,
    ...(t.attributes || {})
  }
  if (t.attributes && t.attributes.Type)
    attributes.Type = declaredType(t.attributes)
  for (const name of [...STORED_ATTRIBUTES, ...keep]) {
    if (storedAttributes[name] !== undefined)
      attributes[name] = storedAttributes[name]
    else if (keep.includes(name)) delete attributes[name]
  }
  delete attributes.Version
  const elements = childElements(s)
  for (const child of childElements(t)) {
    if (keep.includes(child.name)) continue
    const index = elements.findIndex((existing) => existing.name === child.name)
    if (index < 0) elements.push(child)
    else
      elements[index] =
        child.name === 'Customization'
          ? mergeCustomization(elements[index], child, keep)
          : child
  }
  return serialize({ ...s, attributes, elements })
}

/**
 * A field schema XML with `attributes` set; an undefined value leaves the attribute as it is.
 * `Version` is dropped, so the XML can be sent back after the field changed.
 *
 * @param xml - The schema XML
 * @param attributes - The attributes to set
 */
export function setFieldAttributes(
  xml: string,
  attributes: Record<string, any>
): string {
  const field = parseField(xml)
  const merged: Record<string, string> = { ...(field.attributes || {}) }
  for (const name of Object.keys(attributes))
    if (attributes[name] !== undefined) merged[name] = String(attributes[name])
  delete merged.Version
  return serialize({ ...field, attributes: merged })
}

/**
 * A node as SharePoint compares it: attributes sorted and normalised, `Version` left out,
 * whitespace-only text ignored.
 *
 * @param node - The node
 */
function canonical(node: IXmlNode): string {
  if (node.type === 'text' || node.type === 'cdata') {
    // A GUID in text (a taxonomy field's TermSetId) is compared as SharePoint stores it; other text
    // as it is, so "True" and "true" in a CHOICE differ.
    const raw = String(node.text ?? node.cdata ?? '').trim()
    const guid = GUID.exec(raw)
    const text = guid ? guid[1].toLowerCase() : raw
    return text ? JSON.stringify(text) : ''
  }
  if (node.type !== 'element') return ''
  const attributes = node.attributes || {}
  const pairs = Object.keys(attributes)
    .filter((name) => name !== 'Version')
    .sort()
    .map(
      (name) => `${name}=${JSON.stringify(normaliseValue(attributes[name]))}`
    )
  const children = (node.elements || [])
    .map((child) => canonical(child))
    .join('')
  return `<${node.name} ${pairs.join(' ')}>${children}</${node.name}>`
}

/**
 * Whether two field schema XMLs say the same, ignoring attribute order, `Version`, GUID braces and
 * case, and the case of booleans.
 *
 * @param a - A schema XML
 * @param b - Another
 */
export function sameFieldXml(a: string, b: string): boolean {
  return canonical(parseField(a)) === canonical(parseField(b))
}
