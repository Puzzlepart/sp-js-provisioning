import * as xmljs from 'xml-js'
import { ISiteConnection } from '../connect'

/**
 * The parts of a field's `SchemaXml` that describe the column: every attribute of `<Field>` but
 * `Version` (which every update bumps), and the `Default`, `Formula`, `CHOICES` and `FieldRefs`
 * children. GUIDs are lower case without braces and `TRUE`/`FALSE` upper case, so a template's XML
 * and SharePoint's compare.
 */
export interface IFieldSchema {
  attributes: Record<string, string>
  Default?: string
  Formula?: string
  Choices?: string[]
  FieldRefs?: string[]
}

/**
 * A site column as SharePoint stores it.
 */
export interface ISiteFieldState {
  /** Lower case, without braces. */
  Id: string
  InternalName: string
  TypeAsString: string
  Title: string
  Group: string
  Required: boolean
  Hidden: boolean
  Description: string
  schema: IFieldSchema
}

/**
 * What the SiteFields handler provisions on a web: the named site columns, the internal names of
 * every column in their group (a duplicate shows up there), and each column's `SchemaXml`
 * `Version`, kept apart since every update bumps it.
 */
export interface ISiteFieldsState {
  fields: ISiteFieldState[]
  inGroup: string[]
  versions: Record<string, number>
}

const GUID = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i

/**
 * A GUID lower case without braces, `true`/`false` upper case, anything else as is.
 *
 * @param value - An attribute value
 */
function normalise(value: string): string {
  const guid = GUID.exec(value)
  if (guid) return guid[1].toLowerCase()
  return /^(true|false)$/i.test(value) ? value.toUpperCase() : value
}

/**
 * The text of an element (`<Default>Beta</Default>` -> `Beta`).
 *
 * @param element - An xml-js element
 */
function textOf(element: any): string {
  return (element.elements || [])
    .map((child: any) => (child.type === 'text' ? child.text : textOf(child)))
    .join('')
}

/**
 * Parses a field's XML, from a template or from `SchemaXml`, into the parts that describe it.
 *
 * @param xml - `<Field ...>...</Field>`
 */
export function fieldSchema(xml: string): IFieldSchema {
  const field = JSON.parse(xmljs.xml2json(xml)).elements[0]
  const declared: Record<string, string> = field.attributes || {}
  const attributes: Record<string, string> = {}
  for (const name of Object.keys(declared))
    if (name !== 'Version') attributes[name] = normalise(String(declared[name]))
  const schema: IFieldSchema = { attributes }
  for (const child of field.elements || []) {
    if (child.type !== 'element') continue
    if (child.name === 'Default') schema.Default = textOf(child)
    else if (child.name === 'Formula') schema.Formula = textOf(child)
    else if (child.name === 'CHOICES')
      schema.Choices = (child.elements || [])
        .filter((choice: any) => choice.name === 'CHOICE')
        .map(textOf)
    else if (child.name === 'FieldRefs')
      schema.FieldRefs = (child.elements || [])
        .filter((ref: any) => ref.name === 'FieldRef')
        .map((ref: any) => ref.attributes.Name)
  }
  return schema
}

/**
 * The `Version` attribute of a field's `SchemaXml`.
 *
 * @param xml - `SchemaXml`
 */
function versionOf(xml: string): number {
  const match = /^<Field\b[^>]*\sVersion="(\d+)"/.exec(xml)
  return match ? Number(match[1]) : 0
}

/**
 * Reads back the web's site columns with the given internal names, and the internal names of the
 * columns in a group.
 *
 * @param conn - The web
 * @param internalNames - The columns to read
 * @param group - The group to list (`SPJS Itest` by default)
 */
export async function readSiteFields(
  conn: ISiteConnection,
  internalNames: string[],
  group = 'SPJS Itest'
): Promise<ISiteFieldsState> {
  const fields: ISiteFieldState[] = []
  const versions: Record<string, number> = {}
  for (const name of internalNames) {
    const info = await conn.web.fields
      .getByInternalNameOrTitle(name)
      .select(
        'Id',
        'InternalName',
        'TypeAsString',
        'Title',
        'Group',
        'Required',
        'Hidden',
        'Description',
        'SchemaXml'
      )<{
      Id: string
      InternalName: string
      TypeAsString: string
      Title: string
      Group: string
      Required: boolean
      Hidden: boolean
      Description?: string
      SchemaXml: string
    }>()
    fields.push({
      Id: normalise(info.Id),
      InternalName: info.InternalName,
      TypeAsString: info.TypeAsString,
      Title: info.Title,
      Group: info.Group,
      Required: info.Required,
      Hidden: info.Hidden,
      Description: info.Description ?? '',
      schema: fieldSchema(info.SchemaXml)
    })
    versions[info.InternalName] = versionOf(info.SchemaXml)
  }
  const inGroup = (
    await conn.web.fields
      .filter(`Group eq '${group.replace(/'/g, "''")}'`)
      .select('InternalName')<Array<{ InternalName: string }>>()
  )
    .map((field) => field.InternalName)
    // Only the columns of this fixture family: on a web shared by the suites
    // (INTEGRATION_SCRATCH=run), the Lists suites' columns are in the group too.
    .filter((name) => name.startsWith('SPJSItestSf'))
    .sort()
  return { fields, inGroup, versions }
}
