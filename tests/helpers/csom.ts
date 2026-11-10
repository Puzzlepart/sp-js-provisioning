import * as xmljs from 'xml-js'

/** Where a CSOM request is posted, below the web. */
export const CSOM_PATH = '_vti_bin/client.svc/ProcessQuery'

/** The types the engine's requests name by id, by the name SharePoint's object model gives them. */
const TYPES: Record<string, string> = {
  '{3747adcd-a3c3-41b9-bfab-4a64dd2f1e0a}': 'RequestContext',
  '{168f3091-4554-4f14-8866-b20d48e45b54}': 'ContentTypeCreationInformation',
  '{63fb2c92-8f65-4bbb-a658-b6cd294403f4}': 'FieldLinkCreationInformation',
  '{981cbc68-9edc-4f8d-872f-71146fcbb84f}': 'TaxonomySession'
}

/**
 * A parameter or a property's value as a test reads it: a text, a number, a boolean, `null`, a
 * list, a GUID (`{ guid }`, bare and in lower case), another object of the request (`{ ref }`, its
 * path) or a value object (`{ type, properties }`).
 */
export type CsomArg =
  | string
  | number
  | boolean
  | null
  | CsomArg[]
  | { guid: string }
  | { ref: ICsomSegment[] }
  | { type: string; properties: Record<string, CsomArg> }

/** A step of an object's path: a property, or a method with its parameters. */
export interface ICsomSegment {
  name: string
  /** Whether it is a method (`GetById(…)`) and not a property (`ContentTypes`). */
  method: boolean
  args: CsomArg[]
}

/** An action of a CSOM request, as a test reads it. */
export interface ICsomAction {
  /** The action's id: the key of its result in the response. */
  id: number
  kind: 'instantiate' | 'query' | 'set' | 'call'
  /** The object it acts on, as text: `Web.ContentTypes.GetById("0x0100AB")`. */
  path: string
  /** The object it acts on, step by step from the web (or from a static method). */
  segments: ICsomSegment[]
  /** The property (`set`) or method (`call`). */
  name?: string
  /** The value (`set`, one) or the parameters (`call`). */
  args?: CsomArg[]
  /** The properties a `query` reads of the object. */
  properties?: string[]
  /** The properties a `query` reads of each item of a collection. */
  childProperties?: string[]
  /** The action in one line: `call Web.ContentTypes.GetById("0x0100AB").Update(true)`. */
  text: string
}

/** An error for a fake to answer an action with: SharePoint's `ErrorInfo`. */
export class CsomFailure {
  constructor(
    public message: string,
    public typeName = 'Microsoft.SharePoint.SPException',
    public code = -2147467259
  ) {}
}

interface IElement {
  type?: string
  name?: string
  attributes?: Record<string, string>
  elements?: IElement[]
}

const children = (element: IElement, name?: string): IElement[] =>
  ((element && element.elements) || []).filter(
    (child) => child.type === 'element' && (!name || child.name === name)
  )

const textOf = (element: IElement): string =>
  (element.elements || [])
    .filter((child) => child.type === 'text' || child.type === 'cdata')
    .map((child) => String((child as any)[child.type]))
    .join('')

/**
 * A value as text.
 *
 * @param arg - The value
 */
export function argText(arg: CsomArg): string {
  if (arg === null) return 'null'
  if (typeof arg === 'string') return JSON.stringify(arg)
  if (typeof arg !== 'object') return String(arg)
  if (Array.isArray(arg)) return `[${arg.map(argText).join(', ')}]`
  if ('guid' in arg) return `{${arg.guid}}`
  if ('ref' in arg) return `@${pathText(arg.ref)}`
  return `${arg.type}{ ${Object.keys(arg.properties)
    .map((key) => `${key}: ${argText(arg.properties[key])}`)
    .join(', ')} }`
}

/**
 * An object's path as text.
 *
 * @param segments - The path
 */
export function pathText(segments: ICsomSegment[]): string {
  return segments
    .map((segment) =>
      segment.method
        ? `${segment.name}(${segment.args.map(argText).join(', ')})`
        : segment.name
    )
    .join('.')
}

/**
 * The step of an action's path with a name (`GetById`), the last one when there are several.
 *
 * @param action - The action
 * @param name - The property or method
 */
export function segmentOf(
  action: ICsomAction,
  name: string
): ICsomSegment | undefined {
  return action.segments.filter((segment) => segment.name === name).pop()
}

/**
 * Reads a CSOM request into its actions, each with the path of its object written out, so a test
 * can assert what is asked for instead of ids and XML.
 *
 * @param xml - The request
 */
export function parseCsom(xml: string): ICsomAction[] {
  const root = children(
    xmljs.xml2js(xml, { compact: false }) as IElement,
    'Request'
  )[0]
  if (!root) throw new Error(`Not a CSOM request: ${xml.slice(0, 80)}`)
  const paths: Record<string, IElement> = {}
  for (const path of children(children(root, 'ObjectPaths')[0]))
    paths[path.attributes.Id] = path

  const value = (element: IElement): CsomArg => {
    const attributes = element.attributes || {}
    if (attributes.ObjectPathId !== undefined)
      return { ref: segmentsOf(attributes.ObjectPathId) }
    if (attributes.TypeId !== undefined)
      return {
        type: TYPES[attributes.TypeId] || attributes.TypeId,
        properties: children(element, 'Property').reduce(
          (all: Record<string, CsomArg>, property) => {
            all[property.attributes.Name] = value(property)
            return all
          },
          {}
        )
      }
    switch (attributes.Type) {
      case 'Null':
        return null
      case 'String':
        return textOf(element)
      case 'Boolean':
        return textOf(element) === 'true'
      case 'Int32':
        return Number(textOf(element))
      case 'Guid':
        return { guid: textOf(element).replace(/[{}]/g, '').toLowerCase() }
      case 'Array':
        return children(element, 'Object').map(value)
      default:
        throw new Error(`Unknown parameter type ${attributes.Type}`)
    }
  }
  const parameters = (element: IElement): CsomArg[] =>
    children(children(element, 'Parameters')[0], 'Parameter').map(value)

  const segmentsOf = (id: string): ICsomSegment[] => {
    const path = paths[id]
    if (!path) throw new Error(`The request has no object path ${id}`)
    const { Name, ParentId, TypeId } = path.attributes
    switch (path.name) {
      case 'StaticProperty':
        // `RequestContext.Current` is where every path of the web starts: left out.
        return TYPES[TypeId] === 'RequestContext' && Name === 'Current'
          ? []
          : [
              { name: TYPES[TypeId] || TypeId, method: false, args: [] },
              { name: Name, method: false, args: [] }
            ]
      case 'StaticMethod':
        return [
          { name: TYPES[TypeId] || TypeId, method: false, args: [] },
          { name: Name, method: true, args: parameters(path) }
        ]
      case 'Property':
        return [
          ...segmentsOf(ParentId),
          { name: Name, method: false, args: [] }
        ]
      case 'Method':
        return [
          ...segmentsOf(ParentId),
          { name: Name, method: true, args: parameters(path) }
        ]
      default:
        throw new Error(`Unknown object path <${path.name}>`)
    }
  }
  const selected = (element: IElement, name: string): string[] | undefined => {
    const query = children(element, name)[0]
    return query
      ? children(children(query, 'Properties')[0], 'Property').map(
          (property) => property.attributes.Name
        )
      : undefined
  }

  return children(children(root, 'Actions')[0]).map(
    (action): ICsomAction => {
      const id = Number(action.attributes.Id)
      const segments = segmentsOf(action.attributes.ObjectPathId)
      const path = pathText(segments)
      switch (action.name) {
        case 'ObjectPath':
          return {
            id,
            kind: 'instantiate',
            path,
            segments,
            text: `instantiate ${path}`
          }
        case 'Query': {
          const properties = selected(action, 'Query') || []
          const childProperties = selected(action, 'ChildItemQuery')
          return {
            id,
            kind: 'query',
            path,
            segments,
            properties,
            childProperties,
            text: `query ${path} [${properties.join(', ')}]${
              childProperties ? ` items [${childProperties.join(', ')}]` : ''
            }`
          }
        }
        case 'SetProperty': {
          const set = value(children(action, 'Parameter')[0])
          return {
            id,
            kind: 'set',
            path,
            segments,
            name: action.attributes.Name,
            args: [set],
            text: `set ${path}.${action.attributes.Name} = ${argText(set)}`
          }
        }
        case 'Method': {
          const args = parameters(action)
          return {
            id,
            kind: 'call',
            path,
            segments,
            name: action.attributes.Name,
            args,
            text: `call ${path}.${action.attributes.Name}(${args
              .map(argText)
              .join(', ')})`
          }
        }
        default:
          throw new Error(`Unknown action <${action.name}>`)
      }
    }
  )
}

/**
 * A CSOM request as lines, one per action.
 *
 * @param xml - The request
 */
export const describeCsom = (xml: string): string[] =>
  parseCsom(xml).map((action) => action.text)

/**
 * A stand-in for SharePoint's `ProcessQuery`: it reads each request into its actions, asks
 * `answer` for each one's result, and builds the response. Without an answer an `instantiate`
 * finds its object (`IsNull: false`), a `query` reads nothing (`{}`), and a `set` or a `call`
 * returns nothing. An answer that throws a {@link CsomFailure} fails the request there: the
 * actions after it do not run, as in SharePoint.
 */
export class FakeCsom {
  /** The actions that ran, in order; the one a request failed at is marked. */
  public done: Array<ICsomAction & { failed?: boolean }> = []
  /** The actions of each request, in the order the requests came. */
  public requests: ICsomAction[][] = []

  /**
   * @param answer - The result of an action; `undefined` for the default. `scratch` is an object
   * of the request the action belongs to, empty at its first action: what a fake keeps until the
   * request's `Update` or `CommitAll`.
   */
  constructor(
    private answer: (
      action: ICsomAction,
      scratch: Record<string, any>
    ) => any = () => undefined
  ) {}

  /**
   * The response to a request.
   *
   * @param xml - The request
   */
  public respond(xml: string): any[] {
    const actions = parseCsom(xml)
    this.requests.push(actions)
    const header: any = {
      SchemaVersion: '15.0.0.0',
      LibraryVersion: '16.0.0.0',
      ErrorInfo: null,
      TraceCorrelationId: '00000000-0000-0000-0000-000000000000'
    }
    const response: any[] = [header]
    const scratch: Record<string, any> = {}
    for (const action of actions) {
      let result: any
      try {
        result = this.answer(action, scratch)
      } catch (error) {
        if (!(error instanceof CsomFailure)) throw error
        this.done.push({ ...action, failed: true })
        header.ErrorInfo = {
          ErrorMessage: error.message,
          ErrorValue: null,
          TraceCorrelationId: header.TraceCorrelationId,
          ErrorCode: error.code,
          ErrorTypeName: error.typeName
        }
        return [header]
      }
      this.done.push(action)
      if (result === undefined && action.kind === 'instantiate')
        result = { IsNull: false }
      if (result === undefined && action.kind === 'query') result = {}
      if (result !== undefined) response.push(action.id, result)
    }
    return response
  }

  /** The actions that ran, in order, as lines (`FAILED` before the one a request failed at). */
  public get ran(): string[] {
    return this.done.map(line)
  }

  /**
   * The actions that change something, in order, as lines: every `set` and `call`, and each
   * object a request made (an `instantiate` of what an `Add` or a `Create…` returns).
   */
  public get writes(): string[] {
    return this.done.filter(isWrite).map(line)
  }
}

/**
 * Whether an action changes something: a `set`, a `call`, or the making of an object (an
 * `instantiate` of what an `Add` or a `Create…` returns).
 *
 * @param action - The action
 */
export function isWrite(action: ICsomAction): boolean {
  const last = action.segments[action.segments.length - 1]
  return (
    action.kind === 'set' ||
    action.kind === 'call' ||
    (action.kind === 'instantiate' &&
      !!last &&
      last.method &&
      /^(Add|Create)/.test(last.name))
  )
}

/**
 * An action that ran, as a line.
 *
 * @param action - The action
 */
const line = (action: ICsomAction & { failed?: boolean }): string =>
  `${action.failed ? 'FAILED ' : ''}${action.text}`
