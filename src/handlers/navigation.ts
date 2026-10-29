import { isArray } from '@pnp/core'
import { INavigationNodes, IWeb } from '@pnp/sp/presets/all'
import { IProvisioningConfig } from '../provisioningconfig'
import { INavigation, INavigationNode } from '../schema'
import { replaceUrlTokens } from '../util'
import {
  INavigationUrlBase,
  normaliseNavigationUrl
} from '../util/navigationurl'
import { HandlerBase } from './handlerbase'

/**
 * A node as the handler reads it. A collection is read in one request, its children expanded to
 * the template's depth; `Children` is absent below that depth.
 */
interface IStoredNode {
  Id: number
  Title: string
  Url: string
  Children?: IStoredNode[]
}

/**
 * One level of a collection (the collection itself, or the children of a node the template kept)
 * and what the commit does there.
 */
interface ILevel {
  /** For messages: `the Quick Launch`, `the children of 'Prosjekter' in the Quick Launch` */
  label: string
  /** The level's nodes: a fresh queryable on every call */
  nodes: () => INavigationNodes
  /** The level as read, in display order */
  stored: IStoredNode[]
  /** The stored nodes the template does not declare: deleted, their children with them */
  deletes: IStoredNode[]
  /** The template's nodes in its order: the id of the node kept for it, or of the node added */
  order: Array<{ Id: number; Title: string; added: boolean }>
}

/**
 * A collection (the Quick Launch or the top navigation bar) through one run.
 */
interface ICollection {
  label: string
  /** The collection: a fresh queryable on every call */
  root: () => INavigationNodes
  template: INavigationNode[]
  levels: ILevel[]
  /** Kept nodes whose URL the template may set (`IgnoreExisting`, or a duplicated title) */
  updates: Array<{ stored: IStoredNode; node: INavigationNode }>
  added: number
  updated: number
}

/**
 * A write made before anything is deleted or moved, and how to undo it: a node this run added is
 * deleted (its children go with it), a URL it changed is set back.
 */
interface IJournalEntry {
  collection: ICollection
  Id: number
  Title: string
  /** The URL before this run changed it; `undefined` for a node this run added */
  previousUrl?: string
}

/**
 * Describes the Navigation Object Handler
 */
export class Navigation extends HandlerBase {
  /**
   * Creates a new instance of the Navigation class
   *
   * @param config - Provisioning config
   */
  constructor(config: IProvisioningConfig) {
    super('Navigation', config)
  }

  /**
   * Provisioning navigation. The template owns each collection it names (replace mode): a node is
   * matched to a stored node at the same level by exact title and kept (id, URL and audiences), a
   * missing node is added, a stored node the template does not declare is deleted, and the nodes
   * are moved into the template's order. An unchanged re-apply writes nothing.
   *
   * Writes go in the order adds, URL updates, deletes, moves, the Quick Launch before the top
   * navigation bar. The adds and URL updates of both collections are journaled; when SharePoint
   * refuses one (a URL whose file or folder does not exist), they are undone and the run fails, so
   * neither collection changes. A failed delete or move fails the run as it is.
   *
   * @param web - The web
   * @param navigation - The navigation to provision
   */
  public async ProvisionObjects(
    web: IWeb,
    navigation: INavigation
  ): Promise<void> {
    super.scope_started()
    const collections: ICollection[] = []
    if (isArray(navigation.QuickLaunch))
      collections.push(
        Navigation.collection(
          'the Quick Launch',
          () => web.navigation.quicklaunch,
          navigation.QuickLaunch
        )
      )
    if (isArray(navigation.TopNavigationBar))
      collections.push(
        Navigation.collection(
          'the top navigation bar',
          () => web.navigation.topNavigationBar,
          navigation.TopNavigationBar
        )
      )
    const journal: IJournalEntry[] = []
    // Read once, and only when a URL has to be compared. `web` is this handler's own queryable
    // (WebProvisioner hands each handler a fresh one), and the collections never inherit its query.
    let base: Promise<INavigationUrlBase> | undefined
    const webBase = () => {
      if (!base)
        base = web
          .select('Url', 'ServerRelativeUrl')<{
            Url: string
            ServerRelativeUrl: string
          }>()
          .then((info) => ({
            serverRelativeUrl: info.ServerRelativeUrl,
            origin: new URL(info.Url).origin
          }))
      return base
    }
    try {
      try {
        for (const collection of collections) {
          await this.stage(collection, journal)
          await this.updateUrls(collection, journal, webBase)
        }
      } catch (error) {
        await this.rollback(journal, webBase)
        throw error
      }
      for (const collection of collections) await this.commit(collection)
      super.scope_ended()
    } catch (error) {
      super.scope_ended(error)
      throw error
    }
  }

  /**
   * A collection before the run.
   *
   * @param label - For messages
   * @param root - Returns the collection
   * @param template - The template's nodes for it
   */
  private static collection(
    label: string,
    root: () => INavigationNodes,
    template: INavigationNode[]
  ): ICollection {
    return {
      label,
      root,
      template,
      levels: [],
      updates: [],
      added: 0,
      updated: 0
    }
  }

  /**
   * Reads a collection and adds the nodes it lacks, with their children.
   *
   * @param collection - The collection
   * @param journal - Where the adds are recorded
   */
  private async stage(
    collection: ICollection,
    journal: IJournalEntry[]
  ): Promise<void> {
    let stored: IStoredNode[]
    try {
      stored = await Navigation.read(
        collection.root,
        Navigation.depthOf(collection.template)
      )
    } catch (error) {
      throw new Error(
        `Failed to read ${collection.label}: ${Navigation.messageOf(error)}`
      )
    }
    await this.stageLevel(
      collection,
      collection.root,
      collection.label,
      collection.template,
      stored,
      journal
    )
  }

  /**
   * Matches a level's template nodes to its stored nodes, adds the missing ones, and goes on into
   * the children of each node it keeps. A template node without `Children` keeps none: the stored
   * children of its node are deleted.
   *
   * @param collection - The collection
   * @param nodes - Returns the level's nodes
   * @param label - The level, for messages
   * @param template - The template's nodes at this level
   * @param stored - The stored nodes at this level
   * @param journal - Where the adds are recorded
   */
  private async stageLevel(
    collection: ICollection,
    nodes: () => INavigationNodes,
    label: string,
    template: INavigationNode[],
    stored: IStoredNode[],
    journal: IJournalEntry[]
  ): Promise<void> {
    const level: ILevel = { label, nodes, stored, deletes: [], order: [] }
    collection.levels.push(level)
    const kept = stored.map(() => false)
    for (const node of template) {
      // Exact title, the first stored node not matched yet.
      const index = stored.findIndex(
        (candidate, at) => !kept[at] && candidate.Title === node.Title
      )
      if (index === -1) {
        const Id = await this.addNode(collection, nodes, node, journal)
        level.order.push({ Id, Title: node.Title, added: true })
        continue
      }
      kept[index] = true
      const match = stored[index]
      level.order.push({ Id: match.Id, Title: node.Title, added: false })
      // A kept node keeps its stored URL, unless the template says otherwise or its title is
      // stored more often than the template declares it (then it gets the template's URL, as
      // before). Counting the template's too keeps a title it declares twice, stored twice after
      // the first run, from being rewritten on the next.
      const duplicate =
        stored.filter((candidate) => candidate.Title === node.Title).length >
        template.filter((other) => other.Title === node.Title).length
      if (node.IgnoreExisting === true || duplicate)
        collection.updates.push({ stored: match, node })
      const children = node.Children || []
      const childNodes = Navigation.childrenOf(collection, match.Id)
      await this.stageLevel(
        collection,
        childNodes,
        `the children of '${node.Title}' in ${collection.label}`,
        children,
        // Read with the collection; on its own only if SharePoint left them out.
        isArray(match.Children)
          ? match.Children
          : await Navigation.read(childNodes, Navigation.depthOf(children)),
        journal
      )
    }
    level.deletes = stored.filter((_, at) => !kept[at])
  }

  /**
   * Adds a node at the end of a level, then its children under it.
   *
   * @param collection - The collection
   * @param nodes - Returns the level's nodes
   * @param node - The template's node
   * @param journal - Where the add is recorded; `undefined` under a node this run added, whose
   * delete takes its children with it
   */
  private async addNode(
    collection: ICollection,
    nodes: () => INavigationNodes,
    node: INavigationNode,
    journal?: IJournalEntry[]
  ): Promise<number> {
    const url = replaceUrlTokens(node.Url, this.config)
    let Id: number
    try {
      Id = (await nodes().add(node.Title, url)).Id
    } catch (error) {
      throw Navigation.refused(node.Title, url, error)
    }
    if (journal) journal.push({ collection, Id, Title: node.Title })
    collection.added++
    for (const child of node.Children || [])
      await this.addNode(
        collection,
        Navigation.childrenOf(collection, Id),
        child
      )
    return Id
  }

  /**
   * Sets the template's URL on the kept nodes that take it, where it differs from the stored one
   * once both are normalised. Only `Url` is sent, so the node keeps its audiences.
   *
   * @param collection - The collection
   * @param journal - Where the changes are recorded
   * @param webBase - The web the URLs are resolved against
   */
  private async updateUrls(
    collection: ICollection,
    journal: IJournalEntry[],
    webBase: () => Promise<INavigationUrlBase>
  ): Promise<void> {
    for (const { stored, node } of collection.updates) {
      const url = replaceUrlTokens(node.Url, this.config)
      const base = await webBase()
      if (
        normaliseNavigationUrl(url, base) ===
        normaliseNavigationUrl(stored.Url, base)
      )
        continue
      try {
        await collection.root().getById(stored.Id).update({ Url: url })
      } catch (error) {
        throw Navigation.refused(node.Title, url, error)
      }
      journal.push({
        collection,
        Id: stored.Id,
        Title: node.Title,
        previousUrl: stored.Url
      })
      collection.updated++
    }
  }

  /**
   * Undoes the journaled writes, the last first, then reads the collections back: SharePoint
   * reports success for a delete or an update of a node it does not have (probe X1), so a write
   * that did not take is only seen there. Whatever is left is logged as an error naming the node.
   *
   * @param journal - The writes
   * @param webBase - The web the URLs are resolved against
   */
  private async rollback(
    journal: IJournalEntry[],
    webBase: () => Promise<INavigationUrlBase>
  ): Promise<void> {
    if (journal.length === 0) return
    super.log_info(
      'rollback',
      `Undoing ${journal.length} navigation write(s) of this run.`
    )
    const failed: IJournalEntry[] = []
    for (const entry of journal.slice().reverse()) {
      const node = entry.collection.root().getById(entry.Id)
      try {
        await (entry.previousUrl === undefined
          ? node.delete()
          : node.update({ Url: entry.previousUrl }))
      } catch (error) {
        failed.push(entry)
        super.log_error('rollback', Navigation.leftover(entry, error))
      }
    }
    const collections = journal
      .map((entry) => entry.collection)
      .filter((collection, at, all) => all.indexOf(collection) === at)
    for (const collection of collections) {
      const entries = journal.filter(
        (entry) => entry.collection === collection && !failed.includes(entry)
      )
      if (entries.length === 0) continue
      let stored: IStoredNode[]
      let base: INavigationUrlBase | undefined
      try {
        stored = Navigation.flatten(
          await Navigation.read(
            collection.root,
            Navigation.depthOf(collection.template)
          )
        )
        if (entries.some((entry) => entry.previousUrl !== undefined))
          base = await webBase()
      } catch (error) {
        super.log_error(
          'rollback',
          `Failed to read ${
            collection.label
          } back after undoing this run's writes: ${Navigation.messageOf(
            error
          )}`
        )
        continue
      }
      for (const entry of entries) {
        const found = stored.find((node) => node.Id === entry.Id)
        if (!found) continue
        if (
          entry.previousUrl === undefined ||
          normaliseNavigationUrl(found.Url, base) !==
            normaliseNavigationUrl(entry.previousUrl, base)
        )
          super.log_error(
            'rollback',
            Navigation.leftover(entry, 'SharePoint reported success')
          )
      }
    }
  }

  /**
   * Deletes the stored nodes the template does not declare, then moves the nodes into the
   * template's order. A failure is logged and fails the run.
   *
   * @param collection - The collection
   */
  private async commit(collection: ICollection): Promise<void> {
    let deleted = 0
    let moved = 0
    for (const level of collection.levels)
      for (const node of level.deletes) {
        try {
          await collection.root().getById(node.Id).delete()
        } catch (error) {
          throw this.failed(
            `Failed to delete navigation node '${node.Title}' (${
              node.Id
            }) from ${level.label}: ${Navigation.messageOf(error)}`
          )
        }
        deleted++
      }
    for (const level of collection.levels) moved += await this.moveNodes(level)
    super.log_info(
      'commit',
      `Navigation nodes in ${collection.label}: ${collection.added} added, ${collection.updated} URL(s) updated, ${deleted} deleted, ${moved} moved.`
    )
  }

  /**
   * Moves a level's nodes into the template's order, once its other nodes are deleted. SharePoint
   * can only move a node after another one (`MoveAfter`), so each node goes after the one before it
   * in the template, where it is not already there; with only template nodes left, the first one
   * ends up first. The level is then read back, since SharePoint does not check that both nodes
   * are in the collection (probe N4).
   *
   * @param level - The level
   */
  private async moveNodes(level: ILevel): Promise<number> {
    const deleted = new Set(level.deletes.map((node) => node.Id))
    const current = [
      ...level.stored.map((node) => node.Id).filter((Id) => !deleted.has(Id)),
      ...level.order.filter((node) => node.added).map((node) => node.Id)
    ]
    let moved = 0
    for (let at = 1; at < level.order.length; at++) {
      const node = level.order[at]
      const previous = level.order[at - 1]
      if (current[current.indexOf(previous.Id) + 1] === node.Id) continue
      try {
        await level.nodes().moveAfter(node.Id, previous.Id)
      } catch (error) {
        throw this.failed(
          `Failed to move navigation node '${node.Title}' after '${
            previous.Title
          }' in ${level.label}: ${Navigation.messageOf(error)}`
        )
      }
      current.splice(current.indexOf(node.Id), 1)
      current.splice(current.indexOf(previous.Id) + 1, 0, node.Id)
      moved++
    }
    if (moved === 0) return moved
    let ids: number[]
    try {
      ids = (await level.nodes().select('Id')<Array<{ Id: number }>>()).map(
        (node) => node.Id
      )
    } catch (error) {
      throw this.failed(
        `Failed to read ${
          level.label
        } back after moving its nodes: ${Navigation.messageOf(error)}`
      )
    }
    if (ids.join() !== level.order.map((node) => node.Id).join())
      throw this.failed(
        `The nodes in ${
          level.label
        } are not in the template's order after moving them: read ${ids.join(
          ', '
        )}, expected ${level.order.map((node) => node.Id).join(', ')}.`
      )
    return moved
  }

  /**
   * Logs a failure as an error and returns it, to be thrown.
   *
   * @param message - What failed
   */
  private failed(message: string): Error {
    super.log_error('commit', message)
    return new Error(message)
  }

  /**
   * The children of a node, as a level.
   *
   * @param collection - The collection the node is in
   * @param Id - The node
   */
  private static childrenOf(
    collection: ICollection,
    Id: number
  ): () => INavigationNodes {
    // Any collection's getById reaches any node of the web, child or not (probe N6).
    return () => collection.root().getById(Id).children
  }

  /**
   * Reads a level of nodes in one request, with their children `depth` levels down.
   *
   * @param nodes - Returns the level's nodes: a fresh queryable, since `select()` sets the query on
   * the instance
   * @param depth - How many levels of children to read
   */
  private static read(
    nodes: () => INavigationNodes,
    depth: number
  ): Promise<IStoredNode[]> {
    const select: string[] = []
    const expand: string[] = []
    let path = ''
    for (let level = 0; level <= depth; level++) {
      select.push(`${path}Id`, `${path}Title`, `${path}Url`)
      if (level === depth) break
      path += 'Children/'
      expand.push(path.slice(0, -1))
    }
    const query = nodes().select(...select)
    return (expand.length > 0 ? query.expand(...expand) : query)<
      IStoredNode[]
    >()
  }

  /**
   * How many levels a template's nodes have (0 for none).
   *
   * @param nodes - The template's nodes
   */
  private static depthOf(nodes: INavigationNode[]): number {
    return nodes.reduce(
      (depth, node) =>
        Math.max(depth, 1 + Navigation.depthOf(node.Children || [])),
      0
    )
  }

  /**
   * Every node of a read, children included.
   *
   * @param nodes - The nodes read
   */
  private static flatten(nodes: IStoredNode[]): IStoredNode[] {
    return nodes.reduce(
      (all: IStoredNode[], node) => [
        ...all,
        node,
        ...Navigation.flatten(node.Children || [])
      ],
      []
    )
  }

  /**
   * The error for a node SharePoint refused to add, or whose URL it refused to set. SharePoint
   * refuses a URL whose server-relative path is no file or folder in the site collection
   * (`-2130247147, Microsoft.SharePoint.SPException`, probe N2); its message is in the signed-in
   * account's language and does not name the node.
   *
   * @param title - The node's title
   * @param url - The URL sent, tokens replaced
   * @param error - What the request failed with
   */
  private static refused(title: string, url: string, error: any): Error {
    return new Error(
      `Navigation node '${title}' (${url}) was refused by SharePoint: ${Navigation.messageOf(
        error
      )}`
    )
  }

  /**
   * What is left of a write the rollback could not undo.
   *
   * @param entry - The write
   * @param error - Why it could not be undone
   */
  private static leftover(entry: IJournalEntry, error: any): string {
    const reason = Navigation.messageOf(error)
    return entry.previousUrl === undefined
      ? `Navigation node '${entry.Title}' (${entry.Id}), added by this run, is left in ${entry.collection.label}: ${reason}`
      : `Navigation node '${entry.Title}' (${entry.Id}) in ${entry.collection.label} keeps the URL this run set instead of ${entry.previousUrl}: ${reason}`
  }

  /**
   * SharePoint's own message for a failed request (`odata.error`, from the response body PnPjs puts
   * in its error's message), else the error's message.
   *
   * @param error - What a request failed with
   */
  private static messageOf(error: any): string {
    const message = `${(error && error.message) || error}`
    const body = message.indexOf('::> ')
    if (body !== -1)
      try {
        const parsed = JSON.parse(message.slice(body + 4))
        const odata = parsed['odata.error'] || parsed.error
        const value = odata && odata.message && odata.message.value
        if (typeof value === 'string' && value) return value
      } catch {}
    return message
  }
}
