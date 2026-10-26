import { INavigationNodes } from '@pnp/sp/presets/all'
import { ISiteConnection } from '../connect'

/**
 * A navigation node as SharePoint stores it, with its children, in display order.
 */
export interface INavigationNodeState<TId = never> {
  /** Only in {@link readNavigationTree}: the Navigation handler re-adds every node on each run. */
  Id?: TId
  Title: string
  Url: string
  children: Array<INavigationNodeState<TId>>
}

/**
 * What the Navigation handler provisions on a web: the Quick Launch and the top navigation bar.
 */
export interface INavigationState<TId = never> {
  quickLaunch: Array<INavigationNodeState<TId>>
  topNavigationBar: Array<INavigationNodeState<TId>>
}

/**
 * Reads a node collection and, node by node, their children (any depth), in display order.
 *
 * @param nodes - Returns the collection: a fresh queryable on every call, since `select()` sets
 * the query on the instance
 */
async function readNodes(
  nodes: () => INavigationNodes
): Promise<Array<INavigationNodeState<number>>> {
  const infos = await nodes().select('Id', 'Title', 'Url')<
    Array<{ Id: number; Title: string; Url: string }>
  >()
  const states: Array<INavigationNodeState<number>> = []
  for (const info of infos)
    states.push({
      Id: info.Id,
      Title: info.Title,
      Url: info.Url,
      children: await readNodes(() => nodes().getById(info.Id).children)
    })
  return states
}

/**
 * Reads back the web's Quick Launch and top navigation bar, with node ids.
 *
 * @param conn - The web
 */
export async function readNavigationTree(
  conn: ISiteConnection
): Promise<INavigationState<number>> {
  return {
    quickLaunch: await readNodes(() => conn.web.navigation.quicklaunch),
    topNavigationBar: await readNodes(
      () => conn.web.navigation.topNavigationBar
    )
  }
}

/**
 * The node ids in a read-back, depth first.
 *
 * @param state - A read-back with ids
 */
export function navigationIds(state: INavigationState<number>): number[] {
  const ids = (nodes: Array<INavigationNodeState<number>>): number[] =>
    nodes.reduce(
      (all: number[], node) => [...all, node.Id, ...ids(node.children)],
      []
    )
  return [...ids(state.quickLaunch), ...ids(state.topNavigationBar)]
}

/**
 * A read-back without node ids, which change on every run.
 *
 * @param state - A read-back with ids
 */
export function withoutIds(state: INavigationState<number>): INavigationState {
  const strip = (
    nodes: Array<INavigationNodeState<number>>
  ): INavigationNodeState[] =>
    nodes.map((node) => ({
      Title: node.Title,
      Url: node.Url,
      children: strip(node.children)
    }))
  return {
    quickLaunch: strip(state.quickLaunch),
    topNavigationBar: strip(state.topNavigationBar)
  }
}

/**
 * Reads back the web's Quick Launch and top navigation bar: titles, URLs and order, any depth,
 * without ids.
 *
 * @param conn - The web
 */
export async function readNavigation(
  conn: ISiteConnection
): Promise<INavigationState> {
  return withoutIds(await readNavigationTree(conn))
}
