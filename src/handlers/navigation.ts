import { HandlerBase } from './handlerbase'
import { INavigation, INavigationNode } from '../schema'
import { Web, NavigationNodes } from '@pnp/sp'
import { isArray } from '@pnp/common'
import { replaceUrlTokens } from '../util'
import { IProvisioningConfig } from '../provisioningconfig'

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
    super(Navigation.name, config)
  }

  /**
   * Provisioning navigation
   *
   * @param navigation - The navigation to provision
   */
  public async ProvisionObjects(
    web: Web,
    navigation: INavigation
  ): Promise<void> {
    super.scope_started()
    const promises = []
    if (isArray(navigation.QuickLaunch)) {
      promises.push(
        this.processNavTree(web.navigation.quicklaunch, navigation.QuickLaunch)
      )
    }
    if (isArray(navigation.TopNavigationBar)) {
      promises.push(
        this.processNavTree(
          web.navigation.topNavigationBar,
          navigation.TopNavigationBar
        )
      )
    }
    try {
      await Promise.all(promises)
      super.scope_ended()
    } catch (error) {
      super.scope_ended()
      throw error
    }
  }

  private async processNavTree(
    target: NavigationNodes,
    nodes: INavigationNode[]
  ): Promise<void> {
    try {
      const existingNodes = await target.get()
      await this.deleteExistingNodes(target)
      await nodes.reduce(
        (chain: any, node) =>
          chain.then(() => this.processNode(target, node, existingNodes)),
        Promise.resolve()
      )
    } catch (error) {
      throw error
    }
  }

  private async processNode(
    target: NavigationNodes,
    node: INavigationNode,
    existingNodes: any[]
  ): Promise<void> {
    const existingNode = existingNodes.filter((n) => n.Title === node.Title)
    if (existingNode.length === 1 && node.IgnoreExisting !== true) {
      node.Url = existingNode[0].Url
    }
    try {
      const result = await target.add(
        node.Title,
        replaceUrlTokens(node.Url, this.config)
      )
      if (isArray(node.Children)) {
        await this.processNavTree(result.node.children, node.Children)
      }
    } catch (error) {
      throw error
    }
  }

  private async deleteExistingNodes(target: NavigationNodes) {
    try {
      const existingNodes = await target.get()
      await existingNodes.reduce(
        (chain: Promise<void>, n: any) =>
          chain.then(() => this.deleteNode(target, n.Id)),
        Promise.resolve()
      )
    } catch (error) {
      throw error
    }
  }

  private async deleteNode(target: NavigationNodes, id: number): Promise<void> {
    try {
      await target.getById(id).delete()
    } catch (error) {
      throw error
    }
  }
}
