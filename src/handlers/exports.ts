import { IProvisioningConfig } from '../provisioningconfig'
import { ClientSidePages } from './clientsidepages'
import { ComposedLook } from './composedlook'
import { ContentTypes } from './contenttypes'
import { CustomActions } from './customactions'
import { Features } from './features'
import { Files } from './files'
import { HandlerBase } from './handlerbase'
import { Hooks } from './hooks'
import { Lists } from './lists'
import { Navigation } from './navigation'
import { PropertyBagEntries } from './propertybagentries'
import { SiteFields } from './sitefields'
import { WebSettings } from './websettings'

export type Handler =
  | 'ClientSidePages'
  | 'ComposedLook'
  | 'ContentTypes'
  | 'CustomActions'
  | 'Features'
  | 'Files'
  | 'Lists'
  | 'Navigation'
  | 'PropertyBagEntries'
  | 'WebSettings'
  | 'SiteFields'
  | 'Hooks'

export const DefaultHandlerMap = (
  config: IProvisioningConfig
): Record<Handler, HandlerBase> => ({
  ClientSidePages: new ClientSidePages(config),
  ComposedLook: new ComposedLook(config),
  ContentTypes: new ContentTypes(config),
  CustomActions: new CustomActions(config),
  Features: new Features(config),
  Files: new Files(config),
  Lists: new Lists(config),
  Navigation: new Navigation(config),
  PropertyBagEntries: new PropertyBagEntries(config),
  WebSettings: new WebSettings(config),
  SiteFields: new SiteFields(config),
  Hooks: new Hooks(config)
})

export const DefaultHandlerSort: Record<Handler, number> = {
  ClientSidePages: 7,
  ComposedLook: 6,
  ContentTypes: 1,
  CustomActions: 5,
  Features: 2,
  Files: 4,
  Lists: 3,
  Navigation: 9,
  PropertyBagEntries: 8,
  WebSettings: 10,
  SiteFields: 0,
  Hooks: 11
}
