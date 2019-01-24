import { TypedHash } from "@pnp/common";
import { HandlerBase } from "./handlerbase";
import { ComposedLook } from "./composedlook";
import { CustomActions } from "./customactions";
import { Features } from "./features";
import { WebSettings } from "./websettings";
import { Navigation } from "./navigation";
import { Lists } from "./lists";
import { Files } from "./files";
import { ClientSidePages } from "./clientsidepages";
import { PropertyBagEntries } from "./propertybagentries";
import { IProvisioningConfig} from "../provisioningconfig";
import { SiteFields } from "./sitefields";
import { ContentTypes } from "./contenttypes";

export const DefaultHandlerMap = (config: IProvisioningConfig): TypedHash<HandlerBase> => ({
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
});

export const DefaultHandlerSort: TypedHash<number> = {
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
};

