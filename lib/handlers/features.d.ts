import { HandlerBase } from "./handlerbase";
import { IFeature } from "../schema";
import { Web } from "@pnp/sp";
import { IProvisioningConfig } from "../provisioningconfig";
/**
 * Describes the Features Object Handler
 */
export declare class Features extends HandlerBase {
    /**
     * Creates a new instance of the ObjectFeatures class
     *
     * @param {IProvisioningConfig} config Provisioning config
     */
    constructor(config: IProvisioningConfig);
    /**
     * Provisioning features
     *
     * @param {Web} web The web
     * @param {Array<IFeature>} features The features to provision
     */
    ProvisionObjects(web: Web, features: IFeature[]): Promise<void>;
}
