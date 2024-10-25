import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory, MSGraphClientV3 } from "@microsoft/sp-http";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-types";

export interface IUserGraphService {
    getCurrentUserDetails(): Promise<MicrosoftGraph.User>;
}

export class UserGraphService implements IUserGraphService {

    public static readonly serviceKey: ServiceKey<IUserGraphService> = ServiceKey.create<IUserGraphService>("oneorg:IUserGraphService", UserGraphService);

    private _msGraphClientFactory: MSGraphClientFactory;

    constructor(serviceScope: ServiceScope){
        serviceScope.whenFinished(() => {
            this._msGraphClientFactory = serviceScope.consume(MSGraphClientFactory.serviceKey);
        });
    }

    // ==================================================================

    public getCurrentUserDetails(): Promise<MicrosoftGraph.User> {
        return new Promise<MicrosoftGraph.User>((resolve, reject) => {
            this._msGraphClientFactory
                .getClient("3")
                .then((_msGraphClient: MSGraphClientV3) => {
                    _msGraphClient
                        .api("/me")
                        .get((error: any, user: MicrosoftGraph.User, rawResponse?: any) => {
                            //console.log("Error:" + error);
                            resolve(user);
                    })
                .catch((error: any) => {
                    // error
                    reject(error);
                });
            })
        });
    }

}