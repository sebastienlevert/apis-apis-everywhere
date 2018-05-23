import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http";
import CustomDataService from "./CustomDataService";

export default class CustomSecuredDataService extends CustomDataService {

  private _client: AadHttpClient;

  constructor(webPartContext: WebPartContext, listId: string) {
    super(webPartContext, listId);
    this._client = new AadHttpClient(this._webPartContext.serviceScope, "834d48af-54a8-47b7-b1ba-67ec55608513");
  }

  public getTitle(): string {
    return "Custom Azure AD Secured API";
  }

  public getItems(context: WebPartContext): Promise<IHelpDeskItem[]> {

    return new Promise<IHelpDeskItem[]>((resolve, reject) => {
      let apiUrl: string = "https://secured-apis-apis-everywhere.azurewebsites.net/api/GetHelpDeskItems";
      this._client
        .get(apiUrl, AadHttpClient.configurations.v1)
        .then((res: HttpClientResponse): Promise<any> => {
          return res.json();
        })
        .then((res: any): void => {
          let helpDeskItems:IHelpDeskItem[] = [];

          for(let helpDeskListItem of res) {
            helpDeskItems.push(this.buildHelpDeskItem(helpDeskListItem));
          }

          resolve(helpDeskItems);
        }, (err: any): void => {
          console.error(err);
        });
    });
  }

}
