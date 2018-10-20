import { ISessionItem } from "./../models/ISessionItem";
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

  public getItems(context: WebPartContext): Promise<ISessionItem[]> {

    return new Promise<ISessionItem[]>((resolve, reject) => {
      let apiUrl: string = "https://secured-apis-apis-everywhere.azurewebsites.net/api/GetSessionItems";
      this._client
        .get(apiUrl, AadHttpClient.configurations.v1)
        .then((res: HttpClientResponse): Promise<any> => {
          return res.json();
        })
        .then((res: any): void => {
          let sessionItems:ISessionItem[] = [];

          for(let helpDeskListItem of res) {
            sessionItems.push(this.buildSessionItem(helpDeskListItem));
          }

          resolve(sessionItems);
        }, (err: any): void => {
          console.error(err);
        });
    });
  }

}
