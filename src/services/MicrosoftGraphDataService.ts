import { ISessionItem } from "./../models/ISessionItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClient } from "@microsoft/sp-http";

export default class MicrosoftGraphDataService implements IDataService {

  private _webPartContext: WebPartContext;
  private _listId: string;
  private _client: MSGraphClient;

  constructor(webPartContext: WebPartContext, listId: string) {
    this._webPartContext = webPartContext;
    this._listId = listId;
  }

  public getTitle(): string {
    return "Microsoft Graph";
  }

  public isConfigured(): boolean {
    return Boolean(this._listId);
  }

  public getItems(context: WebPartContext): Promise<ISessionItem[]> {

    let graphUrl: string =  `https://graph.microsoft.com/v1.0` +
                            `/sites/${this.getCurrentSiteCollectionGraphId()}` +
                            `/lists/${this._listId}` +
                            `/items?expand=fields(${this.getFieldsToExpand()})`;

    return new Promise<ISessionItem[]>((resolve, reject) => {
      this._webPartContext.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
        client
          .api(graphUrl)
          .get((error, response: any) => {
            if (error) {
              console.error(error);
              return;
            }

            let sessionItems:ISessionItem[] = [];

            for(let sessionItem of response.value) {
              sessionItems.push(this.buildSessionItem(sessionItem));
            }

            resolve(sessionItems);
          });
      });
    });
  }

  public addItem(item: ISessionItem): Promise<void> {
    let graphUrl: string =  `https://graph.microsoft.com/v1.0` +
                            `/sites/${this.getCurrentSiteCollectionGraphId()}` +
                            `/lists/${this._listId}` +
                            `/items`;

    return new Promise<void>((resolve, reject) => {
      const body: any = {
        "fields" : {
          "Title": item.title,
          "SessionDescription": item.description,
          "SessionLevel": item.level
        }
      };

      this._webPartContext.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
        client
          .api(graphUrl)
          .post(body, (error, response: any) => {
            if (error) {
              console.error(error);
              return;
            }

            resolve();
        });
      });
    });
  }

  public deleteItem(id: number): Promise<void> {
    let graphUrl: string =  `https://graph.microsoft.com/v1.0` +
                            `/sites/${this.getCurrentSiteCollectionGraphId()}` +
                            `/lists/${this._listId}` +
                            `/items/${id}`;

    return new Promise<void>((resolve, reject) => {
      this._webPartContext.msGraphClientFactory.getClient().then((client: MSGraphClient): void => {
        client
          .api(graphUrl)
          .delete((error, response: any) => {
            if (error) {
              console.error(error);
              return;
            }

            resolve();
          });
      });
    });
  }

  private getFieldsToExpand(): string {
    return encodeURIComponent("$select=id,Title,SessionDescription,SessionLevel");
  }

  private buildSessionItem(sessionGraphItem: any): ISessionItem {
    return {
      id: sessionGraphItem.id,
      title: sessionGraphItem.fields.Title,
      description: sessionGraphItem.fields.SessionDescription,
      level: sessionGraphItem.fields.SessionLevel
    };
  }

  private getCurrentSiteCollectionGraphId(): string {
    return `${window.location.hostname},${this._webPartContext.pageContext.site.id},${this._webPartContext.pageContext.web.id}`;
  }
}
