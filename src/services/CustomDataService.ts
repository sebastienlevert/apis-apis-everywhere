import { ISessionItem } from "./../models/ISessionItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient } from "@microsoft/sp-http";

export default class CustomDataService implements IDataService {

  public deleteItem(id: number): Promise<void> {
    throw new Error("Method not implemented.");
  }
  public addItem(item: ISessionItem): Promise<void> {
    throw new Error("Method not implemented.");
  }
  protected _webPartContext: WebPartContext;
  protected _listId: string;

  constructor(webPartContext: WebPartContext, listId: string) {
    this._webPartContext = webPartContext;
    this._listId = listId;
  }

  public getTitle(): string {
    return "Custom API";
  }

  public isConfigured(): boolean {
    return Boolean(this._listId);
  }

  public getItems(context: WebPartContext): Promise<ISessionItem[]> {
    return new Promise<ISessionItem[]>((resolve, reject) => {
      context.httpClient
        .get("https://apis-apis-everywhere.azurewebsites.net/api/GetHelpDeskItems", HttpClient.configurations.v1)
        .then(res => res.json())
        .then(res => {
          let sessionItems:ISessionItem[] = [];

          for(let sessionItem of res) {
            sessionItems.push(this.buildSessionItem(sessionItem));
          }

          resolve(sessionItems);
        })
        .catch(err => console.log(err));
    });
  }

  protected buildSessionItem(helpDeskListItem: any): ISessionItem {
    return {
      id: helpDeskListItem.Id,
      title: helpDeskListItem.Title,
      description: helpDeskListItem.SessionDescription,
      level: helpDeskListItem.SessionLevel
    };
  }
}
