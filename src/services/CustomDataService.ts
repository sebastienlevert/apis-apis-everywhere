import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { HttpClient } from "@microsoft/sp-http";

export default class CustomDataService implements IDataService {

  deleteItem(id: number): Promise<void> {
    throw new Error("Method not implemented.");
  }
  addItem(item: IHelpDeskItem): Promise<void> {
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

  public getItems(context: WebPartContext): Promise<IHelpDeskItem[]> {
    return new Promise<IHelpDeskItem[]>((resolve, reject) => {
      context.httpClient
        .get("https://apis-apis-everywhere.azurewebsites.net/api/GetHelpDeskItems", HttpClient.configurations.v1)
        .then(res => res.json())
        .then(res => {
          let helpDeskItems:IHelpDeskItem[] = [];

          for(let helpDeskListItem of res) {
            helpDeskItems.push(this.buildHelpDeskItem(helpDeskListItem));
          }

          resolve(helpDeskItems);
        })
        .catch(err => console.log(err));
    });
  }

  protected buildHelpDeskItem(helpDeskListItem: any): IHelpDeskItem {
    return {
      id: helpDeskListItem.Id,
      title: helpDeskListItem.Title,
      description: helpDeskListItem.HelpDeskDescription,
      level: helpDeskListItem.HelpDeskLevel,
      status: helpDeskListItem.HelpDeskStatus,
      resolution: helpDeskListItem.HelpDeskResolution,
      assignedTo: helpDeskListItem.HelpDeskAssignedTo
    };
  }
}
