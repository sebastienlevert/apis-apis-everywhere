import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient, HttpClientResponse, ISPHttpClientOptions, SPHttpClientResponse } from "@microsoft/sp-http";


export default class SharePointDataService implements IDataService {

  protected _webPartContext: WebPartContext;
  protected _listId: string;
  private listItemEntityTypeName: string = undefined;

  constructor(webPartContext: WebPartContext, listId: string) {
    this._webPartContext = webPartContext;
    this._listId = listId;
  }

  public getTitle(): string {
    return "SharePoint REST API (Direct)";
  }

  public isConfigured(): boolean {
    return Boolean(this._listId);
  }

  public getItems(context: WebPartContext): Promise<IHelpDeskItem[]> {
    return new Promise<IHelpDeskItem[]>((resolve, reject) => {
      context.spHttpClient
        .get( `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/GetById('${this._listId}')/items` +
              `?$select=*,HelpDeskAssignedTo/Title&$expand=HelpDeskAssignedTo`, SPHttpClient.configurations.v1)
        .then(res => res.json())
        .then(res => {
          let helpDeskItems:IHelpDeskItem[] = [];

          for(let helpDeskListItem of res.value) {
            helpDeskItems.push(this.buildHelpDeskItem(helpDeskListItem));
          }

          resolve(helpDeskItems);
        })
        .catch(err => console.log(err));
    });
  }

  public addItem(item: IHelpDeskItem): Promise<void> {
    const currentWebUrl: string = this._webPartContext.pageContext.web.absoluteUrl;

    return new Promise<void>((resolve, reject) => {
      this.getListItemEntityTypeName().then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
        const body: string = JSON.stringify({
          "__metadata": {
            "type": listItemEntityTypeName
          },
          "Title": item.title,
          "HelpDeskDescription": item.description,
          "HelpDeskLevel": item.level
        });

        return this._webPartContext.spHttpClient.post(`${currentWebUrl}/_api/web/lists/GetById('${this._listId}')/items`,
          SPHttpClient.configurations.v1,
          {
            headers: {
              "Accept": "application/json;odata=nometadata",
              "Content-type": "application/json;odata=verbose",
              "odata-version": ""
            },
            body: body
          });
      })
      .then((response: SPHttpClientResponse): Promise<any> => {
        return response.json();
      })
      .then((item: any): void => {
        resolve();
      });
    });
  }

  public deleteItem(id: number): Promise<void> {
    const currentWebUrl: string = this._webPartContext.pageContext.web.absoluteUrl;
    return new Promise<void>((resolve, reject) => {

    if (!window.confirm(`Are you sure you want to delete the item with id ${id}?`)) {
      return;
    }

    return this._webPartContext.spHttpClient.post(`${currentWebUrl}/_api/web/lists/GetById('${this._listId}')/items(${id})`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Accept": "application/json;odata=nometadata",
          "Content-type": "application/json;odata=verbose",
          "odata-version": "",
          "IF-MATCH": "*",
          "X-HTTP-Method": "DELETE"
        }
      })
      .then((response: SPHttpClientResponse): void => {
        resolve();
      });
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
      assignedTo: helpDeskListItem.HelpDeskAssignedTo ? helpDeskListItem.HelpDeskAssignedTo.Title : null
    };
  }

  private getListItemEntityTypeName(): Promise<string> {
    return new Promise<string>((resolve: (listItemEntityTypeName: string) => void, reject: (error: any) => void): void => {
      if (this.listItemEntityTypeName) {
        resolve(this.listItemEntityTypeName);
        return;
      }

      this._webPartContext.spHttpClient.get(
        `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists` +
        `/GetById('${this._listId}')?$select=ListItemEntityTypeFullName`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "odata-version": ""
          }
        })
        .then((response: SPHttpClientResponse): Promise<{ ListItemEntityTypeFullName: string }> => {
          return response.json();
        }, (error: any): void => {
          reject(error);
        })
        .then((response: { ListItemEntityTypeFullName: string }): void => {
          this.listItemEntityTypeName = response.ListItemEntityTypeFullName;
          resolve(this.listItemEntityTypeName);
        });
    });
  }
}
