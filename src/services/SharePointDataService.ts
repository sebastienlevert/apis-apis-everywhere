import { ISessionItem } from "./../models/ISessionItem";
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

  public getItems(context: WebPartContext): Promise<ISessionItem[]> {
    return new Promise<ISessionItem[]>((resolve, reject) => {
      context.spHttpClient
        .get( `${this._webPartContext.pageContext.web.absoluteUrl}/_api/web/lists/GetById('${this._listId}')/items` +
              `?$select=*`, SPHttpClient.configurations.v1)
        .then(res => res.json())
        .then(res => {
          let sessionItems:ISessionItem[] = [];

          for(let sessionItem of res.value) {
            sessionItems.push(this.buildSessionItem(sessionItem));
          }

          resolve(sessionItems);
        })
        .catch(err => console.log(err));
    });
  }

  public addItem(item: ISessionItem): Promise<void> {
    const currentWebUrl: string = this._webPartContext.pageContext.web.absoluteUrl;

    console.log(item); 

    return new Promise<void>((resolve, reject) => {
      this.getListItemEntityTypeName().then((listItemEntityTypeName: string): Promise<SPHttpClientResponse> => {
        const body: string = JSON.stringify({
          "__metadata": {
            "type": listItemEntityTypeName
          },
          "Title": item.title,
          "SessionDescription": item.description,
          "SessionLevel": item.level
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
      .then((newitem: any): void => {
        resolve();
      });
    });
  }

  public deleteItem(id: number): Promise<void> {
    const currentWebUrl: string = this._webPartContext.pageContext.web.absoluteUrl;
    return new Promise<void>((resolve, reject) => {

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

  protected buildSessionItem(sessionItem: any): ISessionItem {
    return {
      id: sessionItem.Id,
      title: sessionItem.Title,
      description: sessionItem.SessionDescription,
      level: sessionItem.SessionLevel
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
