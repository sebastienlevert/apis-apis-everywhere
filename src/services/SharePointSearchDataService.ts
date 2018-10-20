import { ISessionItem } from "./../models/ISessionItem";
import IDataService from './IDataService';

import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, HttpClientResponse } from '@microsoft/sp-http';


export default class SharePointSearchDataService implements IDataService {
  public deleteItem(id: number): Promise<void> {
    throw new Error("Method not implemented.");
  }
  public addItem(item: ISessionItem): Promise<void> {
    throw new Error("Method not implemented.");
  }
  private _webPartContext: IWebPartContext;
  private _listId: string;

  constructor(webPartContext: IWebPartContext, listId: string) {
    this._webPartContext = webPartContext;
    this._listId = listId;
  }

  public getTitle(): string {
    return "SharePoint REST API (Search)";
  }

  public isConfigured(): boolean {
    return Boolean(this._listId);
  }

  public getItems(context: IWebPartContext): Promise<ISessionItem[]> {
    return new Promise<ISessionItem[]>((resolve, reject) => {
      context.spHttpClient
        .get(`${this._webPartContext.pageContext.web.absoluteUrl}/_api/search/query?` +
              `querytext='ContentTypeId:0x0100A829DCD06F34504690C156727F8AEAFE* AND ListID:${this._listId}'` +
              `&selectproperties='ListItemID,Title,SessionDescriptionOWSMTXT,SessionLevelOWSCHCS'` +
              `&orderby='ListItemID asc'`, SPHttpClient.configurations.v1, {
          headers: {
            "odata-version": "3.0"
          }
        })
        .then(res => res.json())
        .then(res => {
          let sessionItems:ISessionItem[] = [];

          if(res.PrimaryQueryResult) {
            for(var row of res.PrimaryQueryResult.RelevantResults.Table.Rows) {
              sessionItems.push(this.buildSessionItem(row));
            }
          }

          resolve(sessionItems);
        })
        .catch(err => console.log(err));
    });
  }

  protected buildSessionItem(helpDeskSearchRow: any): ISessionItem {
    return {
      id: this.getResultValueByKey('ListItemID', helpDeskSearchRow),
      title: this.getResultValueByKey('Title', helpDeskSearchRow),
      description: this.getResultValueByKey('SessionDescriptionOWSMTXT', helpDeskSearchRow),
      level: this.getResultValueByKey("SessionLevelOWSCHCS", helpDeskSearchRow)
    };
  }

  private getResultValueByKey(key, searchResult): any {
    for(var property of searchResult.Cells) {
      if (property.Key.toLowerCase() == key.toLowerCase()) {
        return property.Value;
      }
    }

    return null;
  }
}
