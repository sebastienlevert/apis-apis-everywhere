import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import IDataService from './IDataService';

import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, HttpClientResponse } from '@microsoft/sp-http';


export default class SharePointSearchDataService implements IDataService {
  deleteItem(id: number): Promise<void> {
    throw new Error("Method not implemented.");
  }
  addItem(item: IHelpDeskItem): Promise<void> {
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

  public getItems(context: IWebPartContext): Promise<IHelpDeskItem[]> {
    return new Promise<IHelpDeskItem[]>((resolve, reject) => {
      context.spHttpClient
        .get(`${this._webPartContext.pageContext.web.absoluteUrl}/_api/search/query?` +
              `querytext='ContentTypeId:0x010006746609E953604CACB6A08F020BF357* AND ListID:${this._listId}'` +
              `&selectproperties='ListItemID,Title,HelpDeskDescriptionOWSMTXT,HelpDeskLevelOWSCHCS,HelpDeskStatusOWSCHCS,HelpDeskAssignedToOWSUSER'` +
              `&orderby='ListItemID asc'`, SPHttpClient.configurations.v1, {
          headers: {
            "odata-version": "3.0"
          }
        })
        .then(res => res.json())
        .then(res => {
          let helpDeskItems:IHelpDeskItem[] = [];

          if(res.PrimaryQueryResult) {
            for(var row of res.PrimaryQueryResult.RelevantResults.Table.Rows) {
              helpDeskItems.push(this.buildHelpDeskItem(row));
            }
          }

          resolve(helpDeskItems);
        })
        .catch(err => console.log(err));
    });
  }

  protected buildHelpDeskItem(helpDeskSearchRow: any): IHelpDeskItem {
    return {
      id: this.getResultValueByKey('ListItemID', helpDeskSearchRow),
      title: this.getResultValueByKey('Title', helpDeskSearchRow),
      description: this.getResultValueByKey('HelpDeskDescriptionOWSMTXT', helpDeskSearchRow),
      level: this.getResultValueByKey("HelpDeskLevelOWSCHCS", helpDeskSearchRow),
      status: this.getResultValueByKey("HelpDeskStatusOWSCHCS", helpDeskSearchRow),
      resolution: this.getResultValueByKey("HelpDeskResolutionOWSTEXT", helpDeskSearchRow),
      assignedTo: this.getResultValueByKey("HelpDeskAssignedToOWSUSER", helpDeskSearchRow) ? this.getResultValueByKey("HelpDeskAssignedToOWSUSER", helpDeskSearchRow).split(" | ")[1] : null
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
