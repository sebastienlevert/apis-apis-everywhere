import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { MSGraphClient } from "@microsoft/sp-client-preview";

export default class MicrosoftGraphDataService implements IDataService {

  private _webPartContext: WebPartContext;
  private _listId: string;
  private _client: MSGraphClient;

  constructor(webPartContext: WebPartContext, listId: string) {
    this._webPartContext = webPartContext;
    this._client = this._webPartContext.serviceScope.consume(MSGraphClient.serviceKey);
    this._listId = listId;
  }

  public getTitle(): string {
    return "Microsoft Graph";
  }

  public isConfigured(): boolean {
    return Boolean(this._listId);
  }

  public getItems(context: WebPartContext): Promise<IHelpDeskItem[]> {

    let graphUrl: string =  `https://graph.microsoft.com/v1.0` +
                            `/sites/${this.getCurrentSiteCollectionGraphId()}` +
                            `/lists/${this._listId}` +
                            `/items?expand=fields(${this.getFieldsToExpand()})`;

    return new Promise<IHelpDeskItem[]>((resolve, reject) => {
      this._client
        .api(graphUrl)
        .get((error, response: any) => {
          if (error) {
            console.error(error);
            return;
          }

          let helpDeskItems:IHelpDeskItem[] = [];

          for(let helpDeskListItem of response.value) {
            helpDeskItems.push(this.buildHelpDeskItem(helpDeskListItem));
          }

          resolve(helpDeskItems);
        });
    });
  }

  public addItem(item: IHelpDeskItem): Promise<void> {
    let graphUrl: string =  `https://graph.microsoft.com/v1.0` +
                            `/sites/${this.getCurrentSiteCollectionGraphId()}` +
                            `/lists/${this._listId}` +
                            `/items`;

    return new Promise<void>((resolve, reject) => {
      const body: any = {
        "fields" : {
          "Title": item.title,
          "HelpDeskDescription": item.description,
          "HelpDeskLevel": item.level
        }
      };

      this._client
        .api(graphUrl)
        .post(body, (error, response: any) => {
          if (error) {
            console.error(error);
            return;
          }

          resolve();
        });
    });
  }

  public deleteItem(id: number): Promise<void> {
    let graphUrl: string =  `https://graph.microsoft.com/v1.0` +
                            `/sites/${this.getCurrentSiteCollectionGraphId()}` +
                            `/lists/${this._listId}` +
                            `/items/${id}`;

    return new Promise<void>((resolve, reject) => {
      this._client
        .api(graphUrl)
        .delete((error, response: any) => {
          if (error) {
            console.error(error);
            return;
          }

          resolve();
        });
    });
  }

  private getFieldsToExpand(): string {
    return encodeURIComponent("$select=id,Title,HelpDeskDescription,HelpDeskLevel,HelpDeskStatus,HelpDeskResolution,HelpDeskAssignedTo");
  }

  private buildHelpDeskItem(helpDeskGraphItem: any): IHelpDeskItem {
    return {
      id: helpDeskGraphItem.id,
      title: helpDeskGraphItem.fields.Title,
      description: helpDeskGraphItem.fields.HelpDeskDescription,
      level: helpDeskGraphItem.fields.HelpDeskLevel,
      status: helpDeskGraphItem.fields.HelpDeskStatus,
      resolution: helpDeskGraphItem.fields.HelpDeskResolution,
      assignedTo: helpDeskGraphItem.fields.HelpDeskAssignedTo
    };
  }

  private getCurrentSiteCollectionGraphId(): string {
    return `${window.location.hostname},${this._webPartContext.pageContext.site.id},${this._webPartContext.pageContext.web.id}`;
  }
}
