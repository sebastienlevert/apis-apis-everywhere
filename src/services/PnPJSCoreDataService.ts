import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import SharePointDataService from "./SharePointDataService";

export default class PnPJSCoreDataService extends SharePointDataService {


  public getTitle(): string {
    return "SharePoint REST API (PnP JS Core)";
  }

  public getItems(context: WebPartContext): Promise<IHelpDeskItem[]> {
    return new Promise<IHelpDeskItem[]>((resolve, reject) => {

      sp.web.lists.getById(this._listId).items
        .select("*", "HelpDeskAssignedTo/Title")
        .expand("HelpDeskAssignedTo").getAll().then((sessionItems: any[]) => {
        let helpDeskItems:IHelpDeskItem[] = [];

        for(let helpDeskListItem of sessionItems) {
          helpDeskItems.push(this.buildHelpDeskItem(helpDeskListItem));
        }

        resolve(helpDeskItems);
      });

    });
  }
}
