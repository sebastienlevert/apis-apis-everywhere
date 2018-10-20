import { ISessionItem } from "./../models/ISessionItem";
import IDataService from "./IDataService";

import { WebPartContext } from "@microsoft/sp-webpart-base";
import { sp } from "@pnp/sp";
import SharePointDataService from "./SharePointDataService";

export default class PnPJSCoreDataService extends SharePointDataService {


  public getTitle(): string {
    return "SharePoint REST API (PnP JS Core)";
  }

  public getItems(context: WebPartContext): Promise<ISessionItem[]> {
    return new Promise<ISessionItem[]>((resolve, reject) => {

      sp.web.lists.getById(this._listId).items
        .select("*").getAll().then((sessionItems: any[]) => {
        let sessions:ISessionItem[] = [];

        for(let sessionItem of sessionItems) {
          sessions.push(this.buildSessionItem(sessionItem));
        }

        resolve(sessions);
      });

    });
  }
}
