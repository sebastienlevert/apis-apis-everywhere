import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import IDataService from "./IDataService";

import { IWebPartContext } from "@microsoft/sp-webpart-base";

export default class MockDataService implements IDataService {
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
    return "Mock API";
  }

  public isConfigured(): boolean {
    return true;
  }

  public getItems(context: IWebPartContext): Promise<IHelpDeskItem[]> {
    return new Promise<IHelpDeskItem[]>((resolve, reject) => {
      setTimeout(() => resolve([
        {
          id : 1,
          title : "That doesn't work",
          description : "When I do that, it doesn't work",
          level : "Low",
          status: "Open",
          resolution: "Do this and it will work!",
          assignedTo: "Sébastien Levert",
        }
      ]), 300);
    });
  }
}
