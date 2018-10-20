import { ISessionItem } from "./../models/ISessionItem";
import IDataService from "./IDataService";

import { IWebPartContext } from "@microsoft/sp-webpart-base";

export default class MockDataService implements IDataService {
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
    return "Mock API";
  }

  public isConfigured(): boolean {
    return true;
  }

  public getItems(context: IWebPartContext): Promise<ISessionItem[]> {
    return new Promise<ISessionItem[]>((resolve, reject) => {
      setTimeout(() => resolve([
        {
          id : 1,
          title : "Session #1",
          description : "Awesome session!",
          level : "100"
        }
      ]), 300);
    });
  }
}
