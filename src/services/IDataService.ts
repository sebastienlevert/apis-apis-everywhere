import { ISessionItem } from "./../models/ISessionItem";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export default interface IDataService {
  getTitle(): string;
  isConfigured(): boolean;
  getItems(context: WebPartContext): Promise<ISessionItem[]>;
  addItem(item: ISessionItem): Promise<void>;
  //updateItem(context: WebPartContext, item: ISessionItem): Promise<void>;
  deleteItem(id: number): Promise<void>;
}
