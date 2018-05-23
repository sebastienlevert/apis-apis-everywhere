import { IHelpDeskItem } from "./../models/IHelpDeskItem";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export default interface IDataService {
  getTitle(): string;
  isConfigured(): boolean;
  getItems(context: WebPartContext): Promise<IHelpDeskItem[]>;
  addItem(item: IHelpDeskItem): Promise<void>;
  //updateItem(context: WebPartContext, item: IHelpDeskItem): Promise<void>;
  deleteItem(id: number): Promise<void>;
}
