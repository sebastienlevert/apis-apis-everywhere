import { WebPartContext } from '@microsoft/sp-webpart-base';
import IDataService from './../../../services/IDataService';

export interface IListContentProps {
  dataService: IDataService;
  context: WebPartContext;
  list?: string;
}
