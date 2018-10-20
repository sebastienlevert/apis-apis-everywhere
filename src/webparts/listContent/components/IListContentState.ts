import { ISessionItem } from './../../../models/ISessionItem';

export default interface IListContentState {
  sessionItems?: ISessionItem[];
  isLoading?: boolean;
  hideDialog?: boolean;
  selectedItems?: any[];
}
