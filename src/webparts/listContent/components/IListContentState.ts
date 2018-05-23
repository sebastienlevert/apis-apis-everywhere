import { IHelpDeskItem } from './../../../models/IHelpDeskItem';

export default interface IListContentState {
  helpDeskItems?: IHelpDeskItem[];
  isLoading?: boolean;
  hideDialog?: boolean;
  selectedItems?: any[];
}
