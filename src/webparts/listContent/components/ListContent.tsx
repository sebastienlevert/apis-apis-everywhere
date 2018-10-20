import * as React from "react";
import styles from "./ListContent.module.scss";
import { IListContentProps } from "./IListContentProps";
import IListContentState from "./IListContentState";
import { escape } from "@microsoft/sp-lodash-subset";
import { ListView, IViewField, SelectionMode, GroupOrder, IGrouping } from "@pnp/spfx-controls-react/lib/ListView";
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import MockDataService from "./../../../services/MockDataService";
import { ISessionItem } from "../../../models/ISessionItem";

import {
  Spinner,
  SpinnerSize
} from "office-ui-fabric-react/lib/Spinner";
import { Label } from "office-ui-fabric-react/lib/Label";
import { MessageBar, MessageBarType } from "office-ui-fabric-react/lib/MessageBar";
import { DefaultButton, PrimaryButton } from "office-ui-fabric-react/lib/Button";
import { Dialog, DialogType, DialogFooter } from "office-ui-fabric-react/lib/Dialog";
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { ChoiceGroup } from "office-ui-fabric-react/lib/ChoiceGroup";

export default class ListContent extends React.Component<IListContentProps, IListContentState> {
  private _isConfigured: boolean = false;
  private _itemTitle: string = undefined;
  private _itemDescription: string = undefined;
  private _itemLevel: string = undefined;

  constructor(props: IListContentProps) {
    super(props);

    this.state = {
      sessionItems: [],
      selectedItems: []
    };

    this._onConfigure = this._onConfigure.bind(this);
    this._getSelection = this._getSelection.bind(this);
    this._onTitleChanged = this._onTitleChanged.bind(this);
    this._onDescriptionChanged = this._onDescriptionChanged.bind(this);
    this._onLevelChanged = this._onLevelChanged.bind(this);
  }

  public render(): React.ReactElement<IListContentProps> {
    const viewFields: IViewField[] = [
      {
        name: "id",
        displayName: "ID",
        maxWidth: 25,
        minWidth: 25,
        sorting: true
      },
      {
        name: "title",
        displayName: "Title",
        maxWidth: 100,
        minWidth: 100,
        sorting: true
      },
      {
        name: "description",
        displayName: "Description",
        maxWidth: 200,
        minWidth: 100,
        sorting: false
      },
      {
        name: "level",
        displayName: "Level",
        maxWidth: 50,
        minWidth: 50,
        sorting: true
      }
    ];

    if(this.state.isLoading) {
      return (
        <div>
          <Spinner size={ SpinnerSize.large } label={"Loading data using " + this.props.dataService.getTitle()} ariaLive="assertive" />
        </div>
      );
    } else {
      if (this._isValid()) {
        return (
          <div>
            <MessageBar>Currently displaying content using <span className="ms-fontWeight-semibold">{this.props.dataService.getTitle()}</span></MessageBar>
            <div>&nbsp;</div>
            <ListView
              selectionMode={SelectionMode.multiple}
              selection={this._getSelection}
              items={this.state.sessionItems}
              viewFields={viewFields} />

            <div>&nbsp;</div>
            <DefaultButton
              onClick={ this._showDialog }
              text="Create a new Session"/>

            <DefaultButton
              disabled={ this.state.selectedItems.length <= 0}
              onClick={ this._deleteSelected }
              text="Delete Selected Items"/>

            <Dialog
              hidden={ this.state.hideDialog }
              onDismiss={ this._closeDialog }
              dialogContentProps={ {
                type: DialogType.largeHeader,
                title: "Create a new Session",
                subText: "Use this form to create a new Session"
              } }
              modalProps={ {
                isBlocking: true,
                containerClassName: "ms-dialogMainOverride"
              } }
            >
              <div className="ms-Grid">
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-lg12">
                    <TextField
                      placeholder="Title"
                      required={ true }
                      onBlur={ this._onTitleChanged }
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-lg12">
                    <TextField
                      placeholder="Description"
                      required={ true }
                      onBlur={ this._onDescriptionChanged }
                    />
                  </div>
                </div>
                <div className="ms-Grid-row">
                  <div className="ms-Grid-col ms-lg12">
                    <ChoiceGroup
                      options={ [
                        {
                          key: "100",
                          text: "100"
                        },
                        {
                          key: "200",
                          text: "200",
                          checked: true
                        },
                        {
                          key: "300",
                          text: "300"
                        },
                        {
                          key: "400",
                          text: "400"
                        }
                      ] }
                      onChange={ this._onLevelChanged }
                    />
                  </div>
                </div>
              </div>
              <DialogFooter>
                <PrimaryButton onClick={ this._saveDialog } text="Save" />
                <DefaultButton onClick={ this._closeDialog } text="Cancel" />
              </DialogFooter>
            </Dialog>
          </div>
        );
      } else {
        return (
          <Placeholder
            iconName="Edit"
            iconText="Configure your web part"
            description="Please configure the web part."
            buttonLabel="Configure"
            onConfigure={this._onConfigure} />
        );
      }
    }
  }

  public componentDidMount(): void {
    if(this._isValid()) {
      this.setState({
        isLoading: true
      });
      this._getItems();
    }
  }

  public componentDidUpdate(previousProps: IListContentProps, previousState: IListContentState): void {
    if(this._isValid() &&
      (previousProps && (this.props.list !== previousProps.list || this.props.dataService.getTitle() !== previousProps.dataService.getTitle()))) {
      this.setState({
        isLoading: true
      });
      this._getItems();
    }
  }

  private _showDialog = (): void => {
    this.setState({ hideDialog: false });
  }

  private _closeDialog = (): void => {
    this.setState({ hideDialog: true });
  }

  private _deleteSelected = (): void => {
    for(let item of this.state.selectedItems) {
      
      if (!window.confirm(`Are you sure you want to delete the item with id ${item.id}?`)) {
        return;
      }

      this.props.dataService.deleteItem(item.id).then(() => {
        this._getItems();
      });
    }
  }

  private _saveDialog = (): void => {
    this._addItem().then(() => {
      this.setState({ hideDialog: true });
    });
  }

  private _isValid(): Boolean {
    return this.props.dataService && this.props.dataService.isConfigured();
  }

  private _onConfigure(): void {
    this.props.context.propertyPane.open();
  }

  private _onLevelChanged(event: any, option: any):void {
    this._itemLevel = option.key;
  }

  private _onTitleChanged(event: any):void {
    this._itemTitle = event.target.value;
  }

  private _onDescriptionChanged(event: any):void {
    this._itemDescription = event.target.value;
  }

  private _getItems(): void {
    this.props.dataService.getItems(this.props.context).then(sessionItems => {
      this.setState({
        sessionItems: sessionItems,
        isLoading: false
      });
    });
  }

  private _addItem(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      let sessionItem: ISessionItem = {
        title: this._itemTitle,
        description: this._itemDescription,
        level: this._itemLevel
      };

      this.props.dataService.addItem(sessionItem).then(() => {
        this._getItems();
        resolve();
      });
    });
  }

  private _getSelection(items: any[]): void {
    this.setState({
      selectedItems: items
    });
  }
}
