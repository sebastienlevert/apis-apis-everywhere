import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from "@microsoft/sp-webpart-base";

import * as strings from "ListContentWebPartStrings";
import ListContent from "./components/ListContent";
import { IListContentProps } from "./components/IListContentProps";
import MockDataService from "../../../lib/services/MockDataService";
import SharePointDataService from "../../services/SharePointDataService";
import { PropertyFieldListPicker, PropertyFieldListPickerOrderBy } from "@pnp/spfx-property-controls/lib/PropertyFieldListPicker";
import IDataService from "../../../lib/services/IDataService";
import { setup as pnpSetup } from "@pnp/common";
import PnPJSCoreDataService from "../../services/PnPJSCoreDataService";
import MicrosoftGraphDataService from "../../services/MicrosoftGraphDataService";
import CustomDataService from "../../services/CustomDataService";
import CustomSecuredDataService from "../../services/CustomSecuredDataService";
import SharePointSearchDataService from "../../services/SharePointSearchDataService";

export enum API {
  None,
  Mock,
  SharePointREST,
  SharePointSearch,
  SharePointPnPJSCore,
  MicrosoftGraph,
  MicrosoftGraphPnPJSCore,
  Custom,
  CustomAzureAD
}

export interface IListContentWebPartProps {
  list: string;
  selectedApi: API;
}

export default class ListContentWebPart extends BaseClientSideWebPart<IListContentWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IListContentProps> = React.createElement(
      ListContent,
      {
        context: this.context,
        dataService: this.getDataService(),
        list: this.properties.list
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      pnpSetup({
        spfxContext: this.context
      });
    });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneDropdown("selectedApi", {
                  label: "Select an API",
                  options: [
                    { key: API.None, text: ""},
                    { key: API.Mock, text: "Mock API"},
                    { key: API.SharePointREST, text: "SharePoint REST API (Direct)" },
                    { key: API.SharePointSearch, text: "SharePoint REST API (Search)" },
                    { key: API.SharePointPnPJSCore, text: "SharePoint REST API (PnP JS Core)" },
                    { key: API.MicrosoftGraph, text: "Microsoft Graph"},
                    { key: API.MicrosoftGraphPnPJSCore, text: "Microsoft Graph (PnP JS Core)" },
                    { key: API.Custom, text: "Custom API"},
                    { key: API.CustomAzureAD, text: "Custom Azure AD Secured API"}
                  ],
                  selectedKey: API.None
                }),
                PropertyFieldListPicker("list", {
                  label: "Select a list",
                  selectedList: this.properties.list,
                  includeHidden: false,
                  orderBy: PropertyFieldListPickerOrderBy.Title,
                  disabled: this.properties.selectedApi == API.Mock,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: "listPickerFieldId",
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private getDataService(): IDataService {
    switch(this.properties.selectedApi) {
      case API.Mock:
        return new MockDataService(this.context, "");
      case API.SharePointREST:
        return new SharePointDataService(this.context, this.properties.list);
      case API.SharePointSearch:
        return new SharePointSearchDataService(this.context, this.properties.list);
      case API.SharePointPnPJSCore:
        return new PnPJSCoreDataService(this.context, this.properties.list);
      case API.MicrosoftGraph:
        return new MicrosoftGraphDataService(this.context, this.properties.list);
      case API.Custom:
        return new CustomDataService(this.context, this.properties.list);
      case API.CustomAzureAD:
        return new CustomSecuredDataService(this.context, this.properties.list);
      case API.None:
      default:
        return null;
    }
  }
}
