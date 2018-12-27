import * as React from "react";
import * as ReactDom from "react-dom";
import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from "@microsoft/sp-webpart-base";
import { Dialog } from "@microsoft/sp-dialog";
import MockHttpClient from "./components/MockHttpClient";
import {
  SPHttpClient,
  SPHttpClientResponse
} from "@microsoft/sp-http";
import {
  Environment,
  EnvironmentType
} from "@microsoft/sp-core-library";

import * as strings from "InformaticaFiveQuestionsWebPartStrings";
import InformaticaFiveQuestions from "./components/InformaticaFiveQuestions";
import { IInformaticaFiveQuestionsProps } from "./components/IInformaticaFiveQuestionsProps";

import { Provider } from "react-redux";
import { store } from "./components/redux";
import AppContainer from "./components/AppContainer";

export interface IInformaticaFiveQuestionsWebPartProps {
  WebPartTitle: string;
  ListName: string;
}

export default class InformaticaFiveQuestionsWebPart extends BaseClientSideWebPart<IInformaticaFiveQuestionsWebPartProps> {

  // options for listname dropdown
  private _lists: IPropertyPaneDropdownOption[];
  private _listsDropdownDisabled: boolean = true;
  /**
   * populate listname dropdown
   */
  private loadLists(): Promise<any> {
    return new Promise((resolve, reject) => {
      if (Environment.type === EnvironmentType.Local) {
        this._getMockLists().then((response) => {
          resolve(response);
        });
      } else if (Environment.type === EnvironmentType.SharePoint ||
        Environment.type === EnvironmentType.ClassicSharePoint) {
        this._getLists()
          .then((response) => {
            resolve(response);
          });
      }
    });
  }

  /**
   * get fake data for listname dropdown
   */
  private _getMockLists(): Promise<any> {
    return MockHttpClient.getLists()
      .then((data: any) => {
        return data;
      }) as Promise<any>;
  }
  /**
   * get data for listname dropdown
   */
  private _getLists(): Promise<any> {
    // get data for list name options
    return this.context.spHttpClient.get(
      this.context.pageContext.web.absoluteUrl +
      `/_api/web/Lists?$select=Id,Title&$filter=Hidden eq false and BaseTemplate eq 100&$Orderby=Title desc`,
      SPHttpClient.configurations.v1
    )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      }).then((data) => {
        return data.value.map((item) => {
          return {
            key: item.Id || 0,
            text: item.Title || ""
          };
        });
      }).catch((err) => {
        Dialog.alert("Error occurred while fetching data");
      });
  }

  public render(): void {

    const element: React.ReactElement<any> = React.createElement(
      AppContainer,
      {
        WebPartTitle: this.properties.WebPartTitle,
        context: this.context,
        ListName: this.properties.ListName,
      }
    );

    const element1: any = React.createElement(
      Provider, { store: store }, element
    );

    ReactDom.render(element1, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("WebPartTitle", {
                  label: strings.WebPartTitleFieldLabel
                }),
                PropertyPaneDropdown("ListName", {
                  label: strings.ListNameFieldLabel,
                  options: this._lists,
                  disabled: this._listsDropdownDisabled
                }),
              ]
            }
          ]
        }
      ]
    };
  }

  /**
   * on property pane configuration start event
   */
  protected onPropertyPaneConfigurationStart(): void {
    this._listsDropdownDisabled = !this._lists;
    if (this._lists) {
      return;
    }
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, "_lists");

    this.loadLists()
      .then((listOptions: IPropertyPaneDropdownOption[]): void => {
        this._lists = listOptions;
        this._listsDropdownDisabled = false;
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.context.propertyPane.refresh();
        this.onDispose();
        this.render();
      });
  }

  /**
   * property pane field changed event
   * @param propertyPath
   * @param oldValue
   * @param newValue
   */
  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    this.context.statusRenderer.clearLoadingIndicator(this.domElement);
    this.context.propertyPane.refresh();
    this.onDispose();
    this.render();
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }
}
