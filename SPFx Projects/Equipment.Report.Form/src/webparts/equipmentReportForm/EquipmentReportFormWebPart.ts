import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'EquipmentReportFormWebPartStrings';
import EquipmentReportForm from './components/EquipmentReportForm';
import { IEquipmentReportFormProps } from './components/IEquipmentReportFormProps';

export interface IEquipmentReportFormWebPartProps {
  title: string;
  showTitle: boolean;
  listId: string;
  notificationReceiver: string;
  successMessage: string;
  errorMessage: string;
}

export default class EquipmentReportFormWebPart extends BaseClientSideWebPart<IEquipmentReportFormWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IEquipmentReportFormProps> = React.createElement(
      EquipmentReportForm,
      {
        title: this.properties.title,
        showTitle: this.properties.showTitle,
        listId: this.properties.listId,
        notificationReceiver: this.properties.notificationReceiver,
        successMessage: this.properties.successMessage,
        errorMessage: this.properties.errorMessage,
        isDarkTheme: this._isDarkTheme,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                PropertyPaneToggle('showTitle', {
                  label: 'Show Web Part Title',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneTextField('title', {
                  label: 'Web Part Title'
                }),
                PropertyPaneTextField('listId', {
                  label: 'List ID (GUID)'
                }),
                PropertyPaneTextField('notificationReceiver', {
                  label: 'Notification Receiver'
                }),
                PropertyPaneTextField('successMessage', {
                  label: 'Success Message'
                }),
                PropertyPaneTextField('errorMessage', {
                  label: 'Error Message'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
