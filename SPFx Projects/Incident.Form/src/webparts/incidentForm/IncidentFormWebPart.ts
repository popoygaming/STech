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

import * as strings from 'IncidentFormWebPartStrings';
import IncidentForm from './components/IncidentForm';
import { IIncidentFormProps } from './components/IIncidentFormProps';
import { IIncidentFormWebPartProps } from './IIncidentFormWebPartProps';
export { IIncidentFormWebPartProps };

export default class IncidentFormWebPart extends BaseClientSideWebPart<IIncidentFormWebPartProps> {

  private _isDarkTheme: boolean = false;

  public render(): void {
    const element: React.ReactElement<IIncidentFormProps> = React.createElement(
      IncidentForm,
      {
        title: this.properties.title,
        showTitle: this.properties.showTitle,
        listId: this.properties.listId,
        notificationRecipients: this.properties.notificationRecipients,
        successMessage: this.properties.successMessage,
        errorMessage: this.properties.errorMessage,
        fullNameInfo: this.properties.fullNameInfo,
        occupationRoleInfo: this.properties.occupationRoleInfo,
        addressOfIncidentInfo: this.properties.addressOfIncidentInfo,
        locationNameOfIncidentInfo: this.properties.locationNameOfIncidentInfo,
        dateOfIncidentInfo: this.properties.dateOfIncidentInfo,
        timeOfIncidentInfo: this.properties.timeOfIncidentInfo,
        dateReportedInfo: this.properties.dateReportedInfo,
        timeReportedInfo: this.properties.timeReportedInfo,
        supervisorInfo: this.properties.supervisorInfo,
        witnessInfo: this.properties.witnessInfo,
        typeOfIncidentInfo: this.properties.typeOfIncidentInfo,
        whoWasInjuredInfo: this.properties.whoWasInjuredInfo,
        equipmentDamagedInfo: this.properties.equipmentDamagedInfo,
        equipmentIdentifyingNumberInfo: this.properties.equipmentIdentifyingNumberInfo,
        damagedEquipmentDescriptionInfo: this.properties.damagedEquipmentDescriptionInfo,
        descriptionOfIncidentInfo: this.properties.descriptionOfIncidentInfo,
        immediateActionsTakenInfo: this.properties.immediateActionsTakenInfo,
        attachmentsInfo: this.properties.attachmentsInfo,
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
                  label: 'Show title',
                  onText: 'Show',
                  offText: 'Hide'
                }),
                PropertyPaneTextField('title', {
                  label: 'Title'
                }),
                PropertyPaneTextField('listId', {
                  label: 'List ID (GUID)'
                }),
                PropertyPaneTextField('notificationRecipients', {
                  label: 'Notification Recipients',
                  description: 'Comma-separated email addresses.'
                }),
                PropertyPaneTextField('successMessage', {
                  label: 'Success message'
                }),
                PropertyPaneTextField('errorMessage', {
                  label: 'Error message'
                })
              ]
            },
            {
              groupName: 'Field Information',
              groupFields: [
                PropertyPaneTextField('fullNameInfo', {
                  label: 'Full Name info text'
                }),
                PropertyPaneTextField('occupationRoleInfo', {
                  label: 'Occupation / Role info text'
                }),
                PropertyPaneTextField('addressOfIncidentInfo', {
                  label: 'Address Of Incident info text'
                }),
                PropertyPaneTextField('locationNameOfIncidentInfo', {
                  label: 'Location Name Of Incident info text'
                }),
                PropertyPaneTextField('dateOfIncidentInfo', {
                  label: 'Date Of Incident info text'
                }),
                PropertyPaneTextField('timeOfIncidentInfo', {
                  label: 'Time Of Incident info text'
                }),
                PropertyPaneTextField('dateReportedInfo', {
                  label: 'Date Reported info text'
                }),
                PropertyPaneTextField('timeReportedInfo', {
                  label: 'Time Reported info text'
                }),
                PropertyPaneTextField('supervisorInfo', {
                  label: 'Supervisor info text'
                }),
                PropertyPaneTextField('witnessInfo', {
                  label: 'Witness info text'
                }),
                PropertyPaneTextField('typeOfIncidentInfo', {
                  label: 'Type Of Incident info text'
                }),
                PropertyPaneTextField('whoWasInjuredInfo', {
                  label: 'Who was injured info text'
                }),
                PropertyPaneTextField('equipmentDamagedInfo', {
                  label: 'Equipment Damaged info text'
                }),
                PropertyPaneTextField('equipmentIdentifyingNumberInfo', {
                  label: 'Equipment Identifying Number info text'
                }),
                PropertyPaneTextField('damagedEquipmentDescriptionInfo', {
                  label: 'Damaged Equipment description info text',
                  multiline: true,
                  resizable: false
                }),
                PropertyPaneTextField('descriptionOfIncidentInfo', {
                  label: 'Description of Incident info text',
                  multiline: true,
                  resizable: false
                }),
                PropertyPaneTextField('immediateActionsTakenInfo', {
                  label: 'Immediate Actions Taken info text',
                  multiline: true,
                  resizable: false
                }),
                PropertyPaneTextField('attachmentsInfo', {
                  label: 'Attachments info text'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
