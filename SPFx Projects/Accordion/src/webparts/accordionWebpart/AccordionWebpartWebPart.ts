import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AccordionWebpartWebPartStrings';
import AccordionWebpart from './components/AccordionWebpart';
import { IAccordionWebpartProps } from './components/IAccordionWebpartProps';
import { DEFAULT_ACCORDION_ITEMS_JSON } from './components/constants/accordionDefaults';
import { getAccordionItems } from './components/helpers/accordionDataHelper';
import { IAccordionItem } from './components/models/IAccordionItem';

export interface IAccordionWebpartWebPartProps {
  accordionItemsJson: string;
  textSize: 'small' | 'medium' | 'large';
}

export default class AccordionWebpartWebPart extends BaseClientSideWebPart<IAccordionWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private onInlineSave = (items: IAccordionItem[]): void => {
    this.properties.accordionItemsJson = JSON.stringify(items, null, 2);

    if (this.context.propertyPane.isPropertyPaneOpen()) {
      this.context.propertyPane.refresh();
    }

    this.render();
  };

  public render(): void {
    const element: React.ReactElement<IAccordionWebpartProps> = React.createElement(
      AccordionWebpart,
      {
        items: getAccordionItems(this.properties.accordionItemsJson),
        textSize: this.properties.textSize,
        isEditMode: this.displayMode === DisplayMode.Edit,
        onInlineSave: this.onInlineSave,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    if (!this.properties.accordionItemsJson) {
      this.properties.accordionItemsJson = DEFAULT_ACCORDION_ITEMS_JSON;
    }

    if (!this.properties.textSize) {
      this.properties.textSize = 'medium';
    }

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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

  protected onDisplayModeChanged(_oldDisplayMode: DisplayMode): void {
    this.render();
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
                PropertyPaneTextField('accordionItemsJson', {
                  label: strings.AccordionItemsFieldLabel,
                  description: strings.AccordionItemsFieldDescription,
                  multiline: true,
                  resizable: true,
                  rows: 10
                }),
                PropertyPaneDropdown('textSize', {
                  label: strings.TextSizeFieldLabel,
                  options: [
                    { key: 'small', text: strings.TextSizeSmallLabel },
                    { key: 'medium', text: strings.TextSizeMediumLabel },
                    { key: 'large', text: strings.TextSizeLargeLabel }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
