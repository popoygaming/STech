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

import * as strings from 'FeedbackFormWebPartStrings';
import FeedbackForm from './components/FeedbackForm';
import { IFeedbackFormProps } from './components/IFeedbackFormProps';
import { DEFAULT_TITLE, DEFAULT_SUBTITLE, DEFAULT_SUCCESS_MESSAGE, DEFAULT_ERROR_MESSAGE, DEFAULT_HOME_URL } from './constants';

export interface IFeedbackFormWebPartProps {
  feedbackListId: string;
  title: string;
  showTitle: boolean;
  subTitle: string;
  showSubTitle: boolean;
  successMessage: string;
  errorMessage: string;
  homeUrl: string;
}


export default class FeedbackFormWebPart extends BaseClientSideWebPart<IFeedbackFormWebPartProps> {

  // provide sensible defaults for new properties
  protected onInit(): Promise<void> {
    if (this.properties.showTitle === undefined) {
      this.properties.showTitle = true;
    }
    if (this.properties.showSubTitle === undefined) {
      this.properties.showSubTitle = true;
    }
    if (!this.properties.title) {
      this.properties.title = DEFAULT_TITLE;
    }
    if (!this.properties.subTitle) {
      this.properties.subTitle = DEFAULT_SUBTITLE;
    }
    if (!this.properties.successMessage) {
      this.properties.successMessage = DEFAULT_SUCCESS_MESSAGE;
    }
    if (!this.properties.errorMessage) {
      this.properties.errorMessage = DEFAULT_ERROR_MESSAGE;
    }
    if (!this.properties.homeUrl) {
      this.properties.homeUrl = DEFAULT_HOME_URL;
    }
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const title = this.properties.title || DEFAULT_TITLE;
    const subTitle = this.properties.subTitle || DEFAULT_SUBTITLE;
    const successMessage = this.properties.successMessage || DEFAULT_SUCCESS_MESSAGE;
    const errorMessage = this.properties.errorMessage || DEFAULT_ERROR_MESSAGE;

    const element: React.ReactElement<IFeedbackFormProps> = React.createElement(
      FeedbackForm,
      {
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        feedbackListId: this.properties.feedbackListId,
        title: title,
        showTitle: this.properties.showTitle,
        subTitle: subTitle,
        showSubTitle: this.properties.showSubTitle,
        successMessage: successMessage,
        errorMessage: errorMessage,
        homeUrl: this.properties.homeUrl || DEFAULT_HOME_URL
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('feedbackListId', {            
                    label: "Feedback List ID",
                    description: "Enter the GUID of the Intranet Feedback list"
                  }),
                PropertyPaneTextField('title', {
                  label: "Title",
                  value: this.properties.title || DEFAULT_TITLE
                }),
                PropertyPaneToggle('showTitle', {
                  label: "Show title",
                  checked: this.properties.showTitle !== false
                }),
                PropertyPaneTextField('subTitle', {
                  label: "Subtitle",
                  value: this.properties.subTitle || DEFAULT_SUBTITLE
                }),
                PropertyPaneToggle('showSubTitle', {
                  label: "Show subtitle",
                  checked: this.properties.showSubTitle !== false
                }),
                PropertyPaneTextField('successMessage', {
                  label: "Success message",
                  value: this.properties.successMessage || DEFAULT_SUCCESS_MESSAGE
                }),
                PropertyPaneTextField('errorMessage', {
                  label: "Error message",
                  value: this.properties.errorMessage || DEFAULT_ERROR_MESSAGE
                }),
                PropertyPaneTextField('homeUrl', {
                  label: "Home URL",
                  value: this.properties.homeUrl || DEFAULT_HOME_URL
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
