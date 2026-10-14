import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneLabel

} from '@microsoft/sp-property-pane';
import PropertyPaneMultiSelect from './controls/PropertyPaneMultiSelect/PropertyPaneMultiSelect';

import * as strings from 'CalendarWebPartStrings';
import Calendar from './components/Calendar/Calendar';
import { ICalendarProps } from './components/Calendar/ICalendarProps';
import { PropertyFieldDateTimePicker, DateConvention, TimeConvention, IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';

import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { PropertyPaneHorizontalRule } from "@microsoft/sp-property-pane";

export interface ICalendarWebPartProps {
  title: string;
  siteUrl: string;
  list: string;
  eventStartDate: IDateTimeFieldValue ;
  eventEndDate: IDateTimeFieldValue;
  errorMessage: string;
  /** optional category filter; blank or undefined means all */
  category?: string;
}

import spservices from '../../services/spservices';
import * as moment from 'moment';
import { format } from '@uifabric/utilities';

export default class CalendarWebPart extends BaseClientSideWebPart<ICalendarWebPartProps> {

  private lists: IPropertyPaneDropdownOption[] = [];
  private categoriesOptions: IPropertyPaneDropdownOption[] = [];
  private listsDropdownDisabled: boolean = true;
  private spService: spservices = null;
  private errorMessage: string;

  public constructor() {
    super();

  }

  public render(): void {

    const element: React.ReactElement<ICalendarProps> = React.createElement(
      Calendar,
      {
        title: this.properties.title,
        siteUrl: this.properties.siteUrl,
        list: this.properties.list,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        },
        context: this.context,
        eventStartDate: this.properties.eventStartDate,
        eventEndDate: this.properties.eventEndDate,
        selectedCategory: this.properties.category,
        onCategoryChanged: (value: string) => {
          this.properties.category = value;
          // refresh property pane if open
          this.context.propertyPane.refresh();
          // re-render calendar to pick up prop change
          this.render();
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  // onInit
  public  async onInit(): Promise<void> {

    this.spService = new spservices(this.context);
    this.properties.siteUrl = this.properties.siteUrl ? this.properties.siteUrl : this.context.pageContext.site.absoluteUrl;
    if (!this.properties.eventStartDate){
      this.properties.eventStartDate = { value: moment().subtract(2,'years').startOf('month').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
    }
    if (!this.properties.eventEndDate){
      this.properties.eventEndDate = { value: moment().add(20,'years').endOf('month').toDate(), displayValue: moment().format('ddd MMM MM YYYY')};
    }
    if (this.properties.siteUrl && !this.properties.list) {
     const _lists = await this.loadLists();
     if ( _lists.length > 0 ){
      this.lists = _lists;
      this.properties.list = this.lists[0].key.toString();
     }
    }

    return Promise.resolve();
  }


  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /**
   *
   * @protected
   * @memberof CalendarWebPart
   */
  protected async onPropertyPaneConfigurationStart() {

    try {
      if (this.properties.siteUrl) {
        const _lists = await this.loadLists();
        this.lists = _lists;
        this.listsDropdownDisabled = false;
        // also reload category values for the selected list
        if (this.properties.list) {
          this.categoriesOptions = await this.loadCategories();
        }
        this.context.propertyPane.refresh();

      } else {
        this.lists = [];
        this.properties.list = '';
        this.categoriesOptions = [];
        this.listsDropdownDisabled = false;
        this.context.propertyPane.refresh();
      }

    } catch (error) {

    }
  }

  /**
   *
   * @private
   * @returns {Promise<IPropertyPaneDropdownOption[]>}
   * @memberof CalendarWebPart
   */
  private async loadLists(): Promise<IPropertyPaneDropdownOption[]> {
    const _lists: IPropertyPaneDropdownOption[] = [];
    try {
      const results = await this.spService.getSiteLists(this.properties.siteUrl);
      for (const list of results) {
        _lists.push({ key: list.Id, text: list.Title });
      }
      // push new item value
    } catch (error) {
      this.errorMessage =  `${error.message} -  please check if site url if valid.` ;
      this.context.propertyPane.refresh();
    }
    return _lists;
  }

  /**
   * Load choice field values for the "Category" column on the selected list
   */
  private async loadCategories(): Promise<IPropertyPaneDropdownOption[]> {
    const opts: IPropertyPaneDropdownOption[] = [];
    if (!this.properties.siteUrl || !this.properties.list) {
      return opts;
    }
    try {
      const categories: { key: string; text: string }[] =
        await this.spService.getChoiceFieldOptions(this.properties.siteUrl, this.properties.list, 'Category');
      for (const cat of categories) {
        opts.push({ key: cat.key, text: cat.text });
      }
    } catch (error) {
      // ignore, maybe list doesn't have field
    }
    return opts;
  }

  /**
   *
   *
   * @private
   * @param {string} date
   * @returns
   * @memberof CalendarWebPart
   */
  private onEventStartDateValidation(date:string){
    if (date && this.properties.eventEndDate.value){
      if (moment(date).isAfter(moment(this.properties.eventEndDate.value))){
        return strings.SartDateValidationMessage;
      }
    }
    return '';
  }

  /**
   *
   * @private
   * @param {string} date
   * @returns
   * @memberof CalendarWebPart
   */
  private onEventEndDateValidation(date:string){
    if (date && this.properties.eventEndDate.value){
      if (moment(date).isBefore( moment(this.properties.eventStartDate.value))){
        return strings.EnDateValidationMessage;
      }
    }
    return '';
  }
  /**
   *
   * @private
   * @param {string} value
   * @returns {Promise<string>}
   * @memberof CalendarWebPart
   */

  private onSiteUrlGetErrorMessage(value: string) {
    let returnValue: string = '';
    if (value) {
      returnValue = '';
    } else {
      const previousList: string = this.properties.list;
      const previousSiteUrl: string = this.properties.siteUrl;
      // reset selected item
      this.properties.list = undefined;
      this.properties.siteUrl = undefined;
      this.lists = [];
      this.listsDropdownDisabled = true;
      this.onPropertyPaneFieldChanged('list', previousList, this.properties.list);
      this.onPropertyPaneFieldChanged('siteUrl', previousSiteUrl, this.properties.siteUrl);
      this.context.propertyPane.refresh();
    }
    return returnValue;
  }

  /**
   *
   * @protected
   * @param {string} propertyPath
   * @param {string} oldValue
   * @param {string} newValue
   * @memberof CalendarWebPart
   */
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: string, newValue: string) {
    try {
      // reset any error
      this.properties.errorMessage = undefined;
      this.errorMessage = undefined;
      this.context.propertyPane.refresh();

      if (propertyPath === 'siteUrl' && newValue) {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        const _oldValue = this.properties.list;
        this.onPropertyPaneFieldChanged('list', _oldValue, this.properties.list);
        this.context.propertyPane.refresh();
        const _lists = await this.loadLists();
        this.lists = _lists;
        this.listsDropdownDisabled = false;
        this.properties.list = this.lists.length > 0 ? this.lists[0].key.toString() : undefined;
        // when siteUrl changes we need to reload categories as well
        if (this.properties.list) {
          this.categoriesOptions = await this.loadCategories();
        } else {
          this.categoriesOptions = [];
        }
        this.context.propertyPane.refresh();
        this.render();
      }
      else {
        super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
        // if the list dropdown changed, reload categories
        if (propertyPath === 'list' && newValue) {
          this.categoriesOptions = await this.loadCategories();
          // if the previously selected category no longer exists clear it
          if (this.properties.category && !this.categoriesOptions.find(o => o.key === this.properties.category)) {
            this.properties.category = undefined;
          }
          this.context.propertyPane.refresh();
        }
        // always rerender when any property changes, but make sure category change updates view immediately
        if (propertyPath === 'category') {
          this.render();
        }
      }
    } catch (error) {
      this.errorMessage =  `${error.message} -  please check if site url if valid.` ;
      this.context.propertyPane.refresh();
    }
  }
  /**
   *
   * @protected
   * @returns {IPropertyPaneConfiguration}
   * @memberof CalendarWebPart
   */
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
      // EndDate and Start Date defualt values

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
                PropertyPaneTextField('siteUrl', {
                  label: strings.SiteUrlFieldLabel,
                  onGetErrorMessage: this.onSiteUrlGetErrorMessage.bind(this),
                  value: this.context.pageContext.site.absoluteUrl,
                  deferredValidationTime: 1200,
                }),
                PropertyPaneDropdown('list', {
                  label: strings.ListFieldLabel,
                  options: this.lists,
                  disabled: this.listsDropdownDisabled,
                }),
                PropertyPaneLabel('eventStartDate', {
                  text: strings.eventSelectDatesLabel
                }),
                PropertyFieldDateTimePicker('eventStartDate', {
                  label: 'From',
                  initialDate: this.properties.eventStartDate,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage: this.onEventStartDateValidation,
                  deferredValidationTime: 0,
                  key: 'eventStartDateId'
                }),
                PropertyFieldDateTimePicker('eventEndDate', {
                  label: 'to',
                  initialDate:  this.properties.eventEndDate,
                  dateConvention: DateConvention.Date,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  onGetErrorMessage:  this.onEventEndDateValidation,
                  deferredValidationTime: 0,
                  key: 'eventEndDateId'
                }),
                PropertyPaneLabel('errorMessage', {
                  text:  this.errorMessage,
                }),
                // category filter (multi-select)
                PropertyPaneMultiSelect('category', {
                  label: strings.CategoryLabel,
                  options: [
                    ...this.categoriesOptions
                  ],
                  selectedKeys: this.properties.category ? this.properties.category.split(',') : [],
                  onChanged: (keys: string[]) => {
                    // save as comma separated list; empty array => blank => all
                    this.properties.category = keys && keys.length > 0 ? keys.join(',') : '';
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}