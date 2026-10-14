import { DisplayMode } from '@microsoft/sp-core-library';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue } from '@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker';
export interface ICalendarProps {
  title: string;
  siteUrl: string;
  list: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
  context: WebPartContext;
  eventStartDate:  IDateTimeFieldValue;
  eventEndDate: IDateTimeFieldValue;
  /** optional initial category filter from the web part property pane */
  selectedCategory?: string;
  /** callback invoked when category selection changes in the UI */
  onCategoryChanged?: (value: string) => void;
}
