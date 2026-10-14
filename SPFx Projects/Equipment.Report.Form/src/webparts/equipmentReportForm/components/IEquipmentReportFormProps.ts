import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IEquipmentReportFormProps {
  title: string;
  showTitle: boolean;
  listId: string;
  notificationReceiver: string;
  successMessage: string;
  errorMessage: string;
  isDarkTheme: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
