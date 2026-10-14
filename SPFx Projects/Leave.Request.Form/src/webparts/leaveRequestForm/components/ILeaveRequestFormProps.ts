import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILeaveRequestFormProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
  listId?: string;
  approvers: string[];
  successMessage: string;
  sickLeaveNote?: string;
  showTitle: boolean;
}
