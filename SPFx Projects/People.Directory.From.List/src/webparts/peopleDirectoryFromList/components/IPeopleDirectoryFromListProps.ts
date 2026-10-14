import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IPeopleDirectoryFromListProps {
  description: string;
  employeeListId: string;
  webPartTitle: string;
  showWebPartTitle: boolean;
  pageSize: number;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
