import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface ICustomBannerProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;

  companyMotto?: string;
  backgroundColor?: string;
  fontColor?: string;
}
