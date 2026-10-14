import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IFeedbackFormProps {
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
  context: WebPartContext;
  feedbackListId: string;
  title: string;
  showTitle: boolean;
  subTitle: string;
  showSubTitle: boolean;
  successMessage: string;
  errorMessage: string;
  homeUrl: string;
}
