import { IDropdownOption } from "@fluentui/react/lib/components/Dropdown/Dropdown.types";

export interface IFeedbackFormState {
  feedback: string;
  feedbackType: string;
  submittedBy: string;
  feedbackTypeOptions: IDropdownOption[];
  submitted: boolean;
  feedbackError: string;
  feedbackTypeError: string;
  showDialog: boolean;
  dialogMessage: string;
  dialogIsError: boolean;
}
