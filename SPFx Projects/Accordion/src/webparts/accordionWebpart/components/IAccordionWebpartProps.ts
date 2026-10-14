import { IAccordionItem } from './models/IAccordionItem';

export interface IAccordionWebpartProps {
  items: IAccordionItem[];
  textSize: 'small' | 'medium' | 'large';
  isEditMode: boolean;
  onInlineSave: (items: IAccordionItem[]) => void;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}
