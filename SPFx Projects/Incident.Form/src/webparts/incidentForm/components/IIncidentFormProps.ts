import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IIncidentFormProps {
  title: string;
  showTitle: boolean;
  listId: string;
  notificationRecipients: string;
  successMessage: string;
  errorMessage: string;
  fullNameInfo: string;
  occupationRoleInfo: string;
  addressOfIncidentInfo: string;
  locationNameOfIncidentInfo: string;
  dateOfIncidentInfo: string;
  timeOfIncidentInfo: string;
  dateReportedInfo: string;
  timeReportedInfo: string;
  supervisorInfo: string;
  witnessInfo: string;
  typeOfIncidentInfo: string;
  whoWasInjuredInfo: string;
  equipmentDamagedInfo: string;
  equipmentIdentifyingNumberInfo: string;
  damagedEquipmentDescriptionInfo: string;
  descriptionOfIncidentInfo: string;
  immediateActionsTakenInfo: string;
  attachmentsInfo: string;
  isDarkTheme: boolean;
  userDisplayName: string;
  context: WebPartContext;
}
