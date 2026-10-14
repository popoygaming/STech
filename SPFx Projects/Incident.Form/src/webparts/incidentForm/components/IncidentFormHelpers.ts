import type { IIncidentFormProps } from './IIncidentFormProps';
import type { IIncidentFormState } from './IIncidentFormState';
import { FORM_STATIC_TEXT } from './constants';

export function getDateAtMidnight(value: Date): Date {
  return new Date(value.getFullYear(), value.getMonth(), value.getDate());
}

export function formatTwoDigits(value: number): string {
  return value < 10 ? `0${value}` : String(value);
}

export function getCurrentTimeValue(): string {
  const now: Date = new Date();
  const hours: string = formatTwoDigits(now.getHours());
  const minutes: string = formatTwoDigits(now.getMinutes());
  return `${hours}:${minutes}`;
}

export function clearFeedbackState(): Pick<IIncidentFormState, 'validationMessage' | 'status' | 'statusMessage'> {
  return {
    validationMessage: '',
    status: 'none',
    statusMessage: ''
  };
}

export function createInitialState(props: IIncidentFormProps): IIncidentFormState {
  return {
    fullName: props.userDisplayName || '',
    occupationRole: '',
    addressOfIncident: '',
    locationNameOfIncident: '',
    incidentDate: undefined,
    incidentTime: '',
    dateReported: getDateAtMidnight(new Date()),
    timeReported: getCurrentTimeValue(),
    supervisor: '',
    witness: '',
    incidentType: '',
    whoWasInjured: '',
    equipmentDamaged: '',
    equipmentIdentifyingNumber: '',
    damagedEquipmentDescription: '',
    incidentDescription: '',
    immediateActions: '',
    attachments: [],
    isSubmitting: false,
    status: 'none',
    statusMessage: '',
    validationMessage: ''
  };
}

export function getFieldInfo(props: IIncidentFormProps, fieldName: string): string {
  const infoByField: Record<string, string | undefined> = {
    fullName: props.fullNameInfo,
    occupationRole: props.occupationRoleInfo,
    addressOfIncident: props.addressOfIncidentInfo,
    locationNameOfIncident: props.locationNameOfIncidentInfo,
    dateOfIncident: props.dateOfIncidentInfo,
    timeOfIncident: props.timeOfIncidentInfo,
    dateReported: props.dateReportedInfo,
    timeReported: props.timeReportedInfo,
    supervisor: props.supervisorInfo,
    witness: props.witnessInfo,
    typeOfIncident: props.typeOfIncidentInfo,
    whoWasInjured: props.whoWasInjuredInfo,
    equipmentDamaged: props.equipmentDamagedInfo,
    equipmentIdentifyingNumber: props.equipmentIdentifyingNumberInfo,
    damagedEquipmentDescription: props.damagedEquipmentDescriptionInfo,
    descriptionOfIncident: props.descriptionOfIncidentInfo,
    immediateActionsTaken: props.immediateActionsTakenInfo,
    attachments: props.attachmentsInfo
  };

  return (infoByField[fieldName] || '').trim();
}

export function isDialogVisible(state: IIncidentFormState): boolean {
  return !!state.validationMessage || state.status !== 'none';
}

export function getDialogTitle(state: IIncidentFormState): string {
  if (state.validationMessage) {
    return FORM_STATIC_TEXT.dialogValidationTitle;
  }

  if (state.status === 'success') {
    return FORM_STATIC_TEXT.dialogSuccessTitle;
  }

  if (state.status === 'error') {
    return FORM_STATIC_TEXT.dialogErrorTitle;
  }

  return FORM_STATIC_TEXT.dialogDefaultTitle;
}

export function getDialogMessage(state: IIncidentFormState): string {
  if (state.validationMessage) {
    return state.validationMessage;
  }

  return state.statusMessage;
}

export function validateForm(state: IIncidentFormState, today: Date, hasConfiguredListId: boolean): string {
  if (!hasConfiguredListId) {
    return FORM_STATIC_TEXT.listIdValidationMessage;
  }

  if (state.incidentDate && state.dateReported) {
    const incidentDate: Date = getDateAtMidnight(state.incidentDate);
    const reportedDate: Date = getDateAtMidnight(state.dateReported);

    if (incidentDate > today) {
      return FORM_STATIC_TEXT.incidentDateFutureValidationMessage;
    }

    if (reportedDate > today) {
      return FORM_STATIC_TEXT.reportedDateFutureValidationMessage;
    }

    if (reportedDate < incidentDate) {
      return FORM_STATIC_TEXT.reportedBeforeIncidentValidationMessage;
    }
  } else {
    if (state.incidentDate) {
      const incidentDate: Date = getDateAtMidnight(state.incidentDate);
      if (incidentDate > today) {
        return FORM_STATIC_TEXT.incidentDateFutureValidationMessage;
      }
    }
    if (state.dateReported) {
      const reportedDate: Date = getDateAtMidnight(state.dateReported);
      if (reportedDate > today) {
        return FORM_STATIC_TEXT.reportedDateFutureValidationMessage;
      }
    }
  }

  return '';
}

export function getDateOnlyIsoString(value: Date | undefined): string {
  const normalizedDate: Date = getDateAtMidnight(value as Date);
  const year: number = normalizedDate.getFullYear();
  const monthNumber: number = normalizedDate.getMonth() + 1;
  const dayNumber: number = normalizedDate.getDate();
  const month: string = monthNumber < 10 ? `0${monthNumber}` : String(monthNumber);
  const day: string = dayNumber < 10 ? `0${dayNumber}` : String(dayNumber);
  return `${year}-${month}-${day}`;
}
