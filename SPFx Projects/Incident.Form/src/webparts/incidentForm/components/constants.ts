import { IDropdownOption } from '@fluentui/react';

export const INCIDENT_TYPE_OPTIONS: IDropdownOption[] = [
  { key: 'Near Miss', text: 'Near Miss' },
  { key: 'Minor Injury', text: 'Minor Injury' },
  { key: 'Medical Treatment Injury (MTI)', text: 'Medical Treatment Injury (MTI)' },
  { key: 'Lost Time Injury (LTI)', text: 'Lost Time Injury (LTI)' },
  { key: 'Equipment Damage', text: 'Equipment Damage' },
  { key: 'Process loss', text: 'Process loss' }
];

export const FIELD_LABELS = {
  fullName: 'Full Name',
  occupationRole: 'Occupation / Role',
  addressOfIncident: 'Address Of Incident',
  locationNameOfIncident: 'Location Name Of Incident',
  dateOfIncident: 'Date Of Incident',
  timeOfIncident: 'Time Of Incident',
  dateReported: 'Date Reported',
  timeReported: 'Time Reported',
  supervisor: 'Supervisor',
  witness: 'Witness',
  typeOfIncident: 'Type Of Incident',
  whoWasInjured: 'Who was injured?',
  equipmentDamaged: 'Equipment Damaged',
  equipmentIdentifyingNumber: 'Equipment Identifying Number',
  damagedEquipmentDescription: 'Damaged Equipment description',
  descriptionOfIncident: 'Description of Incident / What Occurred?',
  immediateActionsTaken: 'Immediate Actions Taken',
  attachments: 'Attachments'
};

export const FIELD_ARIA_LABELS = {
  fullName: 'Full Name',
  occupationRole: 'Occupation or Role',
  addressOfIncident: 'Address Of Incident',
  locationNameOfIncident: 'Location Name Of Incident',
  dateOfIncident: 'Date Of Incident',
  timeOfIncident: 'Time Of Incident',
  dateReported: 'Date Reported',
  timeReported: 'Time Reported',
  supervisor: 'Supervisor',
  witness: 'Witness',
  typeOfIncident: 'Type Of Incident',
  whoWasInjured: 'Who was injured',
  equipmentDamaged: 'Equipment Damaged',
  equipmentIdentifyingNumber: 'Equipment Identifying Number',
  damagedEquipmentDescription: 'Damaged Equipment description',
  descriptionOfIncident: 'Description of Incident or What Occurred',
  immediateActionsTaken: 'Immediate Actions Taken',
  attachments: 'Attachments and Photos'
};

export const FORM_STATIC_TEXT = {
  okButton: 'OK',
  datePlaceholder: 'Select a date',
  incidentTypePlaceholder: 'Select an incident type',
  whoWasInjuredPlaceholder: 'Full Name',
  descriptionPlaceholder: 'Describe all relevant background information, sequence of events leading to the incident, what occurred that precipitated the accident, the mechanism of accident and who did you contact after).',
  immediateActionsPlaceholder: 'First aid applied, area cordoned off).',
  attachmentsLabelNote: '(photos of the scene or injury)',
  filesSelectedSuffix: 'file(s) selected',
  submitButton: 'Submit Incident',
  submittingAriaLabel: 'Submitting incident',
  submittingSpinnerLabel: 'Submitting incident...',
  successFallbackMessage: 'Incident submitted successfully.',
  errorFallbackMessage: 'An error occurred while submitting the incident.',
  dialogValidationTitle: 'Validation Message',
  dialogSuccessTitle: 'Success',
  dialogErrorTitle: 'Error',
  dialogDefaultTitle: 'Message',
  listIdValidationMessage: 'Please configure the List ID (GUID) in the web part properties.',
  incidentDateFutureValidationMessage: 'Date of incident cannot be in the future.',
  reportedDateFutureValidationMessage: 'Date reported cannot be in the future.',
  reportedBeforeIncidentValidationMessage: 'Date reported cannot be earlier than the date of incident.'
};
