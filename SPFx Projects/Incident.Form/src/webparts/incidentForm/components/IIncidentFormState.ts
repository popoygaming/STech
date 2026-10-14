export type SubmissionStatus = 'none' | 'success' | 'error';

export interface IIncidentFormState {
  fullName: string;
  occupationRole: string;
  addressOfIncident: string;
  locationNameOfIncident: string;
  incidentDate: Date | undefined;
  incidentTime: string;
  dateReported: Date | undefined;
  timeReported: string;
  supervisor: string;
  witness: string;
  incidentType: string;
  whoWasInjured: string;
  equipmentDamaged: string;
  equipmentIdentifyingNumber: string;
  damagedEquipmentDescription: string;
  incidentDescription: string;
  immediateActions: string;
  attachments: File[];
  isSubmitting: boolean;
  status: SubmissionStatus;
  statusMessage: string;
  validationMessage: string;
}
