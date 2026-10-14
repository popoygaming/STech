import * as React from 'react';
import { DatePicker, DayOfWeek, DefaultButton, Dialog, DialogFooter, DialogType, Dropdown, Icon,
  IDropdownOption, Label, PrimaryButton, Spinner, SpinnerSize, Stack, TextField, TooltipHost
} from '@fluentui/react';
import styles from './IncidentForm.module.scss';
import type { IIncidentFormProps } from './IIncidentFormProps';
import type { IIncidentFormState } from './IIncidentFormState';
import SpService from '../services/SpService';
import { IIncidentListItemPayload } from '../services/ISpServiceInterfaces';
import { FIELD_ARIA_LABELS, FIELD_LABELS, FORM_STATIC_TEXT, INCIDENT_TYPE_OPTIONS } from './constants';
import { clearFeedbackState, createInitialState, getDateAtMidnight, getDateOnlyIsoString, getDialogMessage,
  getDialogTitle, getFieldInfo, isDialogVisible, validateForm
} from './IncidentFormHelpers';

export default class IncidentForm extends React.Component<IIncidentFormProps, IIncidentFormState> {
  private _spService: SpService;
  private readonly _today: Date = getDateAtMidnight(new Date());

  public constructor(props: IIncidentFormProps) {
    super(props);

    this._spService = new SpService(props.context, props.listId);
    this.state = createInitialState(props);
  }

  public componentDidUpdate(prevProps: IIncidentFormProps): void {
    if (prevProps.listId !== this.props.listId) {
      this._spService = new SpService(this.props.context, this.props.listId);
    }
  }

  public render(): React.ReactElement<IIncidentFormProps> {
    const state: IIncidentFormState = this.state;

    return (
      <div className={styles.incidentForm} aria-busy={state.isSubmitting}>
        {this.props.showTitle && (
          <h2 className={styles.title}>{this.props.title}</h2>
        )}

        <Dialog
          hidden={!isDialogVisible(state)}
          onDismiss={this._dismissDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: getDialogTitle(state),
            subText: getDialogMessage(state)
          }}
          modalProps={{ isBlocking: false }}
        >
          <DialogFooter>
            <DefaultButton text={FORM_STATIC_TEXT.okButton} onClick={this._dismissDialog} />
          </DialogFooter>
        </Dialog>

        <form onSubmit={this._onSubmit} noValidate>
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} className={styles.grid}>
              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.fullName, getFieldInfo(this.props, 'fullName'), false)}
                  <TextField
                    value={state.fullName}
                    onChange={this._onFullNameChange}
                    ariaLabel={FIELD_ARIA_LABELS.fullName}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.occupationRole, getFieldInfo(this.props, 'occupationRole'), false)}
                  <TextField
                    value={state.occupationRole}
                    onChange={this._onOccupationRoleChange}
                    ariaLabel={FIELD_ARIA_LABELS.occupationRole}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.addressOfIncident, getFieldInfo(this.props, 'addressOfIncident'), false)}
                  <TextField
                    value={state.addressOfIncident}
                    onChange={this._onAddressOfIncidentChange}
                    ariaLabel={FIELD_ARIA_LABELS.addressOfIncident}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.locationNameOfIncident, getFieldInfo(this.props, 'locationNameOfIncident'), false)}
                  <TextField
                    value={state.locationNameOfIncident}
                    onChange={this._onLocationNameOfIncidentChange}
                    ariaLabel={FIELD_ARIA_LABELS.locationNameOfIncident}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.dateOfIncident, getFieldInfo(this.props, 'dateOfIncident'), false)}
                  <DatePicker
                    firstDayOfWeek={DayOfWeek.Sunday}
                    placeholder={FORM_STATIC_TEXT.datePlaceholder}
                    value={state.incidentDate}
                    onSelectDate={this._onIncidentDateChange}
                    ariaLabel={FIELD_ARIA_LABELS.dateOfIncident}
                    maxDate={this._today}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.timeOfIncident, getFieldInfo(this.props, 'timeOfIncident'), false)}
                  <TextField
                    type="time"
                    value={state.incidentTime}
                    onChange={this._onIncidentTimeChange}
                    ariaLabel={FIELD_ARIA_LABELS.timeOfIncident}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.dateReported, getFieldInfo(this.props, 'dateReported'), false)}
                  <DatePicker
                    firstDayOfWeek={DayOfWeek.Sunday}
                    placeholder={FORM_STATIC_TEXT.datePlaceholder}
                    value={state.dateReported}
                    onSelectDate={this._onDateReportedChange}
                    ariaLabel={FIELD_ARIA_LABELS.dateReported}
                    maxDate={this._today}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.timeReported, getFieldInfo(this.props, 'timeReported'), false)}
                  <TextField
                    type="time"
                    value={state.timeReported}
                    onChange={this._onTimeReportedChange}
                    ariaLabel={FIELD_ARIA_LABELS.timeReported}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.supervisor, getFieldInfo(this.props, 'supervisor'), false)}
                  <TextField
                    value={state.supervisor}
                    onChange={this._onSupervisorChange}
                    ariaLabel={FIELD_ARIA_LABELS.supervisor}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.witness, getFieldInfo(this.props, 'witness'), false, '(if any)')}
                  <TextField
                    value={state.witness}
                    onChange={this._onWitnessChange}
                    ariaLabel={FIELD_ARIA_LABELS.witness}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.typeOfIncident, getFieldInfo(this.props, 'typeOfIncident'), false)}
                  <Dropdown
                    options={INCIDENT_TYPE_OPTIONS}
                    selectedKey={state.incidentType || undefined}
                    onChange={this._onIncidentTypeChange}
                    placeholder={FORM_STATIC_TEXT.incidentTypePlaceholder}
                    ariaLabel={FIELD_ARIA_LABELS.typeOfIncident}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.whoWasInjured, getFieldInfo(this.props, 'whoWasInjured'), false)}
                  <TextField
                    placeholder={FORM_STATIC_TEXT.whoWasInjuredPlaceholder}
                    value={state.whoWasInjured}
                    onChange={this._onWhoWasInjuredChange}
                    ariaLabel={FIELD_ARIA_LABELS.whoWasInjured}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.equipmentDamaged, getFieldInfo(this.props, 'equipmentDamaged'), false, '(if applicable)')}
                  <TextField
                    value={state.equipmentDamaged}
                    onChange={this._onEquipmentDamagedChange}
                    ariaLabel={FIELD_ARIA_LABELS.equipmentDamaged}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.equipmentIdentifyingNumber, getFieldInfo(this.props, 'equipmentIdentifyingNumber'), false)}
                  <TextField
                    value={state.equipmentIdentifyingNumber}
                    onChange={this._onEquipmentIdentifyingNumberChange}
                    ariaLabel={FIELD_ARIA_LABELS.equipmentIdentifyingNumber}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.fullWidth}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.damagedEquipmentDescription, getFieldInfo(this.props, 'damagedEquipmentDescription'), false)}
                  <TextField
                    multiline={true}
                    rows={3}
                    value={state.damagedEquipmentDescription}
                    onChange={this._onDamagedEquipmentDescriptionChange}
                    ariaLabel={FIELD_ARIA_LABELS.damagedEquipmentDescription}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.fullWidth}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.descriptionOfIncident, getFieldInfo(this.props, 'descriptionOfIncident'), false)}
                  <TextField
                    multiline={true}
                    placeholder={FORM_STATIC_TEXT.descriptionPlaceholder}
                    rows={5}
                    value={state.incidentDescription}
                    onChange={this._onDescriptionChange}
                    ariaLabel={FIELD_ARIA_LABELS.descriptionOfIncident}
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.fullWidth}>
                <div className={styles.fieldControl}>
                  {this._renderInfoLabel(FIELD_LABELS.immediateActionsTaken, getFieldInfo(this.props, 'immediateActionsTaken'), false)}
                  <TextField
                    multiline={true}
                    placeholder={FORM_STATIC_TEXT.immediateActionsPlaceholder}
                    rows={4}
                    value={state.immediateActions}
                    onChange={this._onImmediateActionsChange}
                    ariaLabel={FIELD_ARIA_LABELS.immediateActionsTaken}
                  />
                </div>
              </Stack.Item>
            </Stack>

            <div>
              {this._renderInfoLabel(FIELD_LABELS.attachments, getFieldInfo(this.props, 'attachments'), false)}
              <div className={styles.attachmentLabelNote}>{FORM_STATIC_TEXT.attachmentsLabelNote}</div>
              <input
                id="incidentAttachmentInput"
                className={styles.fileInput}
                type="file"
                multiple={true}
                onChange={this._onAttachmentsChange}
                aria-label={FIELD_ARIA_LABELS.attachments}
              />
              {state.attachments.length > 0 && (
                <div className={styles.attachmentText}>{state.attachments.length} {FORM_STATIC_TEXT.filesSelectedSuffix}</div>
              )}
            </div>

            <PrimaryButton text={FORM_STATIC_TEXT.submitButton} type="submit" disabled={state.isSubmitting} />
          </Stack>
        </form>

        {state.isSubmitting && (
          <div className={styles.submittingOverlay} role="status" aria-live="polite" aria-label={FORM_STATIC_TEXT.submittingAriaLabel}>
            <div className={styles.submittingSpinner}>
              <Spinner label={FORM_STATIC_TEXT.submittingSpinnerLabel} size={SpinnerSize.large} />
            </div>
          </div>
        )}
      </div>
    );
  }

  private _renderInfoLabel(label: string, infoText: string, required: boolean, italicSuffix?: string): React.ReactElement {
    return (
      <div className={styles.labelRow}>
        <Label required={required} className={styles.fieldLabel}>
          {label}
          {italicSuffix ? <> <span className={styles.labelSuffix}>{italicSuffix}</span></> : null}
        </Label>
        <TooltipHost content={infoText} calloutProps={{ gapSpace: 8 }}>
          <Icon iconName="Info" className={styles.infoIcon} aria-label={`${label} information`} />
        </TooltipHost>
      </div>
    );
  }

  private readonly _onIncidentDateChange = (date: Date | null | undefined): void => {
    this.setState({ incidentDate: date || undefined, ...clearFeedbackState() });
  };

  private readonly _onDateReportedChange = (date: Date | null | undefined): void => {
    this.setState({ dateReported: date || undefined, ...clearFeedbackState() });
  };

  private readonly _onFullNameChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ fullName: value || '', ...clearFeedbackState() });
  };

  private readonly _onOccupationRoleChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ occupationRole: value || '', ...clearFeedbackState() });
  };

  private readonly _onAddressOfIncidentChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ addressOfIncident: value || '', ...clearFeedbackState() });
  };

  private readonly _onLocationNameOfIncidentChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ locationNameOfIncident: value || '', ...clearFeedbackState() });
  };

  private readonly _onIncidentTimeChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ incidentTime: value || '', ...clearFeedbackState() });
  };

  private readonly _onTimeReportedChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ timeReported: value || '', ...clearFeedbackState() });
  };

  private readonly _onSupervisorChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ supervisor: value || '', ...clearFeedbackState() });
  };

  private readonly _onWitnessChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ witness: value || '', ...clearFeedbackState() });
  };

  private readonly _onIncidentTypeChange = (_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void => {
    this.setState({ incidentType: option ? String(option.key) : '', ...clearFeedbackState() });
  };

  private readonly _onWhoWasInjuredChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ whoWasInjured: value || '', ...clearFeedbackState() });
  };

  private readonly _onEquipmentDamagedChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ equipmentDamaged: value || '', ...clearFeedbackState() });
  };

  private readonly _onEquipmentIdentifyingNumberChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ equipmentIdentifyingNumber: value || '', ...clearFeedbackState() });
  };

  private readonly _onDamagedEquipmentDescriptionChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ damagedEquipmentDescription: value || '', ...clearFeedbackState() });
  };

  private readonly _onDescriptionChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ incidentDescription: value || '', ...clearFeedbackState() });
  };

  private readonly _onImmediateActionsChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ immediateActions: value || '', ...clearFeedbackState() });
  };

  private readonly _onAttachmentsChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files: File[] = [];
    const inputFiles: FileList | null = event.target.files;

    if (inputFiles) {
      for (let i: number = 0; i < inputFiles.length; i++) {
        files.push(inputFiles.item(i) as File);
      }
    }

    this.setState({ attachments: files, ...clearFeedbackState() });
  };

  private readonly _onSubmit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();

    const validationMessage: string = validateForm(this.state, this._today, this._spService.hasConfiguredListId());
    if (validationMessage) {
      this.setState({ validationMessage, status: 'none', statusMessage: '' });
      return;
    }

    this.setState({ isSubmitting: true, ...clearFeedbackState() });

    try {
      const incidentDate: string = getDateOnlyIsoString(this.state.incidentDate);
      const itemPayload: IIncidentListItemPayload = {
        'Title': `${this.state.fullName.trim()}-${incidentDate}`,
        'FullName': this.state.fullName.trim(),
        'Occupation': this.state.occupationRole.trim(),
        'IncidentAddress': this.state.addressOfIncident.trim(),
        'IncidentLocationName': this.state.locationNameOfIncident.trim(),
        'IncidentDate': incidentDate,
        'IncidentTime': this.state.incidentTime,
        'ReportedDate': getDateOnlyIsoString(this.state.dateReported),
        'ReportedTime': this.state.timeReported,
        'Supervisor': this.state.supervisor.trim() || undefined,
        'Witnesses': this.state.witness.trim() || undefined,
        'IncidentType': this.state.incidentType.trim(),
        'WhoWasInjured': this.state.whoWasInjured.trim(),
        'EquipmentDamaged': this.state.equipmentDamaged.trim() || undefined,
        'EquipmentIdentifyingNumber': this.state.equipmentIdentifyingNumber.trim() || undefined,
        'DamagedEquipmentDescription': this.state.damagedEquipmentDescription.trim() || undefined,
        'IncidentDescription': this.state.incidentDescription.trim(),
        'ActionsTaken': this.state.immediateActions.trim(),
        'NotificationRecipients': (this.props.notificationRecipients || '').trim() || undefined
      };

      const createdItemId: number = await this._spService.createIncidentItem(itemPayload);
      await this._spService.uploadAttachments(createdItemId, this.state.attachments);

      this.setState({
        ...createInitialState(this.props),
        status: 'success',
        statusMessage: this.props.successMessage || FORM_STATIC_TEXT.successFallbackMessage
      });
    } catch (error) {
      this.setState({
        isSubmitting: false,
        status: 'error',
        statusMessage: error instanceof Error && error.message
          ? error.message
          : (this.props.errorMessage || FORM_STATIC_TEXT.errorFallbackMessage),
        validationMessage: ''
      });
    }
  };

  private readonly _dismissDialog = (): void => {
    this.setState(clearFeedbackState());
  }
}