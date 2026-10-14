import * as React from 'react';
import {
  ComboBox,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  IComboBox,
  IComboBoxOption,
  Label,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Stack,
  TextField
} from '@fluentui/react';
import styles from './EquipmentReportForm.module.scss';
import type { IEquipmentReportFormProps } from './IEquipmentReportFormProps';
import EquipmentReportService, { IEquipmentReportPayload } from '../services/EquipmentReportService';
import {
  FALLBACK_ASSET_TYPE_CHOICES,
  FALLBACK_URGENCY_CHOICES,
  FIELD_TITLES
} from '../constants/EquipmentReportConstants';

type SubmissionStatus = 'none' | 'success' | 'error';

interface IEquipmentReportFormState {
  fullName: string;
  assetType: string;
  assetTypeOptions: IComboBoxOption[];
  assetIdRegistration: string;
  usageHours: string;
  issueDescription: string;
  urgencyLevel: string;
  urgencyLevelOptions: IComboBoxOption[];
  assetLocation: string;
  attachments: File[];
  isSubmitting: boolean;
  status: SubmissionStatus;
  statusMessage: string;
  validationMessage: string;
}

export default class EquipmentReportForm extends React.Component<IEquipmentReportFormProps> {
  private _equipmentService: EquipmentReportService;

  public constructor(props: IEquipmentReportFormProps) {
    super(props);

    this._equipmentService = new EquipmentReportService(props.context, props.listId);

    this.state = {
      fullName: props.userDisplayName || '',
      assetType: '',
      assetTypeOptions: this._mapChoicesToOptions(FALLBACK_ASSET_TYPE_CHOICES),
      assetIdRegistration: '',
      usageHours: '',
      issueDescription: '',
      urgencyLevel: '',
      urgencyLevelOptions: this._mapChoicesToOptions(FALLBACK_URGENCY_CHOICES),
      assetLocation: '',
      attachments: [],
      isSubmitting: false,
      status: 'none',
      statusMessage: '',
      validationMessage: ''
    };
  }

  public componentDidMount(): void {
    this._loadChoiceFields().catch(() => {
      this.setState({
        assetTypeOptions: this._mapChoicesToOptions(FALLBACK_ASSET_TYPE_CHOICES),
        urgencyLevelOptions: this._mapChoicesToOptions(FALLBACK_URGENCY_CHOICES)
      });
    });
  }

  public componentDidUpdate(prevProps: IEquipmentReportFormProps): void {
    if (prevProps.listId !== this.props.listId) {
      this._equipmentService = new EquipmentReportService(this.props.context, this.props.listId);

      this.setState({
        assetType: '',
        urgencyLevel: '',
        assetTypeOptions: this._mapChoicesToOptions(FALLBACK_ASSET_TYPE_CHOICES),
        urgencyLevelOptions: this._mapChoicesToOptions(FALLBACK_URGENCY_CHOICES),
        status: 'none',
        statusMessage: '',
        validationMessage: ''
      });

      this._loadChoiceFields().catch(() => {
        this.setState({
          assetTypeOptions: this._mapChoicesToOptions(FALLBACK_ASSET_TYPE_CHOICES),
          urgencyLevelOptions: this._mapChoicesToOptions(FALLBACK_URGENCY_CHOICES)
        });
      });
    }
  }

  public render(): React.ReactElement<IEquipmentReportFormProps> {
    const state: IEquipmentReportFormState = this.state as IEquipmentReportFormState;

    return (
      <div className={styles.equipmentReportForm} aria-busy={state.isSubmitting}>
        {this.props.showTitle && (
          <h2 className={styles.title}>{this.props.title || 'Equipment Report Form'}</h2>
        )}

        <Dialog
          hidden={!this._isDialogVisible()}
          onDismiss={this._dismissDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: this._getDialogTitle(),
            subText: this._getDialogMessage()
          }}
          modalProps={{ isBlocking: false }}
        >
          <DialogFooter>
            <DefaultButton text="OK" onClick={this._dismissDialog} />
          </DialogFooter>
        </Dialog>

        <form onSubmit={this._onSubmit} noValidate>
          <Stack tokens={{ childrenGap: 16 }}>
            <Stack horizontal wrap tokens={{ childrenGap: 16 }} className={styles.grid}>
              <Stack.Item grow className={styles.gridItem}>
                <TextField
                  label="Full Name"
                  value={state.fullName}
                  ariaLabel="Full Name"
                  readOnly={true}
                />
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <ComboBox
                  label="Asset Type"
                  options={state.assetTypeOptions}
                  selectedKey={state.assetType || undefined}
                  text={state.assetType}
                  onChange={this._onAssetTypeChange}
                  allowFreeform={true}
                  autoComplete="on"
                  ariaLabel="Asset Type"
                />
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <div>
                  <Label htmlFor="assetIdRegistrationInput">Asset ID/Registration</Label>
                  <TextField
                    id="assetIdRegistrationInput"
                    value={state.assetIdRegistration}
                    onChange={this._onAssetIdChange}
                    ariaLabel="Asset ID or Registration"
                    description='(e.g., Number plate or Serial Number)'
                  />
                </div>
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <TextField
                  label="Current Odometer/Usage Hours"
                  value={state.usageHours}
                  onChange={this._onUsageHoursChange}
                  ariaLabel="Current Odometer or Usage Hours"
                  description='(Crucial for vehicle servicing)'
                />
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <ComboBox
                  label="Urgency Level"
                  options={state.urgencyLevelOptions}
                  selectedKey={state.urgencyLevel || undefined}
                  text={state.urgencyLevel}
                  onChange={this._onUrgencyLevelChange}
                  allowFreeform={true}
                  autoComplete="on"
                  ariaLabel="Urgency Level"
                />
              </Stack.Item>

              <Stack.Item grow className={styles.gridItem}>
                <TextField
                  label="Location of Asset"
                  value={state.assetLocation}
                  onChange={this._onLocationChange}
                  ariaLabel="Location of Asset"
                  description='(Which branch or site is it currently at?)'
                />
              </Stack.Item>
            </Stack>

            <TextField
              label="Description of Issue/Service Required"
              multiline={true}
              rows={4}
              value={state.issueDescription}
              onChange={this._onIssueDescriptionChange}
              ariaLabel="Description of Issue or Service Required"
              description='(Comments section)'
            />

            <div>
              <Label htmlFor="equipmentAttachmentInput">Photo/File Upload</Label>
              <input
                id="equipmentAttachmentInput"
                type="file"
                multiple={true}
                onChange={this._onAttachmentsChange}
                aria-label="Photo or file upload"
              />
              <div style={{ marginTop: 4, fontStyle: 'italic', fontSize: 12 }}>
                (For showing damage or error screens)
              </div>
              {state.attachments.length > 0 && (
                <div className={styles.attachmentText}>{state.attachments.length} file(s) selected</div>
              )}
            </div>

            <PrimaryButton text="Submit Request" type="submit" disabled={state.isSubmitting} />
          </Stack>
        </form>

        {state.isSubmitting && (
          <div className={styles.submittingOverlay} role="status" aria-live="polite" aria-label="Submitting request">
            <div className={styles.submittingSpinner}>
              <Spinner label="Submitting request..." size={SpinnerSize.large} />
            </div>
          </div>
        )}
      </div>
    );
  }

  private readonly _onAssetTypeChange = (
    _event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption,
    _index?: number,
    value?: string
  ): void => {
    this.setState({ assetType: option ? String(option.key) : (value || ''), validationMessage: '', status: 'none', statusMessage: '' });
  };

  private readonly _onAssetIdChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ assetIdRegistration: value || '', validationMessage: '', status: 'none', statusMessage: '' });
  };

  private readonly _onUsageHoursChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ usageHours: value || '', validationMessage: '', status: 'none', statusMessage: '' });
  };

  private readonly _onIssueDescriptionChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ issueDescription: value || '', validationMessage: '', status: 'none', statusMessage: '' });
  };

  private readonly _onUrgencyLevelChange = (
    _event: React.FormEvent<IComboBox>,
    option?: IComboBoxOption,
    _index?: number,
    value?: string
  ): void => {
    this.setState({ urgencyLevel: option ? String(option.key) : (value || ''), validationMessage: '', status: 'none', statusMessage: '' });
  };

  private readonly _onLocationChange = (_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, value?: string): void => {
    this.setState({ assetLocation: value || '', validationMessage: '', status: 'none', statusMessage: '' });
  };

  private readonly _onAttachmentsChange = (event: React.ChangeEvent<HTMLInputElement>): void => {
    const files: File[] = [];
    const inputFiles: FileList | null = event.target.files;

    if (inputFiles) {
      for (let i: number = 0; i < inputFiles.length; i++) {
        files.push(inputFiles.item(i) as File);
      }
    }

    this.setState({ attachments: files, validationMessage: '', status: 'none', statusMessage: '' });
  };

  private readonly _onSubmit = async (event: React.FormEvent<HTMLFormElement>): Promise<void> => {
    event.preventDefault();

    const validationMessage: string = this._validateForm();
    if (validationMessage) {
      this.setState({ validationMessage, status: 'none', statusMessage: '' });
      return;
    }

    this.setState({ isSubmitting: true, validationMessage: '', status: 'none', statusMessage: '' });

    try {
      const state: IEquipmentReportFormState = this.state as IEquipmentReportFormState;
      const payload: IEquipmentReportPayload = {
        title: `Maintenance or Fleet Request - ${state.fullName}-${new Date().toISOString()}`,
        fullName: state.fullName.trim(),
        assetType: state.assetType.trim(),
        assetIdRegistration: state.assetIdRegistration.trim(),
        usageHours: state.usageHours.trim(),
        issueDescription: state.issueDescription.trim(),
        urgencyLevel: state.urgencyLevel.trim(),
        assetLocation: state.assetLocation.trim(),
        notificationReceiver: (this.props.notificationReceiver || '').trim()
      };

      const createdItemId: number = await this._equipmentService.createEquipmentReportItem(payload, FIELD_TITLES);
      await this._equipmentService.uploadAttachments(createdItemId, state.attachments);

      this.setState({
        assetType: '',
        assetIdRegistration: '',
        usageHours: '',
        issueDescription: '',
        urgencyLevel: '',
        assetLocation: '',
        attachments: [],
        isSubmitting: false,
        status: 'success',
        statusMessage: this.props.successMessage || 'Maintenance or fleet request submitted successfully.',
        validationMessage: ''
      });
    } catch {
      this.setState({
        isSubmitting: false,
        status: 'error',
        statusMessage: this.props.errorMessage || 'An error occurred while submitting the maintenance or fleet request.',
        validationMessage: ''
      });
    }
  };

  private _validateForm(): string {
    if (!this._equipmentService.hasConfiguredListId()) {
      return 'Please configure the List ID (GUID) in the web part properties.';
    }

    return '';
  }

  private _isDialogVisible(): boolean {
    const state: IEquipmentReportFormState = this.state as IEquipmentReportFormState;
    return !!state.validationMessage || state.status !== 'none';
  }

  private _getDialogTitle(): string {
    const state: IEquipmentReportFormState = this.state as IEquipmentReportFormState;

    if (state.validationMessage) {
      return 'Validation Message';
    }

    if (state.status === 'success') {
      return 'Success';
    }

    if (state.status === 'error') {
      return 'Error';
    }

    return 'Message';
  }

  private _getDialogMessage(): string {
    const state: IEquipmentReportFormState = this.state as IEquipmentReportFormState;

    if (state.validationMessage) {
      return state.validationMessage;
    }

    return state.statusMessage;
  }

  private readonly _dismissDialog = (): void => {
    this.setState({ validationMessage: '', status: 'none', statusMessage: '' });
  };

  private async _loadChoiceFields(): Promise<void> {
    if (!this._equipmentService.hasConfiguredListId()) {
      this.setState({
        assetTypeOptions: this._mapChoicesToOptions(FALLBACK_ASSET_TYPE_CHOICES),
        urgencyLevelOptions: this._mapChoicesToOptions(FALLBACK_URGENCY_CHOICES)
      });
      return;
    }

    const [assetTypeChoices, urgencyLevelChoices]: [string[], string[]] = await Promise.all([
      this._equipmentService.getChoiceFieldOptionsByTitle(FIELD_TITLES.assetType),
      this._equipmentService.getChoiceFieldOptionsByTitle(FIELD_TITLES.urgencyLevel)
    ]);

    this.setState({
      assetTypeOptions: this._mapChoicesToOptions(assetTypeChoices.length ? assetTypeChoices : FALLBACK_ASSET_TYPE_CHOICES),
      urgencyLevelOptions: this._mapChoicesToOptions(urgencyLevelChoices.length ? urgencyLevelChoices : FALLBACK_URGENCY_CHOICES)
    });
  }

  private _mapChoicesToOptions(choices: string[]): IComboBoxOption[] {
    return choices
      .filter((choice: string) => !!choice && !!choice.trim())
      .map((choice: string) => ({
        key: choice,
        text: choice
      }));
  }
}
