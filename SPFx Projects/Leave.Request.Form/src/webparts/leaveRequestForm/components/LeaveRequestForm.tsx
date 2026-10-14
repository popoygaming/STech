import * as React from 'react';
import {
  TextField,
  Dropdown,
  IDropdownOption,
  DatePicker,
  DayOfWeek,
  PrimaryButton,
  Dialog,
  DialogType,
  DialogFooter,
  Stack,
  Label,
  Spinner,
  SpinnerSize
} from '@fluentui/react';
import styles from './LeaveRequestForm.module.scss';
import type { ILeaveRequestFormProps } from './ILeaveRequestFormProps';
import { SPService, ILeaveRequestItem } from '../services/SPService';
import { GraphService } from '../services/GraphService';

export interface ILeaveRequestFormState {
  workTitle: string;
  department: string;
  approverOptions: IDropdownOption[];
  selectedApprover: string;
  dateFrom: Date | null;
  dateTo: Date | null;
  totalWorkingHours: number;
  hasUserEditedHours: boolean;
  typeOptions: IDropdownOption[];
  selectedType: string;
  attachments: FileList | null;
  additionalDetails: string;
  dialogInfo: {
    show: boolean;
    title: string;
    subText: string;
  };
  loading: boolean;
}

export default class LeaveRequestForm extends React.Component<ILeaveRequestFormProps, ILeaveRequestFormState> {
  private fileInputRef = React.createRef<HTMLInputElement>();
  private spService: SPService | null = null;
  private graphService: GraphService;
  private approverMap: Map<string, { displayName: string; email: string }> = new Map();

  constructor(props: ILeaveRequestFormProps) {
    super(props);
    this.graphService = new GraphService(props.context);
    this.state = {
      workTitle: '',
      department: '',
      approverOptions: [],
      selectedApprover: '',
      dateFrom: null,
      dateTo: null,
      totalWorkingHours: 0,
      hasUserEditedHours: false,
      typeOptions: [],
      selectedType: '',
      attachments: null,
      additionalDetails: '',
      dialogInfo: { show: false, title: '', subText: '' },
      loading: false
    };
  }

  public componentDidMount(): void {
    // initialize service if listId is provided
    if (this.props.listId) {
      this.spService = new SPService(this.props.context, this.props.listId);
    }

    // get user profile using GraphService
    this.graphService
      .getCurrentUserProfile()
      .then(profile => {
        this.setState({
          workTitle: profile.jobTitle,
          department: profile.department
        });
      })
      .catch(() => {
        // ignore
      });

    // load leave types using SPService
    if (this.spService) {
      this.spService
        .getLeaveTypeChoices()
        .then(choices => {
          const opts: IDropdownOption[] = choices.map(c => ({ key: c, text: c }));
          this.setState({ typeOptions: opts });
        })
        .catch(() => {
          // ignore
        });
    }

    // resolve approver display names using GraphService
    if (this.props.approvers && this.props.approvers.length) {
      this.graphService
        .resolveApprovers(this.props.approvers)
        .then(approvers => {
          const opts: IDropdownOption[] = approvers.map(a => ({ key: a.id, text: a.displayName }));
          // Store approver info in map for later use
          approvers.forEach(a => {
            this.approverMap.set(a.id, { displayName: a.displayName, email: a.email });
          });
          this.setState({ approverOptions: opts });
        })
        .catch(() => {
          // ignore
        });
    }
  }

  private calculateWorkingDays = (startDate: Date, endDate: Date): number => {
    const normalizedStart = new Date(startDate.getFullYear(), startDate.getMonth(), startDate.getDate());
    const normalizedEnd = new Date(endDate.getFullYear(), endDate.getMonth(), endDate.getDate());

    if (normalizedEnd < normalizedStart) {
      return 0;
    }

    if (normalizedStart.getTime() === normalizedEnd.getTime()) {
      const dayOfWeek = normalizedStart.getDay();
      return dayOfWeek !== 0 && dayOfWeek !== 6 ? 1 : 0;
    }

    let count = 0;
    const current = new Date(normalizedStart.getTime());

    while (current < normalizedEnd) {
      // Count each working day except the return-to-work day on the end date.
      const dayOfWeek = current.getDay();
      if (dayOfWeek !== 0 && dayOfWeek !== 6) {
        count++;
      }
      current.setDate(current.getDate() + 1);
    }

    return count;
  };

  private onDateChange = (field: 'dateFrom' | 'dateTo', date?: Date): void => {
    const newState: Partial<ILeaveRequestFormState> = { [field]: date || null };

    const selectedDate = date || null;
    if (field === 'dateFrom') {
      if (selectedDate && this.state.dateTo && selectedDate > this.state.dateTo) {
        newState.dateTo = null;
      }
    } else if (field === 'dateTo') {
      if (selectedDate && this.state.dateFrom && selectedDate < this.state.dateFrom) {
        newState.dateTo = null;
      }
    }

    const effectiveDateFrom = field === 'dateFrom' ? selectedDate : this.state.dateFrom;
    const effectiveDateTo = field === 'dateTo' ? selectedDate : this.state.dateTo;
    const finalDateFrom = effectiveDateFrom || (field === 'dateTo' ? this.state.dateFrom : null);
    const finalDateTo = effectiveDateTo || (field === 'dateFrom' ? this.state.dateTo : null);

    if (finalDateFrom && finalDateTo && !this.state.hasUserEditedHours) {
      const workingDays = this.calculateWorkingDays(finalDateFrom, finalDateTo);
      newState.totalWorkingHours = workingDays * 8;
    }

    this.setState(newState as ILeaveRequestFormState);
  };

  private onSubmit = (): void => {
    const missing: string[] = [];
    if (!this.state.selectedApprover) missing.push('Approver');
    if (!this.state.dateFrom) missing.push('Date From');
    if (!this.state.dateTo) missing.push('Date To');
    if (!this.state.selectedType) missing.push('Type of Leave');
    if (!this.state.totalWorkingHours || this.state.totalWorkingHours <= 0) missing.push('Total Working Hours');

    // Validate total working hours does not exceed the available hours between selected dates
    if (this.state.dateFrom && this.state.dateTo) {
      const maxHours = this.calculateWorkingDays(this.state.dateFrom, this.state.dateTo) * 8;
      if (this.state.totalWorkingHours > maxHours) {
        this.setState({
          dialogInfo: {
            show: true,
            title: 'Validation error',
            subText: `Total Working Hours cannot exceed the available hours between Date From and Date To (${maxHours} hours).`
          }
        });
        return;
      }
    }

    if (this.state.dateFrom && this.state.dateTo && this.state.dateFrom > this.state.dateTo) {
      this.setState({
        dialogInfo: {
          show: true,
          title: 'Validation error',
          subText: 'Date From must be earlier than Date To.'
        }
      });
      return;
    }

    if (missing.length) {
      this.setState({
        dialogInfo: {
          show: true,
          title: 'Validation error',
          subText: `The following fields are required: ${missing.join(', ')}`
        }
      });
      return;
    }

    if (!this.spService) {
      this.setState({
        dialogInfo: {
          show: true,
          title: 'Error',
          subText: 'Service not initialized. Please refresh the page.'
        }
      });
      return;
    }

    this.setState({ loading: true });
    
    // Get approver name and email from the map
    const approverInfo = this.approverMap.get(this.state.selectedApprover);
    if (!approverInfo) {
      this.setState({
        loading: false,
        dialogInfo: {
          show: true,
          title: 'Error',
          subText: 'Could not resolve approver information.'
        }
      });
      return;
    }

    const item: ILeaveRequestItem = {
      Title: `${this.props.userDisplayName} - Leave Request`,
      FullName: this.props.userDisplayName,
      WorkTitle: this.state.workTitle,
      ApproverName: approverInfo.displayName,
      ApproverEmail: approverInfo.email,
      DateFrom: this.state.dateFrom!.toISOString(),
      DateTo: this.state.dateTo!.toISOString(),
      TotalWorkingDays: this.state.totalWorkingHours,
      TypeOfLeave: this.state.selectedType,
      AdditionalDetails: this.state.additionalDetails
    };

    this.spService
      .createLeaveRequest(item)
      .then(created => {
        const itemId = created.Id;
        if (this.state.attachments && this.state.attachments.length && this.spService) {
          return this.spService.uploadAttachments(itemId, this.state.attachments);
        }
        return Promise.resolve();
      })
      .then(() => {
        this.setState({
          loading: false,
          dialogInfo: { show: true, title: 'Success', subText: this.props.successMessage || 'Your leave request has been submitted.' },
          selectedApprover: '',
          dateFrom: null,
          dateTo: null,
          totalWorkingHours: 0,
          hasUserEditedHours: false,
          selectedType: '',
          attachments: null,
          additionalDetails: ''
        });
        if (this.fileInputRef.current) {
          this.fileInputRef.current.value = '';
        }
      })
      .catch(err => {
        this.setState({ loading: false, dialogInfo: { show: true, title: 'Error', subText: err.message || 'Failed to save the request.' } });
      });
  };

  public render(): React.ReactElement<ILeaveRequestFormProps> {
    if (!this.props.listId) {
      return (
        <div className={styles.leaveRequestForm}>
          <strong>Configuration required:</strong> please specify a target list GUID in the property pane.
        </div>
      );
    }

    return (
      <section className={styles.leaveRequestForm} style={{ position: 'relative', margin: 0, padding: 0 }}>
        {this.state.loading && (
          <div style={{
            position: 'absolute',
            top: 0,
            left: 0,
            width: '100%',
            height: '100%',
            background: 'rgba(255,255,255,0.7)',
            zIndex: 10,
            display: 'flex',
            flexDirection: 'column',
            alignItems: 'center',
            justifyContent: 'center'
          }}>
            <Spinner size={SpinnerSize.large} label="Your leave request is being submitted, please wait..." labelPosition="bottom" />
          </div>
        )}
        {this.props.showTitle && <h2>Leave Request Form</h2>}
        <Stack tokens={{ childrenGap: 15 }}>
          <TextField label="Full Name" value={this.props.userDisplayName} readOnly />
          <TextField label="Work Title" value={this.state.workTitle} onChange={(_, val) => this.setState({ workTitle: val || '' })} />
          <Dropdown
            label="Approver"
            placeholder="Select an approver"
            options={this.state.approverOptions}
            selectedKey={this.state.selectedApprover}
            onChange={(_, option) => this.setState({ selectedApprover: option?.key as string })}
            required
          />
          <DatePicker
            label="Date From"
            firstDayOfWeek={DayOfWeek.Monday}
            value={this.state.dateFrom || undefined}
            onSelectDate={date => this.onDateChange('dateFrom', date || undefined)}
            // maxDate={this.state.dateTo || undefined}
            isRequired
          />
          <DatePicker
            label="Date To"
            firstDayOfWeek={DayOfWeek.Monday}
            value={this.state.dateTo || undefined}
            onSelectDate={date => this.onDateChange('dateTo', date || undefined)}
            minDate={this.state.dateFrom || undefined}
            isRequired
          />
          <TextField
            label="Total Working Hours"
            type="number"
            step="0.5"
            min="0"
            value={this.state.totalWorkingHours.toString()}
            onChange={(_, val) => this.setState({ totalWorkingHours: parseFloat(val || '0') || 0, hasUserEditedHours: true })}
            required
          />
          <div style={{ color: '#111111', fontSize: '13px', marginTop: '4px', marginBottom: '2px' }}>
              Enter the total number of scheduled hours you will be absent. Do not include weekends or public holidays.
            </div>
          <div style={{ color: '#111111', fontSize: '13px', marginTop: '4px', marginBottom: '8px' }}>
              [A typical working day = 8 hours, a week = 40 hours].
            </div>

          <Dropdown
            label="Type of Leave"
            placeholder="Choose a leave type"
            options={this.state.typeOptions}
            selectedKey={this.state.selectedType}
            onChange={(_, option) => this.setState({ selectedType: option?.key as string })}
            required
          />
          {this.props.sickLeaveNote && this.props.sickLeaveNote.trim() !== '' && (
            <div style={{ color: '#111111', fontSize: '13px', marginTop: '4px', marginBottom: '8px' }}>
              {this.props.sickLeaveNote}
            </div>
          )}

          <Label>Attachments</Label>
          <input
            ref={this.fileInputRef}
            type="file"
            multiple
            onChange={e => this.setState({ attachments: e.target.files })}
          />
          <TextField
            label="Additional Details"
            multiline
            rows={3}
            placeholder="Please provide any further details that may assist your leave request."
            value={this.state.additionalDetails}
            onChange={(_, val) => this.setState({ additionalDetails: val || '' })}
          />
          <PrimaryButton text="Submit" onClick={this.onSubmit} disabled={this.state.loading} />
        </Stack>

        <Dialog
          hidden={!this.state.dialogInfo.show}
          onDismiss={() => this.setState({ dialogInfo: { ...this.state.dialogInfo, show: false } })}
          dialogContentProps={{ type: DialogType.normal, title: this.state.dialogInfo.title, subText: this.state.dialogInfo.subText }}
        >
          <DialogFooter>
            <PrimaryButton onClick={() => this.setState({ dialogInfo: { ...this.state.dialogInfo, show: false } })} text="OK" />
          </DialogFooter>
        </Dialog>
      </section>
    );
  }
}
