import * as React from 'react';
import styles from './FeedbackForm.module.scss';
import type { IFeedbackFormProps } from './IFeedbackFormProps';
import { IFeedbackFormState } from './IFeedbackFormState';
import { IDropdownOption, Dropdown, TextField, PrimaryButton, Stack, Dialog, DialogType, DialogFooter, DefaultButton } from '@fluentui/react';
import { SPHttpClient } from '@microsoft/sp-http';

export default class FeedbackForm extends React.Component<IFeedbackFormProps, IFeedbackFormState> {
    constructor(props: IFeedbackFormProps) {
    super(props);
    this.state = {
      feedback: "",
      feedbackType: "",
      submittedBy: "",
      feedbackTypeOptions: [],
      submitted: false,
      feedbackError: "",
      feedbackTypeError: "",
      showDialog: false,
      dialogMessage: "",
      dialogIsError: false
    };
  }
public componentDidMount(): void {
    // Get current user
    this.props.context.spHttpClient.get(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/currentuser`,
      SPHttpClient.configurations.v1
    ).then(response => response.json())
      .then(user => {
        this.setState({ submittedBy: user.Title });
      });

    // Get Feedback Type choices from list column
    this.props.context.spHttpClient.get(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.props.feedbackListId}')/fields/getbytitle('Feedback Type')`,
      SPHttpClient.configurations.v1
    ).then(response => response.json())
      .then(field => {
        const options: IDropdownOption[] = field.Choices.map((choice: string) => ({
          key: choice,
          text: choice
        }));
        this.setState({ feedbackTypeOptions: options });
      });

  }

  private handleSubmit = (): void => {
    // validation
    let feedbackErr = '';
    let typeErr = '';
    if (!this.state.feedback.trim()) {
      feedbackErr = 'Feedback is required';
    }
    if (!this.state.feedbackType) {
      typeErr = 'Feedback type is required';
    }
    if (feedbackErr || typeErr) {
      this.setState({
        showDialog: true,
        dialogMessage: 'Please complete the required fields.',
        dialogIsError: true,
        feedbackError: feedbackErr,
        feedbackTypeError: typeErr
      });
      return;
    }

    const item = {
      Title: this.state.feedbackType,
      Feedback: this.state.feedback,
      FeedbackType: this.state.feedbackType
    };

    // Save feedback
    this.props.context.spHttpClient.post(
      `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.props.feedbackListId}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          'Accept': 'application/json;odata=nometadata',
          'Content-type': 'application/json;odata=nometadata'
        },
        body: JSON.stringify(item)
      }
    ).then(() => {
      // show dialog thank you
      this.setState({
        showDialog: true,
        dialogMessage: this.props.successMessage,
        dialogIsError: false,
        feedback: '',
        feedbackType: '',
        feedbackError: '',
        feedbackTypeError: ''
      });
    }).catch(() => {
      this.setState({
        showDialog: true,
        dialogMessage: this.props.errorMessage,
        dialogIsError: true
      });
    });
  }

  public render(): React.ReactElement<IFeedbackFormProps> {
    return (
      <div className={styles.container}>
        <Stack tokens={{ childrenGap: 20}} horizontalAlign="stretch">
          {this.props.showTitle && <h3 className={styles.title}>{this.props.title}</h3>}
          {this.props.showSubTitle && (
            <p className={styles.subTitle}>
              {this.props.subTitle}
            </p>
          )}

          <Dialog
            hidden={!this.state.showDialog}
            dialogContentProps={{
              type: DialogType.close,
              title: this.state.dialogIsError ? 'Error' : 'Information',
              subText: this.state.dialogMessage
            }}
            modalProps={{
              isBlocking: true
            }}
            onDismiss={() => this.setState({ showDialog: false })}
          >
            <DialogFooter>
              <DefaultButton
                onClick={() => this.setState({ showDialog: false })}
                text="OK"
              />
              {!this.state.dialogIsError && this.props.homeUrl && (
                <PrimaryButton
                  onClick={() => {
                    this.setState({ showDialog: false });
                    window.location.href = this.props.homeUrl;
                  }}
                  text="Back to Intranet Home"
                />
              )}
            </DialogFooter>
          </Dialog>
          <TextField label='Submitted by' value={this.state.submittedBy} readOnly styles={{ root: { width: '100%', marginTop: '-12px' } }} />
          <Dropdown
            label="Feedback type"
            options={this.state.feedbackTypeOptions}
            selectedKey={this.state.feedbackType}
            onChange={(e, option) => this.setState({ feedbackType: option?.key as string, feedbackTypeError: '' })}
            placeholder="Select a feedback type"
            styles={{ root: { width: '100%' } }}
          />
          {this.state.feedbackTypeError && (
            <div className={styles.errorText}>{this.state.feedbackTypeError}</div>
          )}

          <TextField
            label="Enter your feedback"
            multiline
            rows={4}
            value={this.state.feedback}
            onChange={(e, newValue) => this.setState({ feedback: newValue || "", feedbackError: '' })}
            styles={{ root: { width: '100%' } }}
            errorMessage={this.state.feedbackError}
          />

          <PrimaryButton text="Submit" onClick={this.handleSubmit} styles={{ root: { width: '100%' } }} />
        </Stack>
      </div>
    );
  }
}
