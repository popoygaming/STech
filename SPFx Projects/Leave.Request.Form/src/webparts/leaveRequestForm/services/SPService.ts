import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface ILeaveRequestItem {
  Title: string;
  FullName: string;
  WorkTitle: string;
  ApproverName: string; 
  ApproverEmail: string;
  DateFrom: string;
  DateTo: string;
  TotalWorkingDays: number;
  TypeOfLeave: string;
  AdditionalDetails: string;
}

export class SPService {
  private context: WebPartContext;
  private listId: string;

  constructor(context: WebPartContext, listId: string) {
    this.context = context;
    this.listId = listId;
  }

  /**
   * Fetch leave type choices from the "Type of Leave" field in the list
   */
  public async getLeaveTypeChoices(): Promise<string[]> {
    try {
      const response = await this.context.spHttpClient.get(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.listId}')/fields/getbytitle('Leave Type')?$select=Choices`,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch leave types: ${response.statusText}`);
      }

      const data = await response.json();
      return data.Choices || [];
    } catch (error) {
      console.error('Error fetching leave type choices:', error);
      throw error;
    }
  }

  /**
   * Create a new leave request item in the list
   */
  public async createLeaveRequest(item: ILeaveRequestItem): Promise<{ Id: number }> {
    try {
      const response = await this.context.spHttpClient.post(
        `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.listId}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json',
            'Content-Type': 'application/json'
          },
          body: JSON.stringify(item)
        }
      );

      if (!response.ok) {
        const errorData = await response.json();
        throw new Error(errorData.error?.message || 'Failed to create leave request');
      }

      const createdItem = await response.json();
      return { Id: createdItem.Id };
    } catch (error) {
      console.error('Error creating leave request:', error);
      throw error;
    }
  }

  /**
   * Upload attachments to a list item
   */
  public async uploadAttachments(itemId: number, files: FileList): Promise<void> {
    try {
      console.log(`Uploading ${files.length} attachment(s) for item ID ${itemId}...`);
      let failedCount = 0;
      for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const uploadUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this.listId}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
        const response = await this.context.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
          headers: {
            'Accept': 'application/json'
          },
          body: file
        });
        if (!response.ok) {
          failedCount++;
        }
      }
      if (failedCount > 0) {
        throw new Error(`Failed to upload ${failedCount} attachment(s)`);
      }
    } catch (error) {
      console.error('Error uploading attachments:', error);
      throw error;
    }
  }
}
