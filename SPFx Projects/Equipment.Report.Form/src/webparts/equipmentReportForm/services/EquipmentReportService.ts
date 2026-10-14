import { SPHttpClient } from '@microsoft/sp-http';

interface ISPHttpContext {
  pageContext: {
    web: {
      absoluteUrl: string;
    };
  };
  spHttpClient: {
    get: typeof SPHttpClient.prototype.get;
    post: typeof SPHttpClient.prototype.post;
  };
}

export interface IEquipmentReportPayload {
  title: string;
  fullName: string;
  assetType: string;
  assetIdRegistration: string;
  usageHours: string;
  issueDescription: string;
  urgencyLevel: string;
  assetLocation: string;
  notificationReceiver?: string;
}

interface IChoiceFieldResponse {
  Choices?: string[] | { results: string[] };
}

interface IFieldInfo {
  Title: string;
  InternalName: string;
}

interface IFieldCollectionResponse {
  value: IFieldInfo[];
}

interface ISharePointListItemResponse {
  Id: number;
}

export default class EquipmentReportService {
  private context: ISPHttpContext;
  private listId: string;
  private fieldTitleToInternalName: Map<string, string>;

  public constructor(context: ISPHttpContext, listId: string) {
    this.context = context;
    this.listId = listId;
    this.fieldTitleToInternalName = new Map<string, string>();
  }

  public hasConfiguredListId(): boolean {
    return !!this._getNormalizedListId();
  }

  public async getChoiceFieldOptionsByTitle(fieldTitle: string): Promise<string[]> {
    const safeFieldTitle: string = fieldTitle.replace(/'/g, "''");
    const response = await this.context.spHttpClient.get(
      `${this._getListApiRoot()}/fields/getbytitle('${safeFieldTitle}')?$select=Choices`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch choices for ${fieldTitle}: ${response.statusText}`);
    }

    const data: IChoiceFieldResponse = await response.json();
    if (Array.isArray(data.Choices)) {
      return data.Choices;
    }

    if (data.Choices && Array.isArray(data.Choices.results)) {
      return data.Choices.results;
    }

    return [];
  }

  public async createEquipmentReportItem(payload: IEquipmentReportPayload, fieldTitles: {
    fullName: string;
    assetType: string;
    assetIdRegistration: string;
    usageHours: string;
    issueDescription: string;
    urgencyLevel: string;
    assetLocation: string;
    notificationReceiver: string;
  }): Promise<number> {
    const [
      fullNameInternal,
      assetTypeInternal,
      assetIdInternal,
      usageHoursInternal,
      issueDescriptionInternal,
      urgencyLevelInternal,
      assetLocationInternal,
      notificationReceiverInternal
    ]: string[] = await Promise.all([
      this._getInternalNameByFieldTitle(fieldTitles.fullName),
      this._getInternalNameByFieldTitle(fieldTitles.assetType),
      this._getInternalNameByFieldTitle(fieldTitles.assetIdRegistration),
      this._getInternalNameByFieldTitle(fieldTitles.usageHours),
      this._getInternalNameByFieldTitle(fieldTitles.issueDescription),
      this._getInternalNameByFieldTitle(fieldTitles.urgencyLevel),
      this._getInternalNameByFieldTitle(fieldTitles.assetLocation),
      this._getInternalNameByFieldTitle(fieldTitles.notificationReceiver)
    ]);

    const itemBody: Record<string, string> = {
      Title: payload.title,
      [fullNameInternal]: payload.fullName,
      [assetTypeInternal]: payload.assetType,
      [assetIdInternal]: payload.assetIdRegistration,
      [usageHoursInternal]: payload.usageHours,
      [issueDescriptionInternal]: payload.issueDescription,
      [urgencyLevelInternal]: payload.urgencyLevel,
      [assetLocationInternal]: payload.assetLocation
    };

    if (payload.notificationReceiver && notificationReceiverInternal) {
      itemBody[notificationReceiverInternal] = payload.notificationReceiver;
    }

    const response = await this.context.spHttpClient.post(
      `${this._getListApiRoot()}/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(itemBody)
      }
    );

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error?.message || 'Failed to create equipment report item.');
    }

    const createdItem: ISharePointListItemResponse = await response.json();
    return createdItem.Id;
  }

  public async uploadAttachments(itemId: number, attachments: File[]): Promise<void> {
    if (!attachments.length) {
      return;
    }

    let failedCount: number = 0;

    for (let i: number = 0; i < attachments.length; i++) {
      const file: File = attachments[i];
      const uploadUrl: string = `${this._getListApiRoot()}/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;
      const response = await this.context.spHttpClient.post(uploadUrl, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json'
        },
        body: file
      });

      if (!response.ok) {
        failedCount++;
      }
    }

    if (failedCount > 0) {
      throw new Error(`Failed to upload ${failedCount} attachment(s).`);
    }
  }

  private async _getInternalNameByFieldTitle(fieldTitle: string): Promise<string> {
    const normalizedTitle: string = (fieldTitle || '').trim();
    if (!normalizedTitle) {
      return '';
    }

    const cachedInternalName: string | undefined = this.fieldTitleToInternalName.get(normalizedTitle.toLowerCase());
    if (cachedInternalName) {
      return cachedInternalName;
    }

    await this._ensureFieldMap();

    return this.fieldTitleToInternalName.get(normalizedTitle.toLowerCase()) || '';
  }

  private async _ensureFieldMap(): Promise<void> {
    if (this.fieldTitleToInternalName.size > 0) {
      return;
    }

    const response = await this.context.spHttpClient.get(
      `${this._getListApiRoot()}/fields?$select=Title,InternalName`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch list fields: ${response.statusText}`);
    }

    const data: IFieldCollectionResponse = await response.json();
    const fields: IFieldInfo[] = data.value || [];

    fields.forEach((field: IFieldInfo) => {
      this.fieldTitleToInternalName.set((field.Title || '').toLowerCase(), field.InternalName || '');
    });
  }

  private _getListApiRoot(): string {
    return `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this._getNormalizedListId()}')`;
  }

  private _getNormalizedListId(): string {
    return (this.listId || '').trim().replace('{', '').replace('}', '');
  }
}
