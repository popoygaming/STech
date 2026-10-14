import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  ISharePointListItemResponse,
  IIncidentListItemPayload,
  IChoiceFieldResponse,
  ISharePointFieldResponse,
  ISharePointFieldCollectionResponse
} from './ISpServiceInterfaces';
export { ISharePointListItemResponse, IIncidentListItemPayload };

export default class SpService {
  private context: WebPartContext;
  private listId: string;
  private fieldInternalNameMapPromise: Promise<Map<string, string>> | undefined;
  private readonly fieldTitleAliases: Record<string, string[]> = {
    'type of incident': ['type of incident', 'incident type'],
    'location name of incident': ['location name of incident', 'location'],
    'description of incident / what occurred?': [
      'description of incident / what occurred?',
      'description of incident',
      'incident description'
    ],
    witness: ['witness', 'witnesses'],
    'immediate actions taken': ['immediate actions taken', 'actions taken']
  };

  public constructor(context: WebPartContext, listId: string) {
    this.context = context;
    this.listId = listId;
  }

  public hasConfiguredListId(): boolean {
    return !!this._getNormalizedListId();
  }

  public async getChoiceFieldOptionsByTitle(fieldTitle: string): Promise<string[]> {
    const safeFieldTitle: string = fieldTitle.replace(/'/g, "''");
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this._getNormalizedListId()}')/fields/getbytitle('${safeFieldTitle}')?$select=Choices`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch incident types: ${response.statusText}`);
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

  public async createIncidentItem(itemPayload: IIncidentListItemPayload): Promise<number> {
    const mappedPayload: Record<string, string> = await this.mapPayloadToInternalNames(itemPayload);
    const response = await this.context.spHttpClient.post(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this._getNormalizedListId()}')/items`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(mappedPayload)
      }
    );

    if (!response.ok) {
      const errorData = await response.json();
      throw new Error(errorData.error?.message || 'Failed to create incident item');
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
      const uploadUrl: string = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this._getNormalizedListId()}')/items(${itemId})/AttachmentFiles/add(FileName='${encodeURIComponent(file.name)}')`;

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
      throw new Error(`Failed to upload ${failedCount} attachment(s)`);
    }
  }

  private async mapPayloadToInternalNames(itemPayload: IIncidentListItemPayload): Promise<Record<string, string>> {
    const fieldInternalNameMap: Map<string, string> = await this.getFieldInternalNameMap();
    const mappedPayload: Record<string, string> = {};
    const missingFields: string[] = [];

    Object.keys(itemPayload).forEach((fieldTitle: string) => {
      const value: string | undefined = itemPayload[fieldTitle];

      if (value === undefined || value === null || value === '') {
        return;
      }

      const internalName: string | undefined = this.resolveInternalFieldName(fieldTitle, fieldInternalNameMap);
      if (!internalName) {
        missingFields.push(fieldTitle);
        return;
      }

      mappedPayload[internalName] = value;
    });

    if (missingFields.length > 0) {
      throw new Error(`The SharePoint list is missing these fields: ${missingFields.join(', ')}`);
    }

    return mappedPayload;
  }

  private async getFieldInternalNameMap(): Promise<Map<string, string>> {
    if (!this.fieldInternalNameMapPromise) {
      this.fieldInternalNameMapPromise = this.loadFieldInternalNameMap();
    }

    return this.fieldInternalNameMapPromise;
  }

  private async loadFieldInternalNameMap(): Promise<Map<string, string>> {
    const response = await this.context.spHttpClient.get(
      `${this.context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${this._getNormalizedListId()}')/fields?$select=Title,InternalName`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json',
          'Content-Type': 'application/json'
        }
      }
    );

    if (!response.ok) {
      throw new Error(`Failed to fetch list fields: ${response.statusText}`);
    }

    const data: ISharePointFieldCollectionResponse = await response.json();
    const fields: ISharePointFieldResponse[] = Array.isArray(data.value)
      ? data.value
      : (data.d?.results || []);
    const fieldMap: Map<string, string> = new Map<string, string>();

    fields.forEach((field: ISharePointFieldResponse) => {
      if (!field.Title || !field.InternalName) {
        return;
      }

      fieldMap.set(field.Title.trim().toLowerCase(), field.InternalName);
      fieldMap.set(this.normalizeFieldKey(field.Title), field.InternalName);
      fieldMap.set(field.InternalName.trim().toLowerCase(), field.InternalName);
      fieldMap.set(this.normalizeFieldKey(field.InternalName), field.InternalName);
    });

    return fieldMap;
  }

  private resolveInternalFieldName(fieldTitle: string, fieldInternalNameMap: Map<string, string>): string | undefined {
    const normalizedFieldTitle: string = fieldTitle.trim().toLowerCase();
    const aliases: string[] = this.fieldTitleAliases[normalizedFieldTitle] || [fieldTitle];

    for (let i: number = 0; i < aliases.length; i++) {
      const candidate: string = aliases[i];
      const exactMatch: string | undefined = fieldInternalNameMap.get(candidate.trim().toLowerCase());
      if (exactMatch) {
        return exactMatch;
      }

      const normalizedMatch: string | undefined = fieldInternalNameMap.get(this.normalizeFieldKey(candidate));
      if (normalizedMatch) {
        return normalizedMatch;
      }
    }

    return undefined;
  }

  private normalizeFieldKey(value: string): string {
    return value.toLowerCase().replace(/[^a-z0-9]/g, '');
  }

  private _getNormalizedListId(): string {
    return (this.listId || '').trim().replace('{', '').replace('}', '');
  }
}
