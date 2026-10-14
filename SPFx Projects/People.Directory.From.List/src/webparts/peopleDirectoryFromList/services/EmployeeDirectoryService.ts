import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import EmployeeInfo, { IEmployeeProfile } from '../models/EmployeeInfo';

export type { IEmployeeProfile } from '../models/EmployeeInfo';

interface IEmployeeDirectoryResponse {
  value: Record<string, unknown>[];
}

export default class EmployeeDirectoryService {
  private readonly _spHttpClient: SPHttpClient;
  private readonly _webAbsoluteUrl: string;

  public constructor(spHttpClient: SPHttpClient, webAbsoluteUrl: string) {
    this._spHttpClient = spHttpClient;
    this._webAbsoluteUrl = webAbsoluteUrl;
  }

  public async getEmployees(listId: string): Promise<IEmployeeProfile[]> {
    try {
      const normalizedListId: string = listId.replace(/[{}]/g, '').trim();

      if (!normalizedListId) {
        throw new Error('Employee list ID is required.');
      }

      const selectQuery: string = EmployeeInfo.selectFields.join(',');
      const endpoint: string = `${this._webAbsoluteUrl}/_api/web/lists(guid'${normalizedListId}')/items?$select=${selectQuery}&$top=5000`;

      const response: SPHttpClientResponse = await this._spHttpClient.get(
        endpoint,
        SPHttpClient.configurations.v1
      );

      if (!response.ok) {
        throw new Error(`Failed to fetch employee items: ${response.status} ${response.statusText}`);
      }

      const data: IEmployeeDirectoryResponse = await response.json();

      return data.value.map((item: Record<string, unknown>): IEmployeeProfile => ({
        id: Number(item[EmployeeInfo.columns.id]),
        fullName: this._getStringValue(item, EmployeeInfo.columns.fullName) || 'Unknown User',
        role: this._getStringValue(item, EmployeeInfo.columns.role) || 'N/A',
        birthday: this._getStringValue(item, EmployeeInfo.columns.birthday),
        employmentStartDate: this._getStringValue(item, EmployeeInfo.columns.employmentStartDate),
        phone: this._getStringValue(item, EmployeeInfo.columns.phone),
        email: this._getStringValue(item, EmployeeInfo.columns.email),
        state: this._getStringValue(item, EmployeeInfo.columns.state)
      }));
    } catch (error) {
      console.error('Error fetching employees:', error);
      throw error;
    }
  }

  private _getStringValue(item: Record<string, unknown>, key: string): string | undefined {
    const value: unknown = item[key];

    if (typeof value === 'string' && value.trim().length > 0) {
      return value.trim();
    }

    return undefined;
  }
}
