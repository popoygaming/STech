import { WebPartContext } from '@microsoft/sp-webpart-base';

export interface IUserProfile {
  jobTitle: string;
  department: string;
}

export interface IApproverInfo {
  id: string;
  displayName: string;
  email: string;
}

export class GraphService {
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Get current user's profile information (jobTitle, department)
   */
  public async getCurrentUserProfile(): Promise<IUserProfile> {
    try {
      const client = await this.context.msGraphClientFactory.getClient('3');
      const me = await client.api('/me').select('jobTitle,department').get();

      return {
        jobTitle: me.jobTitle || '',
        department: me.department || ''
      };
    } catch (error) {
      console.error('Error fetching current user profile:', error);
      throw error;
    }
  }

  /**
   * Resolve multiple approver identifiers to display names
   * Accepts user IDs, emails, or UPNs
   */
  public async resolveApprovers(approverIds: string[]): Promise<IApproverInfo[]> {
    try {
      const results = await Promise.all(
        approverIds.map(id => this.getUserInfo(id))
      );
      return results;
    } catch (error) {
      console.error('Error resolving approvers:', error);
      throw error;
    }
  }

  /**
   * Get user information by ID, email, or UPN with email details
   */
  private async getUserInfo(userIdentifier: string): Promise<IApproverInfo> {
    try {
      const client = await this.context.msGraphClientFactory.getClient('3');
      const encoded = encodeURIComponent(userIdentifier);

      const user = await client
        .api(`/users/${encoded}`)
        .select('displayName,mail,userPrincipalName')
        .get();

      return {
        id: userIdentifier,
        displayName: user.displayName || user.mail || user.userPrincipalName || userIdentifier,
        email: user.mail || user.userPrincipalName || userIdentifier
      };
    } catch (error) {
      console.warn(`Could not resolve user ${userIdentifier}:`, error);
      // Return the identifier as fallback
      return {
        id: userIdentifier,
        displayName: userIdentifier,
        email: userIdentifier
      };
    }
  }
}
