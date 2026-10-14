import { MSGraphClientV3 } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';

interface IUserLookupResult {
  id?: string;
  mail?: string;
  userPrincipalName?: string;
}

interface IUserPresenceResult {
  availability?: string;
  activity?: string;
}

export interface IEmployeeProfileVisuals {
  profileImageUrl?: string;
  availability?: string;
  activity?: string;
}

export default class EmployeeProfileService {
  private readonly _context: WebPartContext;

  public constructor(context: WebPartContext) {
    this._context = context;
  }

  public async getProfileVisuals(emailOrUpn?: string): Promise<IEmployeeProfileVisuals> {
    const identifier: string | undefined = emailOrUpn?.trim();

    if (!identifier) {
      return {};
    }

    try {
      const client: MSGraphClientV3 = await this._context.msGraphClientFactory.getClient('3');
      const user: IUserLookupResult = await client
        .api(`/users/${encodeURIComponent(identifier)}`)
        .select('id,mail,userPrincipalName')
        .get();

      const photoKey: string = (user.userPrincipalName || user.mail || identifier).trim();

      let presence: IUserPresenceResult = {};

      if (user.id) {
        try {
          const presenceResponse: IUserPresenceResult = await client
            .api(`/users/${encodeURIComponent(user.id)}/presence`)
            .get();

          presence = {
            availability: presenceResponse.availability,
            activity: presenceResponse.activity
          };
        } catch (presenceError) {
          console.warn(`Unable to load presence for user ${identifier}:`, presenceError);
        }
      }

      return {
        profileImageUrl: this._getUserPhotoUrl(photoKey),
        availability: presence.availability,
        activity: presence.activity
      };
    } catch (error) {
      console.warn(`Unable to resolve user profile for ${identifier}:`, error);
      return {
        profileImageUrl: this._getUserPhotoUrl(identifier)
      };
    }
  }

  private _getUserPhotoUrl(emailOrUpn: string): string {
    return `${this._context.pageContext.web.absoluteUrl}/_layouts/15/userphoto.aspx?size=L&accountname=${encodeURIComponent(emailOrUpn)}`;
  }
}
