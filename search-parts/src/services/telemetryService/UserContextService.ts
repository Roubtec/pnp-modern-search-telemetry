import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';

/**
 * Extended user context information for telemetry
 */
export interface IExtendedUserContext {
  /** User's login name (domain\\username or email) */
  loginName: string;
  /** User's display name */
  displayName: string;
  /** User's email address */
  email: string;
  /** User's preferred language */
  preferredLanguage?: string;
  /** Whether the user is a site administrator */
  isSiteAdmin: boolean;
}

/**
 * Site context information for telemetry
 */
export interface ISiteContext {
  /** Current site URL */
  siteUrl: string;
  /** Site collection URL */
  siteCollectionUrl: string;
  /** Site title */
  siteTitle: string;
}

/**
 * Service for extracting comprehensive user and site context for telemetry purposes
 */
export class UserContextService {
  public static readonly serviceKey: ServiceKey<UserContextService> = ServiceKey.create<UserContextService>('SearchTelemetry:UserContextService', UserContextService);

  private _pageContext: PageContext;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
    });
  }

  /**
   * Gets basic user context information (fast, no API calls)
   */
  public getBasicUserContext(): Pick<IExtendedUserContext, 'loginName' | 'displayName' | 'email'> {
    const user = this._pageContext.user;
    return {
      loginName: user.loginName,
      displayName: user.displayName,
      email: user.email,
    };
  }

  /**
   * Gets extended user context information
   */
  public getExtendedUserContext(): IExtendedUserContext {
    const user = this._pageContext.user;

    return {
      loginName: user.loginName,
      displayName: user.displayName,
      email: user.email,
      preferredLanguage: this._pageContext.cultureInfo.currentCultureName,
      isSiteAdmin: (this._pageContext.legacyPageContext as any)?.isSiteAdmin || false,
    };
  }

  /**
   * Gets site context information
   */
  public getSiteContext(): ISiteContext {
    const web = this._pageContext.web;
    const site = this._pageContext.site;

    return {
      siteUrl: web.absoluteUrl,
      siteCollectionUrl: site.absoluteUrl,
      siteTitle: web.title,
    };
  }

  /**
   * Gets session context information
   */
  public getSessionContext(): {
    sessionId: string;
    browserInfo: string;
    timestamp: Date;
    pageUrl: string;
    referrer: string;
  } {
    return {
      sessionId: this._generateSessionId().sessionId,
      browserInfo: navigator.userAgent,
      timestamp: new Date(),
      pageUrl: window.location.href,
      referrer: document.referrer,
    };
  }

  /**
   * Generates a session identifier
   */
  private _generateSessionId(): { userHash: string; timestamp: number; sessionId: string } {
    // Create a session ID based on user, timestamp, and random component
    const timestamp = Date.now();
    const random = Math.random().toString(36).substring(2, 15);
    const userHash = this._pageContext.user.loginName.split('').reduce((a: number, b: string) => {
      a = (a << 5) - a + b.charCodeAt(0);
      return a & a;
    }, 0);

    const absUserHash = Math.abs(userHash).toString();
    return { userHash: absUserHash, timestamp, sessionId: `${absUserHash}-${timestamp}-${random}` };
  }
}
