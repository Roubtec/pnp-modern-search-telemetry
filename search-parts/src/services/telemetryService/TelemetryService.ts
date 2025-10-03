import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { FilterContextService, IFilterContext, IFilterTelemetryData } from './FilterContextService';
import { TelemetryError, TelemetryErrorCode, TelemetryValidator, TelemetryLogger } from './ErrorHandling';

const whitespaceWordBoundaryRegex = /\s+/;
const specialCharRegex = /[!@#$%^&*()_+\-=[\]{};':"|,.<>?/~`]/;

/**
 * Interface defining the structure of telemetry data to be sent
 */
export interface ITelemetryData {
  /** The search query text entered by the user */
  queryText: string;
  /** Current user's unique identifier */
  userId: string;
  /** Current user's display name */
  userDisplayName: string;
  /** Current user's email address */
  userEmail: string;
  /** SharePoint site URL where the search was performed */
  siteUrl: string;
  /** Page URL where the search was performed */
  pageUrl: string;
  /** Timestamp when the search was performed */
  timestamp: Date;
  /** Additional metadata about the search context */
  metadata: {
    /** Number of characters in the query */
    queryLength: number;
    /** Number of words in the query */
    wordCount: number;
    /** Whether the query contains special characters */
    hasSpecialChars: boolean;
    /** User agent string */
    userAgent: string;
    /** Session identifier (if available) */
    sessionId?: string;
    searchContext: {
      /** PnP Modern Search filter information */
      filters: IFilterTelemetryData;
      /** Query enhancement context */
      queryEnhancement: {
        hasFiltersApplied: boolean;
        isSearchWithFilters: boolean;
        isFilterOnlySearch: boolean;
        isPureTextSearch: boolean;
      };
      /** Page and component context */
      pageContext: {
        hasSearchFiltersWebPart: boolean;
        isKnownFilterWebPartInstanceId: boolean;
        hasFilterDataInUrl: boolean;
      };
    };
    /** Any additional metadata provided */
    [key: string]: any;
  };
}

/**
 * Configuration interface for the telemetry service
 */
export interface ITelemetryConfiguration {
  /** External endpoint URL where telemetry data will be sent */
  endpointUrl: string;
  /** Whether telemetry collection is enabled */
  enabled: boolean;
  /** API key or authorization header for the external endpoint */
  apiKey?: string;
  /** Timeout in milliseconds for HTTP requests (default: 5000) */
  timeout?: number;
  /** Whether to include user personal information (email, display name) */
  includePersonalInfo?: boolean;
  /** Whether to log telemetry operations for debugging */
  enableLogging?: boolean;
  /** The instance ID of the PnP Search Filters web part to monitor */
  filterWebPartId?: string;
}

/**
 * Service for collecting and sending search telemetry data to external endpoints
 */
export class TelemetryService {
  public static readonly serviceKey: ServiceKey<TelemetryService> = ServiceKey.create<TelemetryService>('SearchTelemetry:TelemetryService', TelemetryService);

  private _httpClient: HttpClient;
  private _pageContext: PageContext;
  private _configuration: ITelemetryConfiguration;
  private _filterContextService: FilterContextService;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._httpClient = serviceScope.consume(HttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._filterContextService = serviceScope.consume(FilterContextService.serviceKey);
    });

    // Default configuration
    this._configuration = {
      endpointUrl: '',
      enabled: false,
      timeout: 5000,
      includePersonalInfo: true,
      enableLogging: false,
      filterWebPartId: '76abee26-57ed-47ad-b309-6f514f50e6d7', // Default PnP Search Filters web part ID
    };
  }

  /**
   * Updates the telemetry service configuration
   * @param config New configuration settings
   */
  public updateConfiguration(config: Partial<ITelemetryConfiguration>): void {
    // Validate configuration
    if (config.endpointUrl && !TelemetryValidator.validateEndpointUrl(config.endpointUrl)) {
      throw new TelemetryError('Invalid endpoint URL provided', TelemetryErrorCode.INVALID_CONFIG, {
        endpointUrl: config.endpointUrl,
      });
    }

    if (config.apiKey && !TelemetryValidator.validateApiKey(config.apiKey)) {
      throw new TelemetryError('Invalid API key provided', TelemetryErrorCode.INVALID_CONFIG);
    }

    this._configuration = { ...this._configuration, ...config };

    // Update logger settings
    TelemetryLogger.setLoggingEnabled(this._configuration.enableLogging || false);

    // Update FilterContextService configuration if available
    if (this._filterContextService && config.filterWebPartId) {
      this._filterContextService.updateConfiguration({
        filterWebPartId: config.filterWebPartId,
      });
    }

    TelemetryLogger.info('Configuration updated:', {
      enabled: this._configuration.enabled,
      endpointUrl: this._configuration.endpointUrl ? 'SET' : 'NOT_SET',
      apiKey: this._configuration.apiKey ? 'SET' : 'NOT_SET',
      includePersonalInfo: this._configuration.includePersonalInfo,
      enableLogging: this._configuration.enableLogging,
      filterWebPartId: this._configuration.filterWebPartId ? 'SET' : 'NOT_SET',
    });
  }

  /**
   * Gets the current configuration (excluding sensitive data like API keys)
   */
  public getConfiguration(): Omit<ITelemetryConfiguration, 'apiKey'> {
    // eslint-disable-next-line @typescript-eslint/no-unused-vars
    const { apiKey, ...safeConfig } = this._configuration;
    return safeConfig;
  }

  /**
   * Captures and sends telemetry data for a search query
   * @param queryText The search query text
   * @param additionalMetadata Optional additional metadata
   */
  public async captureSearchQuery(queryText: string, additionalMetadata?: any): Promise<void> {
    // Early return if telemetry is disabled
    if (!this._configuration.enabled) {
      TelemetryLogger.debug('Telemetry disabled, skipping capture');
      return;
    }

    // Early return if no endpoint is configured
    if (!this._configuration.endpointUrl) {
      TelemetryLogger.warn('No endpoint URL configured, skipping telemetry');
      return;
    }

    // Validate query text
    if (!TelemetryValidator.validateQueryText(queryText)) {
      TelemetryLogger.warn('Invalid query text provided, skipping telemetry');
      return;
    }

    try {
      // Store original query length before any processing
      const originalQueryLength = queryText.length;

      // Build the telemetry data
      const telemetryData = this._buildTelemetryData(queryText, originalQueryLength, additionalMetadata);

      // Validate the telemetry data
      if (!TelemetryValidator.validateTelemetryData(telemetryData)) {
        throw new TelemetryError('Invalid telemetry data structure', TelemetryErrorCode.INVALID_CONFIG);
      }

      TelemetryLogger.info('Sending telemetry data:', {
        queryLength: queryText.length,
      });

      // Send the data to the endpoint (non-blocking)
      await this._sendTelemetryData(telemetryData);

      TelemetryLogger.info('Telemetry data sent successfully');
    } catch (error) {
      // Log the error but don't throw to avoid breaking search functionality
      if (error instanceof TelemetryError) {
        TelemetryLogger.error('Telemetry error:', error, {
          code: error.code,
          details: error.details,
        });
      } else {
        TelemetryLogger.error('Unexpected error capturing telemetry:', error);
      }

      // Re-throw network errors for retry logic if needed
      if (error instanceof TelemetryError && error.code === TelemetryErrorCode.NETWORK_ERROR) {
        throw error;
      }
    }
  }

  /**
   * Tests the connection to the configured telemetry endpoint
   */
  public async testConnection(): Promise<boolean> {
    if (!this._configuration.endpointUrl) {
      TelemetryLogger.warn('No endpoint URL configured for connection test');
      return false;
    }

    try {
      TelemetryLogger.info('Testing connection to telemetry endpoint:', this._configuration.endpointUrl);

      const response: HttpClientResponse = await this._httpClient.get(this._configuration.endpointUrl, HttpClient.configurations.v1, {
        headers: this._buildHeaders(),
      });

      const success = response.ok || response.status === 404; // 404 is acceptable for GET on a POST-only endpoint
      if (success) {
        TelemetryLogger.info('Connection test successful');
      } else {
        TelemetryLogger.warn('Connection test failed:', response.status, response.statusText);
      }

      return success;
    } catch (error) {
      TelemetryLogger.error('Connection test failed with error:', error);
      return false;
    }
  }

  /**
   * Builds the telemetry data object from the search query and context
   */
  private _buildTelemetryData(queryText: string, originalQueryLength: number, additionalMetadata?: any): ITelemetryData {
    const currentUser = this._pageContext.user;
    const web = this._pageContext.web;

    // Capture filter context for enhanced telemetry
    let filterContext: IFilterContext | null = null;
    let filterTelemetryData: IFilterTelemetryData | any = {};

    try {
      if (this._filterContextService) {
        filterContext = this._filterContextService.getCurrentFilterContext();
        filterTelemetryData = this._filterContextService.getFilterTelemetryData(filterContext);
      }
    } catch (error) {
      TelemetryLogger.warn('Failed to capture filter context:', error);
    }

    const telemetryData: ITelemetryData = {
      queryText,
      userId: currentUser.loginName,
      userDisplayName: this._configuration.includePersonalInfo ? currentUser.displayName : '[REDACTED]',
      userEmail: this._configuration.includePersonalInfo ? currentUser.email : '[REDACTED]',
      siteUrl: web.absoluteUrl,
      pageUrl: window.location.href,
      timestamp: new Date(),
      metadata: {
        queryLength: originalQueryLength,
        wordCount: queryText.trim().split(whitespaceWordBoundaryRegex).length,
        hasSpecialChars: specialCharRegex.test(queryText),
        userAgent: navigator.userAgent,
        sessionId: this._generateSessionId(),
        searchContext: {
          filters: filterTelemetryData,
          queryEnhancement: {
            hasFiltersApplied: filterTelemetryData.hasActiveFilters || false,
            isSearchWithFilters: (filterTelemetryData.hasActiveFilters || false) && queryText.trim().length > 0,
            isFilterOnlySearch: (filterTelemetryData.hasActiveFilters || false) && queryText.trim().length === 0,
            isPureTextSearch: !(filterTelemetryData.hasActiveFilters || false) && queryText.trim().length > 0,
          },
          pageContext: {
            hasSearchFiltersWebPart: !!filterTelemetryData.webPartInstanceId,
            isKnownFilterWebPartInstanceId: filterTelemetryData.webPartInstanceId === '76abee26-57ed-47ad-b309-6f514f50e6d7',
            hasFilterDataInUrl: filterTelemetryData.filterUrlPresent || false,
          },
        },
        ...additionalMetadata,
      },
    };

    return telemetryData;
  }

  /**
   * Sends telemetry data to the configured external endpoint
   */
  private async _sendTelemetryData(data: ITelemetryData): Promise<void> {
    const response: HttpClientResponse = await this._httpClient.post(this._configuration.endpointUrl, HttpClient.configurations.v1, {
      headers: this._buildHeaders(),
      body: JSON.stringify(data),
    });

    if (!response.ok) {
      throw new TelemetryError(`Failed to send telemetry data: ${response.statusText}`, TelemetryErrorCode.NETWORK_ERROR, { status: response.status });
    }
  }

  /**
   * Builds HTTP headers for telemetry requests
   */
  private _buildHeaders(): { [key: string]: string } {
    const headers: { [key: string]: string } = {
      'Content-Type': 'application/json',
      'User-Agent': 'PnP-Modern-Search-Telemetry/1.0.0',
    };

    // Add authorization header if API key is configured
    if (this._configuration.apiKey) {
      headers['Authorization'] = `Bearer ${this._configuration.apiKey}`;
    }

    return headers;
  }

  /**
   * Generates a session identifier for grouping related searches
   */
  private _generateSessionId(): string {
    // Simple session ID based on user and timestamp
    const timestamp = Date.now();
    const userHash = this._pageContext.user.loginName.split('').reduce((a: number, b: string) => {
      a = (a << 5) - a + b.charCodeAt(0);
      return a & a;
    }, 0);

    return Math.abs(userHash) + '-' + timestamp;
  }
}
