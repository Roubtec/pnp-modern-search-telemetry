/**
 * Telemetry error codes
 */
export enum TelemetryErrorCode {
  INVALID_CONFIG = 'INVALID_CONFIG',
  INVALID_QUERY = 'INVALID_QUERY',
  NETWORK_ERROR = 'NETWORK_ERROR',
  TIMEOUT = 'TIMEOUT',
  UNKNOWN = 'UNKNOWN',
}

/**
 * Custom error class for telemetry operations
 */
export class TelemetryError extends Error {
  public code: TelemetryErrorCode;
  public details?: any;

  constructor(message: string, code: TelemetryErrorCode = TelemetryErrorCode.UNKNOWN, details?: any) {
    super(message);
    this.name = 'TelemetryError';
    this.code = code;
    this.details = details;

    // Maintains proper stack trace for where our error was thrown (only available on V8)
    if ((Error as any).captureStackTrace) {
      (Error as any).captureStackTrace(this, TelemetryError);
    }
  }
}

/**
 * Validator class for telemetry data
 */
export class TelemetryValidator {
  /**
   * Validates an endpoint URL
   */
  public static validateEndpointUrl(url: string): boolean {
    if (!url || typeof url !== 'string') {
      return false;
    }

    // Basic URL validation
    try {
      const urlObj = new URL(url);
      return urlObj.protocol === 'http:' || urlObj.protocol === 'https:';
    } catch {
      return false;
    }
  }

  /**
   * Validates an API key
   */
  public static validateApiKey(apiKey: string): boolean {
    if (!apiKey || typeof apiKey !== 'string') {
      return false;
    }

    // Basic validation - should be non-empty and not too short
    return apiKey.trim().length >= 8;
  }

  /**
   * Validates query text
   */
  public static validateQueryText(queryText: string): boolean {
    if (queryText === null || queryText === undefined) {
      return false;
    }

    // Allow empty strings as they represent valid searches
    return typeof queryText === 'string';
  }

  /**
   * Validates telemetry data payload
   */
  public static validateTelemetryData(data: any): boolean {
    if (!data || typeof data !== 'object') {
      return false;
    }

    // Check required fields
    const requiredFields = ['queryText', 'userId', 'siteUrl', 'pageUrl', 'timestamp', 'metadata'];
    return requiredFields.every((field) => field in data);
  }
}

/**
 * Logger class for telemetry operations
 */
export class TelemetryLogger {
  private static _loggingEnabled: boolean = false;

  /**
   * Enables or disables logging
   */
  public static setLoggingEnabled(enabled: boolean): void {
    TelemetryLogger._loggingEnabled = enabled;
  }

  /**
   * Logs an info message
   */
  public static info(message: string, ...args: any[]): void {
    if (TelemetryLogger._loggingEnabled) {
      console.log(`[TelemetryService] INFO: ${message}`, ...args);
    }
  }

  /**
   * Logs a warning message
   */
  public static warn(message: string, ...args: any[]): void {
    if (TelemetryLogger._loggingEnabled) {
      console.warn(`[TelemetryService] WARN: ${message}`, ...args);
    }
  }

  /**
   * Logs an error message
   */
  public static error(message: string, error?: Error | any, ...args: any[]): void {
    if (TelemetryLogger._loggingEnabled) {
      console.error(`[TelemetryService] ERROR: ${message}`, error, ...args);
    }
  }

  /**
   * Logs a debug message
   */
  public static debug(message: string, ...args: any[]): void {
    if (TelemetryLogger._loggingEnabled) {
      console.debug(`[TelemetryService] DEBUG: ${message}`, ...args);
    }
  }
}
