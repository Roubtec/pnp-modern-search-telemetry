/**
 * Interface representing a filter applied by the user
 */
export interface IAppliedFilter {
  /** The internal field name used for filtering */
  filterField: string;
  /** The display name of the filter */
  displayName: string;
  /** The selected values for this filter */
  selectedValues: string[];
  /** The operator used between values (AND/OR) */
  valuesOperator: 'AND' | 'OR';
  /** The type of filter (refiner, static, date, etc.) */
  filterType: 'refiner' | 'static' | 'date' | 'people' | 'combo';
  /** Whether multiple values are allowed */
  isMultiValue: boolean;
}

/**
 * Interface representing the complete filter context
 */
export interface IFilterContext {
  /** List of all applied filters */
  appliedFilters: IAppliedFilter[];
  /** The operator used between different filters (AND/OR) */
  filtersOperator: 'AND' | 'OR';
  /** The Search Filters web part instance ID */
  webPartInstanceId?: string;
  /** URL hash/query string containing filter data */
  filterUrlData?: string;
  /** Total number of active filters */
  totalActiveFilters: number;
  /** Categories of filters being used */
  filterCategories: string[];
}

/**
 * Interface representing detailed filter information for telemetry
 */
export interface IFilterTelemetryDetail {
  /** The internal field name */
  field: string;
  /** The display name of the filter */
  displayName: string;
  /** The type of filter */
  type: IAppliedFilter['filterType'];
  /** Number of selected values */
  valueCount: number;
  /** Whether multiple values are allowed */
  isMultiValue: boolean;
  /** The operator used between values */
  valuesOperator: 'AND' | 'OR';
  /** Values (array of strings) */
  values: string[];
}

/**
 * Interface representing telemetry data extracted from filter context
 */
export interface IFilterTelemetryData {
  /** Whether there are any active filters */
  hasActiveFilters: boolean;
  /** Total number of active filters */
  totalActiveFilters: number;
  /** The operator used between different filters */
  filtersOperator: 'AND' | 'OR';
  /** Categories of filters being used */
  filterCategories: string[];
  /** List of filter field names */
  filterFields: string[];
  /** List of filter types being used */
  filterTypes: IAppliedFilter['filterType'][];
  /** Number of multi-value filters */
  multiValueFilters: number;
  /** Number of single-value filters */
  singleValueFilters: number;
  /** Detailed information about each filter */
  filterDetails: IFilterTelemetryDetail[];
  /** The Search Filters web part instance ID */
  webPartInstanceId?: string;
  /** Whether filter URL data is present */
  filterUrlPresent: boolean;
}

export interface IQueryAnalysis {
  isEmpty: boolean;
  isWildcard: boolean;
  containsQuotes: boolean;
  containsOperators: boolean;
  containsWildcards: boolean;
  wordCount: number;
}

export interface ITechnicalContext {
  webPartType: string;
  pnpModernSearchVersion: string;
  captureMethod: string;
  browserLanguage: string;
  currentPath: string;
}

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
    queryAnalysis?: IQueryAnalysis;
    technicalContext?: ITechnicalContext;
    /** Any additional metadata provided */
    [key: string]: any;
  };
}
