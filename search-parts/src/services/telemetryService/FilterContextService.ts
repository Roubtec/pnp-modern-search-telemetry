import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { TelemetryLogger } from './ErrorHandling';

/**
 * FilterContextService - Extracts filter information from PnP Modern Search components
 *
 * URL Parameter Detection:
 * - PnP Modern Search hardcodes the filter parameter as 'f' in SearchFiltersContainer.tsx
 *   (see constant: DEEPLINK_QUERYSTRING_PARAM = 'f')
 * - This service defaults to 'f' but can be configured via filterUrlParameter
 * - As a fallback, it also checks 'filters' in case of custom implementations
 *
 * The filter URL parameter is NOT exposed in the Search Filters web part properties,
 * so runtime detection is not possible. Use the configuration option if you've customized
 * the PnP Modern Search source code.
 */

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
  /** Hashed values for privacy (array of hash strings) */
  hashedValues: string[];
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

/**
 * Configuration interface for the filter context service
 */
export interface IFilterContextConfiguration {
  /** The instance ID of the PnP Search Filters web part to monitor */
  filterWebPartId: string;
  /** The URL query parameter name used for filter deep linking (default: 'f' for PnP Modern Search) */
  filterUrlParameter: string;
}

/**
 * Service for extracting and managing filter context from PnP Modern Search components
 */
export class FilterContextService {
  public static readonly serviceKey: ServiceKey<FilterContextService> = ServiceKey.create<FilterContextService>('SearchTelemetry:FilterContextService', FilterContextService);

  private _urlObserver: MutationObserver | null = null;
  private _lastKnownFilterContext: IFilterContext | null = null;
  private _configuration: IFilterContextConfiguration;

  constructor(serviceScope: ServiceScope) {
    // Initialize default configuration
    this._configuration = {
      filterWebPartId: '76abee26-57ed-47ad-b309-6f514f50e6d7', // Default PnP Search Filters web part ID
      filterUrlParameter: 'f', // PnP Modern Search standard parameter name (hardcoded in SearchFiltersContainer.tsx)
    };

    // Initialize observer when service scope is ready
    serviceScope.whenFinished(() => {
      this._initializeFilterObserver();
    });
  }

  /**
   * Updates the filter context service configuration
   * @param config New configuration settings
   *
   * @example
   * // Use default PnP Modern Search parameter (recommended)
   * filterContextService.updateConfiguration({
   *   filterUrlParameter: 'f'  // This is the default
   * });
   *
   * @example
   * // Use custom parameter if you've modified PnP Modern Search source
   * filterContextService.updateConfiguration({
   *   filterUrlParameter: 'filters'
   * });
   *
   * @remarks
   * The filterUrlParameter defaults to 'f' which is hardcoded in PnP Modern Search
   * (see SearchFiltersContainer.tsx DEEPLINK_QUERYSTRING_PARAM constant).
   * Only change this if you've customized the PnP Modern Search source code.
   */
  public updateConfiguration(config: Partial<IFilterContextConfiguration>): void {
    this._configuration = { ...this._configuration, ...config };

    TelemetryLogger.info('FilterContextService configuration updated:', {
      filterWebPartId: this._configuration.filterWebPartId ? 'SET' : 'NOT_SET',
      filterUrlParameter: this._configuration.filterUrlParameter || 'default',
    });
  }

  /**
   * Gets the current configuration
   */
  public getConfiguration(): IFilterContextConfiguration {
    return { ...this._configuration };
  }

  /**
   * Gets the current filter context by analyzing the page state
   */
  public getCurrentFilterContext(): IFilterContext {
    try {
      const context: IFilterContext = {
        appliedFilters: [],
        filtersOperator: 'AND',
        totalActiveFilters: 0,
        filterCategories: [],
      };

      // Extract filters from URL parameters
      const urlFilters = this._extractFiltersFromUrl();
      if (urlFilters) {
        context.appliedFilters = urlFilters.filters;
        context.filtersOperator = urlFilters.operator;
        context.filterUrlData = urlFilters.rawData;
      }

      // Extract filters from DOM elements (if available)
      const domFilters = this._extractFiltersFromDOM();
      if (domFilters.length > 0) {
        context.appliedFilters = this._mergeFilterData(context.appliedFilters, domFilters);
      }

      // Find web part instance ID
      context.webPartInstanceId = this._findFilterWebPartInstanceId();

      // Calculate derived properties
      context.totalActiveFilters = context.appliedFilters.length;
      context.filterCategories = Array.from(new Set(context.appliedFilters.map((f) => f.filterType)));

      // Cache the result
      this._lastKnownFilterContext = context;

      return context;
    } catch (error) {
      TelemetryLogger.warn('Failed to get current filter context:', error);

      // Return empty context on error
      return {
        appliedFilters: [],
        filtersOperator: 'AND',
        totalActiveFilters: 0,
        filterCategories: [],
      };
    }
  }

  /**
   * Converts filter context to a simplified object for telemetry
   */
  public getFilterTelemetryData(context?: IFilterContext): IFilterTelemetryData {
    const filterContext = context || this.getCurrentFilterContext();

    return {
      hasActiveFilters: filterContext.totalActiveFilters > 0,
      totalActiveFilters: filterContext.totalActiveFilters,
      filtersOperator: filterContext.filtersOperator,
      filterCategories: filterContext.filterCategories,
      filterFields: filterContext.appliedFilters.map((f) => f.filterField),
      filterTypes: Array.from(new Set(filterContext.appliedFilters.map((f) => f.filterType))),
      multiValueFilters: filterContext.appliedFilters.filter((f) => f.isMultiValue).length,
      singleValueFilters: filterContext.appliedFilters.filter((f) => !f.isMultiValue).length,
      filterDetails: filterContext.appliedFilters.map((f) => ({
        field: f.filterField,
        displayName: f.displayName,
        type: f.filterType,
        valueCount: f.selectedValues.length,
        isMultiValue: f.isMultiValue,
        valuesOperator: f.valuesOperator,
        hashedValues: f.selectedValues.map((v) => this._hashValue(v)),
      })),
      webPartInstanceId: filterContext.webPartInstanceId,
      filterUrlPresent: !!filterContext.filterUrlData,
    };
  }

  /**
   * Extracts filter information from URL parameters (PnP Modern Search deep linking)
   */
  private _extractFiltersFromUrl(): {
    filters: IAppliedFilter[];
    operator: 'AND' | 'OR';
    rawData: string;
  } | null {
    try {
      const urlParams = new URLSearchParams(window.location.search);
      const hashParams = new URLSearchParams(window.location.hash.substring(1));

      // Check both URL params and hash params for filter data
      // Use configured parameter name first, then fall back to common alternatives
      const primaryParam = this._configuration.filterUrlParameter || 'f';
      let filterData =
        urlParams.get(primaryParam) || hashParams.get(primaryParam) || urlParams.get('f') || urlParams.get('filters') || hashParams.get('f') || hashParams.get('filters');

      TelemetryLogger.debug('FilterContextService: URL filter extraction', {
        configuredParam: primaryParam,
        hasFilterData: !!filterData,
        filterDataLength: filterData?.length || 0,
        url: window.location.href,
      });

      if (!filterData) {
        return null;
      }

      const filters: IAppliedFilter[] = [];
      let operator: 'AND' | 'OR' = 'AND';

      // Try to parse the filter data (format varies by PnP Modern Search version)
      try {
        const parsedData = JSON.parse(decodeURIComponent(filterData));

        if (Array.isArray(parsedData)) {
          // Array format
          parsedData.forEach((filterItem: any) => {
            filters.push(this._parseFilterItem(filterItem));
          });
        } else if (typeof parsedData === 'object') {
          // Object format
          operator = parsedData.operator || 'AND';
          if (parsedData.filters && Array.isArray(parsedData.filters)) {
            parsedData.filters.forEach((filterItem: any) => {
              filters.push(this._parseFilterItem(filterItem));
            });
          }
        }
      } catch (parseError) {
        TelemetryLogger.debug('Failed to parse filter URL data:', parseError);
      }

      TelemetryLogger.debug('FilterContextService: Successfully extracted filters from URL', {
        filterCount: filters.length,
        operator,
        filterFields: filters.map((f) => f.filterField),
      });

      return {
        filters,
        operator,
        rawData: filterData,
      };
    } catch (error) {
      TelemetryLogger.debug('Error extracting filters from URL:', error);
      return null;
    }
  }

  /**
   * Parses a single filter item from URL data
   */
  private _parseFilterItem(filterItem: any): IAppliedFilter {
    // Handle PnP Modern Search filter structure
    let selectedValues: string[] = [];
    let valuesOperator: 'AND' | 'OR' = 'OR';

    if (Array.isArray(filterItem.values)) {
      // Extract actual values from the PnP structure: values[{name, value, operator}]
      selectedValues = filterItem.values.map((v: any) => v.name || v.value || v.toString());
    } else if (filterItem.value) {
      selectedValues = [filterItem.value];
    }

    // Map PnP operator format
    if (filterItem.operator) {
      if (typeof filterItem.operator === 'string') {
        valuesOperator = filterItem.operator.toUpperCase() as 'AND' | 'OR';
      } else {
        // PnP sometimes uses numeric operators: 0 = OR, 1 = AND
        valuesOperator = filterItem.operator === 1 ? 'AND' : 'OR';
      }
    }

    return {
      filterField: filterItem.filterName || filterItem.field || filterItem.name || '',
      displayName: filterItem.displayName || filterItem.filterName || filterItem.field || '',
      selectedValues,
      valuesOperator,
      filterType: this._determineFilterType(filterItem),
      isMultiValue: selectedValues.length > 1,
    };
  }

  /**
   * Extracts filter information from DOM elements
   */
  private _extractFiltersFromDOM(): IAppliedFilter[] {
    const filters: IAppliedFilter[] = [];

    try {
      // Look for common filter selectors in PnP Modern Search
      const filterSelectors = ['[data-is-filter-value="true"]', '.pnp-filter-value-selected', '[data-automation-id*="filter"]'];

      filterSelectors.forEach((selector) => {
        const elements = document.querySelectorAll(selector);
        elements.forEach((element: HTMLElement) => {
          const filterInfo = this._extractFilterInfoFromElement(element);
          if (filterInfo) {
            filters.push(filterInfo);
          }
        });
      });
    } catch (error) {
      TelemetryLogger.debug('Error extracting filters from DOM:', error);
    }

    return filters;
  }

  /**
   * Extracts filter information from a DOM element
   */
  private _extractFilterInfoFromElement(element: HTMLElement): IAppliedFilter | null {
    try {
      const field = element.getAttribute('data-filter-field') || element.getAttribute('data-field') || 'Unknown';
      const value = element.getAttribute('data-filter-value') || element.textContent?.trim() || '';

      return {
        filterField: field,
        displayName: field,
        selectedValues: [value],
        valuesOperator: 'OR',
        filterType: this._determineFilterTypeFromElement(element),
        isMultiValue: false,
      };
    } catch (error) {
      TelemetryLogger.debug('Error extracting filter info from element:', error);
      return null;
    }
  }

  /**
   * Merges filter data from different sources (URL and DOM)
   */
  private _mergeFilterData(urlFilters: IAppliedFilter[], domFilters: IAppliedFilter[]): IAppliedFilter[] {
    const mergedFilters = [...urlFilters];

    for (const domFilter of domFilters) {
      const existingIndex = mergedFilters.findIndex((f) => f.filterField === domFilter.filterField);
      if (existingIndex === -1) {
        // Add new filter
        mergedFilters.push(domFilter);
      } else {
        // Merge values
        const existingFilter = mergedFilters[existingIndex];
        const combinedValues = Array.from(new Set([...existingFilter.selectedValues, ...domFilter.selectedValues]));
        mergedFilters[existingIndex] = {
          ...existingFilter,
          selectedValues: combinedValues,
          isMultiValue: combinedValues.length > 1,
        };
      }
    }

    return mergedFilters;
  }

  /**
   * Determines filter type from filter data object
   */
  private _determineFilterType(filterData: any): IAppliedFilter['filterType'] {
    if (filterData.type) {
      const type = filterData.type.toLowerCase();
      if (type.includes('date')) return 'date';
      if (type.includes('people')) return 'people';
      if (type.includes('combo')) return 'combo';
      if (type.includes('static')) return 'static';
    }

    // Fallback based on field name patterns
    const fieldName = (filterData.filterName || filterData.field || '').toLowerCase();
    if (fieldName.includes('date') || fieldName.includes('created') || fieldName.includes('modified')) {
      return 'date';
    }
    if (fieldName.includes('author') || fieldName.includes('people') || fieldName.includes('user')) {
      return 'people';
    }

    return 'refiner'; // Default
  }

  /**
   * Determines filter type from DOM element
   */
  private _determineFilterTypeFromElement(element: HTMLElement): IAppliedFilter['filterType'] {
    // Check element type and classes for hints
    const inputElement = element as HTMLInputElement;
    if (inputElement.type === 'checkbox') return 'refiner';
    if (inputElement.type === 'date') return 'date';
    if (element.classList.contains('people-picker')) return 'people';
    if (element.tagName.toLowerCase() === 'select') return 'combo';

    return 'refiner'; // Default
  }

  /**
   * Tries to find the Search Filters web part instance ID from the DOM
   */
  private _findFilterWebPartInstanceId(): string | undefined {
    try {
      // Look for PnP Modern Search web part containers
      const webPartContainers = document.querySelectorAll('[data-sp-feature-tag*="PnPModernSearch"]');
      for (let i = 0; i < webPartContainers.length; i++) {
        const container = webPartContainers[i];
        const instanceId = container.getAttribute('data-sp-webpart-instance-id');
        if (instanceId) {
          return instanceId;
        }
      }
    } catch (error) {
      TelemetryLogger.debug('Error finding filter web part instance ID:', error);
    }

    return undefined;
  }

  /**
   * Initializes observers to detect filter changes
   */
  private _initializeFilterObserver(): void {
    try {
      // Observe URL changes
      window.addEventListener('hashchange', () => this._onUrlChange());
      window.addEventListener('popstate', () => this._onUrlChange());

      // Observe DOM changes for filter updates
      this._urlObserver = new MutationObserver(() => this._onDOMChange());

      // Start observing the document with the configured parameters
      this._urlObserver.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        attributeFilter: ['data-is-filter-value', 'data-filter-value'],
      });

      TelemetryLogger.debug('Filter observer initialized');
    } catch (error) {
      TelemetryLogger.warn('Failed to initialize filter observer:', error);
    }
  }

  /**
   * Handles URL changes that might indicate filter updates
   */
  private _onUrlChange(): void {
    try {
      this._lastKnownFilterContext = this.getCurrentFilterContext();
    } catch (error) {
      TelemetryLogger.debug('Error handling URL change:', error);
    }
  }

  /**
   * Handles DOM changes that might indicate filter updates
   */
  private _onDOMChange(): void {
    try {
      // Debounce DOM change detection to avoid excessive processing
      // We'll just update the cached context
      const currentContext = this.getCurrentFilterContext();

      // Only update if there's a meaningful change
      if (JSON.stringify(currentContext) !== JSON.stringify(this._lastKnownFilterContext)) {
        this._lastKnownFilterContext = currentContext;
        TelemetryLogger.debug('Filter context updated from DOM change');
      }
    } catch (error) {
      TelemetryLogger.debug('Error handling DOM change:', error);
    }
  }

  /**
   * Creates a simple hash of a value for privacy
   */
  private _hashValue(value: string): string {
    let hash = 0;
    if (value.length === 0) return '0';

    for (let i = 0; i < value.length; i++) {
      const char = value.charCodeAt(i);
      hash = (hash << 5) - hash + char;
      hash = hash & hash; // Convert to 32bit integer
    }

    return Math.abs(hash).toString(16);
  }

  /**
   * Cleanup method
   */
  public dispose(): void {
    if (this._urlObserver) {
      this._urlObserver.disconnect();
      this._urlObserver = null;
    }
  }
}
