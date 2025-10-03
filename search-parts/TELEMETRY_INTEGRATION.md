# Search Telemetry Integration

This document describes the integrated search telemetry functionality in the PnP Modern Search - Search Results Web Part.

## Overview

The Search Results web part now includes built-in telemetry capabilities that can capture search queries, applied filters, and user context information, sending this data to an external analytics endpoint.

## Features

- **Query Capture**: Automatically captures search queries whenever results are displayed
- **Filter Context**: Detects and includes applied filters from PnP Search Filters web parts
- **User Context**: Captures user information (with privacy controls)
- **Site Context**: Includes SharePoint site information
- **Configurable Privacy**: Option to redact personal information
- **Debug Logging**: Enable detailed logging for troubleshooting

## Configuration

### Web Part Properties

In the Search Results web part property pane, navigate to the "Connections & Query" page. You'll find a "Search Telemetry" section with the following settings:

1. **Enable Search Telemetry** (Toggle)

   - Turn on/off telemetry collection
   - Default: Off

2. **Telemetry Endpoint URL** (Text)

   - The HTTPS endpoint where telemetry data will be sent
   - Example: `https://your-telemetry-endpoint.com/api/search-analytics`
   - Required when telemetry is enabled

3. **API Key** (Text, Optional)

   - Authorization key for the telemetry endpoint
   - Sent as `Authorization: Bearer {apiKey}` header
   - Leave empty if your endpoint doesn't require authentication

4. **Include Personal Information** (Toggle)

   - When enabled: Sends user display name and email
   - When disabled: Redacts personal information as `[REDACTED]`
   - Default: Enabled

5. **Enable Debug Logging** (Toggle)

   - Enables console logging for telemetry operations
   - Useful for troubleshooting
   - Default: Off

6. **Search Filters Web Part ID** (Text)
   - The instance ID of the PnP Search Filters web part to monitor
   - Default: `76abee26-57ed-47ad-b309-6f514f50e6d7`
   - Used to capture filter context from the page

## Telemetry Data Structure

The telemetry payload sent to your endpoint follows this structure:

```json
{
  "queryText": "search query text",
  "userId": "user@domain.com",
  "userDisplayName": "User Name",
  "userEmail": "user@domain.com",
  "siteUrl": "https://tenant.sharepoint.com/sites/site",
  "pageUrl": "https://tenant.sharepoint.com/sites/site/pages/page.aspx",
  "timestamp": "2025-10-03T10:30:00.000Z",
  "metadata": {
    "queryLength": 15,
    "wordCount": 2,
    "hasSpecialChars": false,
    "userAgent": "Mozilla/5.0...",
    "sessionId": "12345-1696328400000",
    "searchContext": {
      "filters": {
        "hasActiveFilters": true,
        "totalActiveFilters": 2,
        "filtersOperator": "AND",
        "filterCategories": ["refiner", "date"],
        "filterFields": ["FileType", "Modified"],
        "filterTypes": ["refiner", "date"],
        "multiValueFilters": 1,
        "singleValueFilters": 1,
        "filterDetails": [
          {
            "field": "FileType",
            "displayName": "File Type",
            "type": "refiner",
            "valueCount": 2,
            "isMultiValue": true,
            "valuesOperator": "OR",
            "hashedValues": ["abc123", "def456"]
          }
        ],
        "webPartInstanceId": "76abee26-57ed-47ad-b309-6f514f50e6d7",
        "filterUrlPresent": true
      },
      "queryEnhancement": {
        "hasFiltersApplied": true,
        "isSearchWithFilters": true,
        "isFilterOnlySearch": false,
        "isPureTextSearch": false
      },
      "pageContext": {
        "hasSearchFiltersWebPart": true,
        "isKnownFilterWebPartInstanceId": true,
        "hasFilterDataInUrl": true
      }
    },
    "resultsCount": 42,
    "pageNumber": 1,
    "dataSourceKey": "SharePointSearch"
  }
}
```

## Endpoint Requirements

Your telemetry endpoint should:

1. Accept POST requests with JSON payload
2. Return HTTP 200/201 for successful receipt
3. Support the `Content-Type: application/json` header
4. (Optional) Validate the `Authorization: Bearer {apiKey}` header if using API keys

Example endpoint implementation (Node.js/Express):

```javascript
app.post("/api/search-analytics", (req, res) => {
  // Validate API key if required
  const apiKey = req.headers.authorization?.replace("Bearer ", "");
  if (apiKey !== process.env.EXPECTED_API_KEY) {
    return res.status(401).send("Unauthorized");
  }

  // Process telemetry data
  const telemetryData = req.body;
  console.log("Received search telemetry:", telemetryData);

  // Store in database, send to analytics service, etc.
  // ...

  res.status(200).send("OK");
});
```

## Testing and Debugging

### Browser Console Commands

When debug logging is enabled or the web part is in edit mode, you can test telemetry from the browser console:

```javascript
// Get the web part instance (replace {instanceId} with your web part's instance ID)
var wp = window["SearchResultsWebPart_{instanceId}"];

// Check telemetry status
wp.getTelemetryStatus();

// Test connection to endpoint
await wp.testTelemetryConnection();

// Send a test telemetry event
await wp.sendTestTelemetry();
```

### Finding Your Instance ID

To find your web part's instance ID:

1. Open the page in edit mode
2. Open browser developer tools (F12)
3. In the console, type:
   ```javascript
   Object.keys(window).filter((k) => k.startsWith("SearchResultsWebPart_"));
   ```
4. This will list all Search Results web part instances on the page

## Privacy Considerations

- **Personal Information**: When "Include Personal Information" is disabled, user display names and email addresses are redacted as `[REDACTED]`
- **Filter Values**: Filter values are hashed using a simple hash function for privacy
- **No Content Storage**: The actual content of search results is not captured, only metadata
- **User Control**: Users can see what data is being sent in the browser's Network tab

## Troubleshooting

### Telemetry not being sent

1. Verify "Enable Search Telemetry" is turned on
2. Check that "Telemetry Endpoint URL" is configured
3. Enable "Debug Logging" to see console messages
4. Check browser Network tab for failed requests
5. Verify your endpoint is accessible and returning 200 OK

### Connection test fails

1. Verify the endpoint URL is correct (must start with https://)
2. Check CORS configuration on your endpoint (must allow requests from SharePoint domains)
3. Verify firewall/network policies allow outbound HTTPS connections
4. Test the endpoint with a tool like Postman

### No filter data captured

1. Verify "Search Filters Web Part ID" matches your filters web part
2. Check that filters are actually applied on the page
3. Enable debug logging to see filter detection messages

## Implementation Details

The telemetry integration consists of these key components:

- **TelemetryService**: Core service for sending telemetry data
- **FilterContextService**: Extracts filter information from the page
- **UserContextService**: Provides user and site context
- **SearchResultsContainer**: Triggers telemetry capture after data retrieval

The telemetry capture is **non-blocking** and designed to not impact search performance. If telemetry fails, it will not affect the search functionality.

## Architecture

```
SearchResultsWebPart
  ├── initializeWebPartServices()
  │   ├── TelemetryService
  │   ├── FilterContextService
  │   └── UserContextService
  │
  └── render()
      └── SearchResultsContainer
          └── getDataFromDataSource()
              └── onCaptureSearchTelemetry()
                  └── TelemetryService.captureSearchQuery()
                      ├── Build telemetry data
                      ├── Get filter context
                      ├── Get user context
                      └── POST to endpoint
```

## Future Enhancements

Possible future improvements:

- Batch multiple telemetry events for efficiency
- Retry logic for failed requests
- Client-side caching to avoid duplicate events
- Support for additional metadata fields
- Integration with Application Insights or other analytics platforms
- Anonymization options beyond simple redaction

## Support

For issues or questions about the telemetry integration, please file an issue in the GitHub repository.
