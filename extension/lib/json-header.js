// "Show _api pages as JSON": a dynamic declarativeNetRequest rule that sets the Accept
// header of page loads (main frame only) of SharePoint REST URLs, so the browser shows
// JSON instead of XML. Requests made by the extension itself are not page loads.
// Pure: no chrome.* APIs.

export const JSON_RULE_ID = 1;
export const JSON_ACCEPT = 'application/json;odata=nometadata';
// https://<tenant>.sharepoint.com/<anything>/_api/...
export const JSON_URL_REGEX = '^https://[a-z0-9-]+\\.sharepoint\\.com/(?:[^?#]*/)?_api/';

export function jsonRule() {
  return {
    id: JSON_RULE_ID,
    priority: 1,
    action: { type: 'modifyHeaders', requestHeaders: [{ header: 'Accept', operation: 'set', value: JSON_ACCEPT }] },
    condition: { regexFilter: JSON_URL_REGEX, resourceTypes: ['main_frame'] },
  };
}
