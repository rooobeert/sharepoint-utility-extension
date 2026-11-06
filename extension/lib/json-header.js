// "JSON" tab: a dynamic declarativeNetRequest rule that sets the Accept header of page
// loads (main frame only) of SharePoint URLs whose path contains one of the listed texts
// (default "/_api/"), so the browser shows JSON instead of XML. Requests made by the
// extension itself are not page loads. Pure: no chrome.* APIs.

export const JSON_RULE_ID = 1;
export const JSON_ACCEPT = 'application/json;odata=nometadata';
export const DEFAULT_PATHS = ['/_api/'];

const escapeRegex = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

/** A path text as typed -> the text stored, or null if it can't be used (empty, spaces, ? or #). */
export function normalizePath(input) {
  const s = String(input || '').trim();
  return s && !/[\s?#]/.test(s) ? s : null;
}

/** Add a path (no duplicates, ignoring case). Returns the new list. */
export function addPath(paths, input) {
  const p = normalizePath(input);
  if (!p || paths.some((x) => x.toLowerCase() === p.toLowerCase())) return paths;
  return [...paths, p];
}

/** https://<tenant>.sharepoint.com/ followed by a path that contains one of the texts. */
export function pathsRegex(paths) {
  return `^https://[a-z0-9-]+\\.sharepoint\\.com(?:/[^?#]*)?(?:${paths.map(escapeRegex).join('|')})`;
}

/** The rule for the list, or null when the list is empty. */
export function jsonRule(paths = DEFAULT_PATHS) {
  const list = paths.map(normalizePath).filter(Boolean);
  if (!list.length) return null;
  return {
    id: JSON_RULE_ID,
    priority: 1,
    action: { type: 'modifyHeaders', requestHeaders: [{ header: 'Accept', operation: 'set', value: JSON_ACCEPT }] },
    condition: { regexFilter: pathsRegex(list), isUrlFilterCaseSensitive: false, resourceTypes: ['main_frame'] },
  };
}
