// "Headers" tab: a dynamic declarativeNetRequest rule that sets request headers on page
// loads (main frame only) of SharePoint URLs whose path contains one of the listed texts
// (default "/_api/"). Default header: Accept JSON, so the browser shows JSON instead of XML;
// Accept-Language and others can be added. Requests made by the extension itself and the
// Batch Runner are not page loads and are not touched. Pure: no chrome.* APIs.

export const RULE_ID = 1;
export const DEFAULT_PATHS = ['/_api/'];
export const DEFAULT_HEADERS = [{ name: 'Accept', value: 'application/json;odata=nometadata' }];

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

/** Header as typed -> { name, value }, or null (name must be a plain token, value text without line breaks). */
export function normalizeHeader(name, value) {
  const n = String(name || '').trim();
  const v = String(value ?? '').trim();
  if (!/^[A-Za-z0-9!#$%&'*+.^_`|~-]+$/.test(n) || !v || /[\r\n\0]/.test(v)) return null;
  return { name: n, value: v };
}

/** Add a header, or change the value of the one with the same name (ignoring case). Returns the new list. */
export function setHeader(headers, name, value) {
  const h = normalizeHeader(name, value);
  if (!h) return headers;
  const i = headers.findIndex((x) => x.name.toLowerCase() === h.name.toLowerCase());
  return i < 0 ? [...headers, h] : headers.map((x, k) => (k === i ? h : x));
}

/** https://<tenant>.sharepoint.com/ followed by a path that contains one of the texts. */
export function pathsRegex(paths) {
  return `^https://[a-z0-9-]+\\.sharepoint\\.com(?:/[^?#]*)?(?:${paths.map(escapeRegex).join('|')})`;
}

/** The rule for the lists, or null when either list is empty. */
export function headerRule(paths = DEFAULT_PATHS, headers = DEFAULT_HEADERS) {
  const list = paths.map(normalizePath).filter(Boolean);
  const set = headers.map((h) => normalizeHeader(h.name, h.value)).filter(Boolean);
  if (!list.length || !set.length) return null;
  return {
    id: RULE_ID,
    priority: 1,
    action: { type: 'modifyHeaders', requestHeaders: set.map((h) => ({ header: h.name, operation: 'set', value: h.value })) },
    condition: { regexFilter: pathsRegex(list), isUrlFilterCaseSensitive: false, resourceTypes: ['main_frame'] },
  };
}
