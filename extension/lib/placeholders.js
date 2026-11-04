// Placeholders in batch bodies.
//
// ID placeholders, resolved just before the operations that use them run:
//   {{WebId}}                         -> web GUID
//   {{ListId:<reference>}}            -> list GUID (no braces). The reference may be
//       a title                       {{ListId:Change Requests}}
//       a site-relative URL           {{ListId:Lists/crRequests}}, {{ListId:crDocumentVault}}
//       a server-relative URL         {{ListId:/sites/x/Lists/crRequests}}
//       a full URL                    {{ListId:https://tenant.sharepoint.com/sites/x/Lists/crRequests}}
//   @{variables('<any name>')?['<reference>']}   Power Automate form, same references
//
// URL placeholders, replaced with the target site before anything else:
//   {{WebUrl}} {{WebServerRelativeUrl}} {{SiteUrl}} {{SiteServerRelativeUrl}}
//   (no trailing slash; the server-relative forms are empty for a root site)

const ID_PATTERNS = [
  { re: /\{\{\s*WebId\s*\}\}/gi, toRef: () => ({ type: 'web', ref: '' }) },
  { re: /\{\{\s*ListId\s*:\s*([^{}]+?)\s*\}\}/gi, toRef: (m) => ({ type: 'list', ref: m[1] }) },
  { re: /@\{\s*variables\('[^']+'\)\?\['([^'\]]+)'\]\s*\}/g, toRef: (m) => ({ type: 'list', ref: m[1] }) },
];

export const URL_PLACEHOLDER_NAMES = ['WebUrl', 'WebServerRelativeUrl', 'SiteUrl', 'SiteServerRelativeUrl'];
const URL_RE = /\{\{\s*(WebUrl|WebServerRelativeUrl|SiteUrl|SiteServerRelativeUrl)\s*\}\}/gi;
const canonicalUrlName = (name) => URL_PLACEHOLDER_NAMES.find((n) => n.toLowerCase() === name.toLowerCase());

/**
 * How a list reference is matched:
 * - 'url'      starts with "/" or "https://": server-relative or full URL
 * - 'relative' contains "/": URL relative to the web ("Lists/crRequests")
 * - 'name'     a single word: list title, or a library's URL name ("crDocumentVault")
 */
export function listRefKind(ref) {
  const r = ref.trim();
  if (/^(\/|https?:\/\/)/i.test(r)) return 'url';
  return r.includes('/') ? 'relative' : 'name';
}

function normalizeListRef(ref) {
  let r = ref.trim();
  const kind = listRefKind(r);
  if (kind === 'url' && /^https?:\/\//i.test(r)) r = new URL(r).pathname;
  if (kind !== 'name') r = safeDecode(r).replace(/^\.\//, '').replace(/\/+$/, '');
  return r;
}

function safeDecode(s) {
  try {
    return decodeURIComponent(s);
  } catch {
    return s;
  }
}

/** Stable key so the same list referenced twice (in any syntax) is looked up once. */
export function placeholderKey(type, ref) {
  return type === 'web' ? 'web' : `list:${listRefKind(ref)}:${normalizeListRef(ref).toLowerCase()}`;
}

/** All distinct ID placeholders in a text, in order of first appearance. */
export function findPlaceholders(text) {
  const found = new Map();
  const hits = [];
  for (const { re, toRef } of ID_PATTERNS) {
    for (const m of String(text).matchAll(re)) hits.push({ at: m.index, token: m[0], ...toRef(m) });
  }
  hits.sort((a, b) => a.at - b.at);
  for (const h of hits) {
    const key = placeholderKey(h.type, h.ref);
    if (!found.has(key)) {
      found.set(key, { key, type: h.type, ref: h.type === 'web' ? '' : normalizeListRef(h.ref), tokens: [] });
    }
    const entry = found.get(key);
    if (!entry.tokens.includes(h.token)) entry.tokens.push(h.token);
  }
  return [...found.values()];
}

// ---------------------------------------------------------------- URL placeholders

/** Names of the URL placeholders used in a text (canonical spelling). */
export function findUrlPlaceholders(text) {
  return [...new Set([...String(text).matchAll(URL_RE)].map((m) => canonicalUrlName(m[1])))];
}

const pathOf = (url) => {
  const p = new URL(url).pathname.replace(/\/+$/, '');
  return p === '/' ? '' : p;
};

/** Values for the URL placeholders from the verified web and site collection URLs. */
export function urlPlaceholderValues({ webUrl, siteCollectionUrl }) {
  const web = String(webUrl).replace(/\/+$/, '');
  const site = String(siteCollectionUrl || webUrl).replace(/\/+$/, '');
  return { WebUrl: web, WebServerRelativeUrl: pathOf(web), SiteUrl: site, SiteServerRelativeUrl: pathOf(site) };
}

export function applyUrlPlaceholders(text, values) {
  return String(text).replace(URL_RE, (token, name) => values[canonicalUrlName(name)] ?? token);
}

// ---------------------------------------------------------------- resolution

// Same approach as the Power Automate flow: read all lists of the web once and
// build an object { "<list server-relative URL>": "<list id>" }.
export const LISTS_PATH = '/_api/web/lists?$select=Id,Title,RootFolder/ServerRelativeUrl&$expand=RootFolder&$top=5000';
export const WEB_PATH = '/_api/web?$select=Id,ServerRelativeUrl';

/** Build the flow's objListGuids object (plus a title index) from the lists response. */
export function buildListGuids(json) {
  const rows = (json && (json.value || (json.d && json.d.results))) || [];
  const byUrl = {};
  const byTitle = {};
  for (const l of rows) {
    if (!l || !l.Id) continue;
    const id = stripBraces(l.Id);
    const url = l.RootFolder && l.RootFolder.ServerRelativeUrl;
    if (url) byUrl[url] = id;
    if (l.Title) byTitle[l.Title.toLowerCase()] = id;
  }
  return { byUrl, byTitle };
}

function byUrl(guids, serverRelativeUrl) {
  const want = serverRelativeUrl.toLowerCase();
  const hit = Object.keys(guids.byUrl).find((u) => u.replace(/\/+$/, '').toLowerCase() === want);
  return hit ? guids.byUrl[hit] : null;
}

const joinWeb = (webPath, rel) => `${(webPath || '').replace(/\/+$/, '')}/${rel}`;

/** Find a list for a normalized reference. Returns { id } or { error }. */
export function findList(guids, ref, webServerRelativeUrl) {
  const kind = listRefKind(ref);
  if (kind === 'url') return byUrl(guids, ref) ? { id: byUrl(guids, ref) } : { error: 'List not found.' };
  if (kind === 'relative') {
    const id = byUrl(guids, joinWeb(webServerRelativeUrl, ref));
    return id ? { id } : { error: 'List not found.' };
  }
  const urlId = byUrl(guids, joinWeb(webServerRelativeUrl, ref));
  const titleId = guids.byTitle[ref.toLowerCase()] || null;
  if (urlId && titleId && urlId !== titleId) {
    return { error: `Ambiguous: one list has the title "${ref}", another has the URL ".../${ref}". Use the full URL, e.g. {{ListId:/sites/.../${ref}}}.` };
  }
  return urlId || titleId ? { id: urlId || titleId } : { error: 'List not found.' };
}

export function stripBraces(guid) {
  return String(guid).trim().replace(/^\{/, '').replace(/\}$/, '');
}

const httpMessage = (res) => `HTTP ${res.status}${res.message ? ': ' + res.message : ''}`;

/**
 * Resolve placeholders through `fetchJson(path)`, which resolves to
 * { status, json, message } (or throws). The web is read once (for {{WebId}}
 * and site-relative references) and the lists once.
 * Returns { resolved, listGuids }: entries with id or error filled in, and the
 * URL -> id object (null if the lists were not read).
 */
export async function resolvePlaceholders(placeholders, fetchJson) {
  const resolved = placeholders.map((p) => ({ ...p, id: null, error: null }));
  const lists = resolved.filter((p) => p.type === 'list');
  const webs = resolved.filter((p) => p.type === 'web');
  let listGuids = null;
  let web = null;

  if (webs.length || lists.some((p) => listRefKind(p.ref) !== 'url')) {
    try {
      const res = await fetchJson(WEB_PATH);
      const json = res.json && (res.json.d || res.json);
      if (res.status >= 200 && res.status < 300 && json && json.Id) web = json;
      else for (const p of webs) p.error = httpMessage(res);
    } catch (e) {
      for (const p of webs) p.error = (e && e.message) || String(e);
    }
    for (const p of webs) if (web) p.id = stripBraces(web.Id);
  }

  if (lists.length) {
    try {
      const res = await fetchJson(LISTS_PATH);
      if (res.status >= 200 && res.status < 300) {
        const guids = buildListGuids(res.json);
        listGuids = guids.byUrl;
        for (const p of lists) {
          const hit = findList(guids, p.ref, web && web.ServerRelativeUrl);
          p.id = hit.id || null;
          p.error = hit.error || null;
        }
      } else {
        for (const p of lists) p.error = `Reading the lists failed: ${httpMessage(res)}`;
      }
    } catch (e) {
      for (const p of lists) p.error = `Reading the lists failed: ${(e && e.message) || e}`;
    }
  }
  return { resolved, listGuids };
}

/** Replace every token of every resolved placeholder. Unresolved ones are left as is. */
export function applyPlaceholders(text, resolved) {
  let result = String(text);
  for (const p of resolved) {
    if (!p.id) continue;
    for (const token of p.tokens) result = result.split(token).join(p.id);
  }
  return result;
}

export function describePlaceholder(p) {
  return p.type === 'web' ? '{{WebId}}' : `List "${p.ref}"`;
}

/**
 * The operation in `operations` (before index `before`) that creates the list
 * a reference points to: POST .../_api/web/lists whose Title matches the
 * reference's title or last URL segment. Used to explain and plan splits.
 */
export function findCreator(operations, ref, before = operations.length) {
  const last = ref.replace(/\/+$/, '').split('/').pop().toLowerCase();
  for (let i = 0; i < before; i++) {
    const op = operations[i];
    if (op.method !== 'POST' || !/\/_api\/web\/lists\/?$/i.test(op.url || '')) continue;
    let title = null;
    try {
      title = JSON.parse(op.body).Title;
    } catch {
      const m = /"Title"\s*:\s*"([^"]*)"/.exec(op.body || '');
      title = m && m[1];
    }
    if (title && (title.toLowerCase() === ref.toLowerCase() || title.toLowerCase() === last)) return op;
  }
  return null;
}
