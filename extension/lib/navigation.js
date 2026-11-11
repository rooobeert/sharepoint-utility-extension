// "Navigation" builder: list grouping, the editable navigation tree, and the
// plan that turns the edited tree into REST calls. Pure functions: no DOM, no
// chrome.* APIs.
//
// Tree node: { key, id, title, url, children }. `id` is the SharePoint
// navigation node id (null for nodes added in the builder), `url` is '' for a
// heading without a link.

export const NAV_TARGETS = {
  quicklaunch: { label: 'Quick Launch', path: '/_api/web/navigation/QuickLaunch' },
  topnav: { label: 'Top navigation', path: '/_api/web/navigation/TopNavigationBar' },
  // The hub navigation is the top navigation of the hub site's own web, so it is read and
  // written there (same path, on the hub site's URL).
  hub: { label: 'Hub navigation', path: '/_api/web/navigation/TopNavigationBar' },
};

/**
 * Hub site of a web, from /_api/web/HubSiteData (with nometadata the answer is { value: "<json>" },
 * for a site without hub null/empty). Returns { url, name } or null.
 */
export function parseHubSite(json) {
  let data = json;
  try {
    if (data && typeof data.value === 'string') data = JSON.parse(data.value);
    else if (typeof data === 'string') data = JSON.parse(data);
  } catch {
    return null;
  }
  const url = data && typeof data.url === 'string' ? data.url.trim().replace(/\/+$/, '') : '';
  return /^https:\/\//i.test(url) ? { url, name: String(data.name || '') } : null;
}

const NO_GUID = /^\{?0{8}-0{4}-0{4}-0{4}-0{12}\}?$/;
const bareGuid = (g) => String(g || '').replace(/[{}]/g, '').toLowerCase();

/** The hub a site collection belongs to, from /_api/site (HubSiteId, IsHubSite, Url) and /_api/hubsites. Returns { url, name } or null. */
export function hubFromSite(site, hubSites) {
  if (site && site.IsHubSite && site.Url) return { url: String(site.Url).replace(/\/+$/, ''), name: '' };
  const id = bareGuid(site && site.HubSiteId);
  if (!id || NO_GUID.test(id)) return null;
  const list = (hubSites && (hubSites.value || (hubSites.d && hubSites.d.results))) || [];
  const hit = list.find((h) => bareGuid(h.SiteId || h.ID || h.Id) === id);
  const url = hit && (hit.SiteUrl || hit.Url);
  return url ? { url: String(url).replace(/\/+$/, ''), name: String(hit.Title || '') } : null;
}

/** SharePoint's URL for a heading that is only a label (what the modern UI uses). */
export const LINKLESS = 'http://linkless.header/';

/** SharePoint shows at most three levels in both navigations. */
export const MAX_DEPTH = 3;

/**
 * Which navigation the site shows, from its template and navigation settings
 * (/_api/web: WebTemplate, QuickLaunchEnabled, HorizontalQuickLaunch, MegaMenuEnabled).
 * Communication sites show the top navigation. Team sites show the Quick
 * Launch, on the left or, with the horizontal layout, at the top.
 * Returns { target, layout, text }.
 */
export function detectNavigation(web) {
  const w = web || {};
  if (/^SITEPAGEPUBLISHING/i.test(w.WebTemplate || '')) {
    return { target: 'topnav', layout: 'top', text: `Communication site: the navigation at the top (${w.MegaMenuEnabled ? 'mega menu' : 'cascading'}) is the top navigation.` };
  }
  if (w.QuickLaunchEnabled === false) {
    return { target: 'quicklaunch', layout: 'off', text: 'Team site with the site navigation turned off (Change the look > Navigation). Changes are saved but not shown until it is turned on again.' };
  }
  if (w.HorizontalQuickLaunch) {
    return { target: 'quicklaunch', layout: 'horizontal', text: 'Team site with horizontal navigation: the bar at the top is stored as the Quick Launch.' };
  }
  return { target: 'quicklaunch', layout: 'vertical', text: 'Team site: the navigation on the left is the Quick Launch.' };
}

export const defaultTarget = (web) => detectNavigation(web).target;

const urlName = (serverRelativeUrl) => String(serverRelativeUrl || '').split('/').pop();

/**
 * Solution tag at the start of a list's internal (URL) name: the lowercase
 * letters before the first capital, digit or underscore, e.g. "hrpLocations"
 * -> "hrp", "tdasTermDeals" -> "tdas". "Documents" or "SitePages" -> null.
 */
export function listPrefix(serverRelativeUrl) {
  const m = /^([a-z]{2,8})(?=[A-Z0-9_])/.exec(urlName(serverRelativeUrl));
  return m ? m[1] : null;
}

/**
 * Non-system lists grouped by solution tag. A tag needs at least `minLists`
 * lists; the rest go into a last group with prefix null.
 * Returns [{ prefix, label, lists }].
 */
export function groupListsByPrefix(lists, { minLists = 2 } = {}) {
  const byTitle = (a, b) => a.title.localeCompare(b.title, undefined, { sensitivity: 'base' });
  const groups = new Map();
  const other = [];
  for (const l of lists.filter((x) => !x.system)) {
    const p = listPrefix(l.url);
    if (!p) other.push(l);
    else groups.set(p, [...(groups.get(p) || []), l]);
  }
  const out = [];
  for (const [prefix, members] of groups) {
    if (members.length >= minLists) out.push({ prefix, label: prefix.toUpperCase(), lists: members.sort(byTitle) });
    else other.push(...members);
  }
  out.sort((a, b) => a.label.localeCompare(b.label));
  if (other.length) out.push({ prefix: null, label: 'Other lists and libraries', lists: other.sort(byTitle) });
  return out;
}

// ---------------------------------------------------------------- URLs

const isHeadingUrl = (url) => !url || /^https?:\/\/linkless\.header\/?$/i.test(url);

/** '' for a heading, the URL otherwise (trimmed). */
export const editorUrl = (url) => (isHeadingUrl(String(url || '').trim()) ? '' : String(url).trim());

/** What is sent to SharePoint: LINKLESS for a heading. */
export const sharePointUrl = (url) => editorUrl(url) || LINKLESS;

/** Absolute URLs (and headings) must be created with IsExternal = true. */
export const isExternalUrl = (url) => /^https?:\/\//i.test(sharePointUrl(url));

/** Lower-case, decoded, server-relative path of a link on `origin` (null for other hosts). */
export function linkPath(url, origin) {
  let u = editorUrl(url);
  if (!u) return null;
  if (/^https?:\/\//i.test(u)) {
    if (!origin || !u.toLowerCase().startsWith(origin.toLowerCase() + '/')) return null;
    u = u.slice(origin.length);
  }
  u = u.split(/[?#]/)[0];
  try {
    u = decodeURIComponent(u);
  } catch {
    /* keep as is */
  }
  return u.replace(/\/+$/, '').toLowerCase();
}

/** True if the link opens the list (its root folder or a page below it, e.g. a view). */
export function linksToList(url, listUrl, origin) {
  const path = linkPath(url, origin);
  const list = String(listUrl || '').replace(/\/+$/, '').toLowerCase();
  return !!path && !!list && (path === list || path.startsWith(list + '/'));
}

// ---------------------------------------------------------------- tree

let nextKey = 1;
export const newKey = () => `new${nextKey++}`;

/** Editor tree from REST nodes: [{ Id, Title, Url, Children: [...] }]. */
export function treeFromNodes(nodes) {
  return (nodes || []).map((n) => ({
    key: `id${n.Id}`,
    id: n.Id,
    title: n.Title || '',
    url: editorUrl(n.Url),
    children: treeFromNodes(n.Children),
  }));
}

export const cloneTree = (tree) => tree.map((n) => ({ ...n, children: cloneTree(n.children) }));

/** Comparable form of a tree: same text means nothing to save. */
export const treeSignature = (tree) =>
  JSON.stringify(tree.map(function sig(n) {
    return [n.id, n.title.trim(), editorUrl(n.url), n.children.map(sig)];
  }));

export function walkTree(tree, fn, depth = 1, parent = null) {
  for (const n of tree) {
    fn(n, depth, parent);
    walkTree(n.children, fn, depth + 1, n);
  }
}

/** Levels below and including the node (a leaf is 1). */
export const subtreeHeight = (node) => 1 + Math.max(0, ...node.children.map(subtreeHeight));

/** { node, parent, siblings, index, depth } for a key, or null. */
export function findNode(tree, key, depth = 1, parent = null) {
  for (const [index, node] of tree.entries()) {
    if (node.key === key) return { node, parent, siblings: tree, index, depth };
    const hit = findNode(node.children, key, depth + 1, node);
    if (hit) return hit;
  }
  return null;
}

/** Problems that block saving: empty titles, bad URLs, too many levels. */
export function validateTree(tree) {
  const problems = [];
  walkTree(tree, (n, depth, parent) => {
    const where = parent ? `under "${parent.title.trim() || '(no title)'}"` : 'at the top level';
    if (!n.title.trim()) problems.push(`A ${n.url ? 'link' : 'heading'} ${where} has no title.`);
    const url = editorUrl(n.url);
    if (url && !/^(https?:\/\/|\/)/i.test(url)) {
      problems.push(`"${n.title.trim()}": the URL must start with "/" (on this tenant) or "https://" (or "http://").`);
    }
    if (depth > MAX_DEPTH) problems.push(`"${n.title.trim()}" is on level ${depth}. SharePoint shows at most ${MAX_DEPTH} levels.`);
  });
  return problems;
}

// ---------------------------------------------------------------- save plan

/**
 * What to send so that SharePoint ends up with `desired`.
 *
 * SharePoint can only add a node at the end of its parent's children and
 * can't move a node to another parent. So per parent, existing nodes are kept
 * while they are still in their original order; from the first new or moved
 * node on, the rest of that level is created at the end (moved nodes are
 * deleted and recreated, with their children).
 *
 * Returns {
 *   deletes: [{ id, title }]               top-most removed nodes (children go with them)
 *   updates: [{ id, title, url, from }]    kept nodes with a new title or URL
 *   creates: [{ key, parent, title, url, external, recreated }]  in order;
 *            parent is { id } (existing), { key } (created earlier) or null (top level)
 *   removed: [{ title }]                   every node that is gone afterwards and not recreated
 * }
 */
export function planSave(original, desired) {
  const orig = new Map();
  walkTree(original, (n, depth, parent) => orig.set(n.id, { node: n, parentId: parent ? parent.id : null }));
  const desiredIds = new Set();
  walkTree(desired, (n) => n.id != null && desiredIds.add(n.id));

  const deletes = [];
  const updates = [];
  const creates = [];

  const create = (n, parent) => {
    const url = editorUrl(n.url);
    creates.push({ key: n.key, parent, title: n.title.trim(), url, external: isExternalUrl(url), recreated: n.id != null });
    for (const c of n.children) create(c, { key: n.key });
  };

  const level = (parentId, origChildren, wanted, parentRef) => {
    const kept = new Set();
    let last = -1;
    let append = false;
    for (const n of wanted) {
      const o = n.id != null ? orig.get(n.id) : null;
      const at = o && o.parentId === parentId ? origChildren.findIndex((c) => c.id === n.id) : -1;
      if (!append && at > last) {
        kept.add(n.id);
        last = at;
        const title = n.title.trim();
        const url = editorUrl(n.url);
        if (title !== o.node.title.trim() || url !== editorUrl(o.node.url)) {
          updates.push({ id: n.id, title, url, from: { title: o.node.title, url: editorUrl(o.node.url) } });
        }
        level(n.id, o.node.children, n.children, { id: n.id });
      } else {
        append = true;
        create(n, parentRef);
      }
    }
    for (const c of origChildren) if (!kept.has(c.id)) deletes.push({ id: c.id, title: c.title });
  };
  level(null, original, desired, null);

  const removed = [];
  const gone = (n) => {
    if (!desiredIds.has(n.id)) removed.push({ title: n.title });
    n.children.forEach(gone);
  };
  for (const d of deletes) gone(orig.get(d.id).node);

  return { deletes, updates, creates, removed };
}

export const planIsEmpty = (plan) => !plan.deletes.length && !plan.updates.length && !plan.creates.length;

// ---------------------------------------------------------------- copy and paste between sites

export const EXPORT_FORMAT = 'sp-utility-navigation';
const MAX_IMPORT_NODES = 1000;
const WEB_REL = '{{WebServerRelativeUrl}}';

/**
 * A link that works on any site: a link into this web becomes {{WebServerRelativeUrl}}/..., so
 * it follows the site it is pasted on. Other links (other sites, other hosts, headings) stay.
 */
export function portableUrl(url, { origin, webServerRelativeUrl }) {
  const u = editorUrl(url);
  if (!u) return '';
  const web = String(webServerRelativeUrl || '').replace(/\/+$/, '');
  let path = u;
  if (/^https?:\/\//i.test(u)) {
    if (!origin || !u.toLowerCase().startsWith(origin.toLowerCase() + '/')) return u;
    path = u.slice(origin.length);
  }
  if (!path.startsWith('/')) return u;
  const lower = path.toLowerCase();
  if (!web) return WEB_REL + path; // root web: every path on the host is inside it
  if (lower === web.toLowerCase() || lower.startsWith(web.toLowerCase() + '/') || lower.startsWith(web.toLowerCase() + '?')) return WEB_REL + path.slice(web.length);
  return u;
}

/** The pasted link for the target web (placeholders resolved). */
export function resolvePortableUrl(url, { site, webServerRelativeUrl }) {
  const web = String(webServerRelativeUrl || '').replace(/\/+$/, '');
  const out = String(url || '')
    .replace(/\{\{\s*WebServerRelativeUrl\s*\}\}/gi, web)
    .replace(/\{\{\s*WebUrl\s*\}\}/gi, String(site || '').replace(/\/+$/, ''));
  return out === '' && /WebServerRelativeUrl/i.test(String(url || '')) ? '/' : out;
}

/** Text to copy: the tree as JSON with portable links. */
export function exportNavigation(tree, { target, site, origin, webServerRelativeUrl }) {
  const nodes = (list) => list.map((n) => {
    const url = portableUrl(n.url, { origin, webServerRelativeUrl });
    return { title: n.title, ...(url ? { url } : {}), ...(n.children.length ? { children: nodes(n.children) } : {}) };
  });
  return JSON.stringify({ format: EXPORT_FORMAT, version: 1, target, source: site, nodes: nodes(tree) }, null, 2);
}

/**
 * Pasted text -> { nodes, target, source } with editor nodes (new keys, links for the target web).
 * Accepts the exported object or a bare array of { title, url?, children? }. Throws Error with a readable message.
 */
export function parseNavigationExport(text, { site, webServerRelativeUrl }) {
  let data;
  try {
    data = JSON.parse(String(text).trim());
  } catch {
    throw new Error('The text is not a copied navigation (it is not valid JSON).');
  }
  if (data && !Array.isArray(data) && data.format !== EXPORT_FORMAT) throw new Error('The text is not a navigation copied from this extension.');
  const list = Array.isArray(data) ? data : data && data.nodes;
  if (!Array.isArray(list)) throw new Error('The copied navigation has no "nodes" list.');
  let count = 0;
  const build = (items, path) => items.map((n, i) => {
    const where = `${path}${i + 1}`;
    if (!n || typeof n !== 'object' || typeof n.title !== 'string') throw new Error(`Item ${where} has no title.`);
    if (n.url != null && typeof n.url !== 'string') throw new Error(`Item ${where} ("${n.title}") has a link that is not text.`);
    if (n.children != null && !Array.isArray(n.children)) throw new Error(`Item ${where} ("${n.title}") has children that are not a list.`);
    if (++count > MAX_IMPORT_NODES) throw new Error(`More than ${MAX_IMPORT_NODES} items.`);
    return {
      key: newKey(),
      id: null,
      title: n.title,
      url: editorUrl(resolvePortableUrl(n.url || '', { site, webServerRelativeUrl })),
      children: build(n.children || [], `${where}.`),
    };
  });
  const nodes = build(list, '');
  return {
    nodes,
    count,
    target: !Array.isArray(data) && NAV_TARGETS[data.target] ? data.target : null,
    source: !Array.isArray(data) && typeof data.source === 'string' ? data.source : '',
  };
}

// ---------------------------------------------------------------- backups of the previous navigation

export const MAX_BACKUPS_PER_NAV = 5;
const MAX_BACKUP_KEYS = 30;

/** Backups are kept per site and navigation. */
export const backupKey = (site, target) => `${String(site).replace(/\/+$/, '').toLowerCase()}|${target}`;

/** Titles, links and nesting only: what is needed to rebuild the tree. */
export const snapshotTree = (tree) => tree.map((n) => ({ title: n.title, url: editorUrl(n.url), children: snapshotTree(n.children) }));

export const countNodes = (tree) => tree.reduce((n, x) => n + 1 + countNodes(x.children || []), 0);

/** New list of backups for `key` with `entry` first (newest), at most MAX_BACKUPS_PER_NAV; other keys beyond MAX_BACKUP_KEYS drop the oldest. */
export function addBackup(all, key, entry) {
  const next = { ...all, [key]: [entry, ...(all[key] || [])].slice(0, MAX_BACKUPS_PER_NAV) };
  const keys = Object.keys(next);
  if (keys.length > MAX_BACKUP_KEYS) {
    const newest = (k) => next[k][0]?.at || '';
    for (const k of keys.filter((x) => x !== key).sort((a, b) => newest(a).localeCompare(newest(b))).slice(0, keys.length - MAX_BACKUP_KEYS)) delete next[k];
  }
  return next;
}

/** Editor tree (new nodes) from a stored snapshot. */
export const treeFromSnapshot = (snap) => snap.map((n) => ({ key: newKey(), id: null, title: n.title, url: n.url || '', children: treeFromSnapshot(n.children || []) }));
