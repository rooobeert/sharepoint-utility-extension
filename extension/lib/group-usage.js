// "Where is a group used?": crawl the webs of a site collection and their lists
// and libraries, and report where a SharePoint group holds a role assignment.
// Scope: web and list/library level only. Items, folders, columns, sharing
// links, alerts and flows are not looked at.
//
// `get(webUrl, path)` returns parsed JSON (the permission client's GET with
// throttling retries). It is injected so the crawl can be tested without a tab.

const results = (json) => (json && (json.value || (json.d && json.d.results))) || [];

/** Permission level names a group holds in a role assignments response; [] if it has none. */
export function groupRoles(json, groupId) {
  const roles = [];
  for (const ra of results(json)) {
    if (ra.PrincipalId !== groupId) continue;
    for (const r of results({ value: ra.RoleDefinitionBindings })) roles.push({ name: r.Name, hidden: !!r.Hidden });
  }
  return roles;
}

/** True if every permission level is "Limited Access" (SharePoint adds it itself so a user can reach a parent). */
export const limitedOnly = (roles) => roles.length > 0 && roles.every((r) => r.hidden);

const RA_PATH = '/roleassignments?$select=PrincipalId&$expand=RoleDefinitionBindings&$top=5000';
const WEB_SELECT = 'Title,Url,ServerRelativeUrl,HasUniqueRoleAssignments';
const LIST_SELECT = 'Id,Title,Hidden,BaseType,BaseTemplate,HasUniqueRoleAssignments,IsSystemList,RootFolder/ServerRelativeUrl';

/** Run `fn` over `items` with at most `n` calls at a time, keeping the order of the results. */
async function pool(items, n, fn) {
  const out = new Array(items.length);
  let next = 0;
  const worker = async () => {
    while (next < items.length) {
      const i = next++;
      out[i] = await fn(items[i], i);
    }
  };
  await Promise.all(Array.from({ length: Math.min(n, items.length) }, worker));
  return out;
}

/**
 * Crawl `rootUrl` (the site collection's root web) and all its subwebs.
 *
 * Returns { webs, lists, places, errors, stopped }:
 *   places: [{ type: 'web' | 'list' | 'library', title, url, webTitle, webUrl, roles: [name], limited }]
 *   errors: [{ where, message }]  parts that could not be read; the crawl goes on
 *   webs, lists: how many were looked at
 * Only objects that define their own permissions can name the group; the ones
 * that inherit are not asked (their parent shows the group).
 *
 * @param {object} deps
 * @param {(webUrl: string, path: string) => Promise<object>} deps.get
 * @param {(event: object) => void} [deps.onEvent]  { type: 'web', url, webs } before each web
 * @param {() => boolean} [deps.shouldStop]
 */
export async function scanGroupUsage({ get, onEvent = () => {}, shouldStop = () => false }, rootUrl, groupId, { concurrency = 4 } = {}) {
  const out = { webs: 0, lists: 0, places: [], errors: [], stopped: false };
  const fail = (where, e) => out.errors.push({ where, message: (e && e.message) || String(e) });
  const queue = [rootUrl.replace(/\/+$/, '')];
  const seen = new Set();

  while (queue.length) {
    if (shouldStop()) {
      out.stopped = true;
      break;
    }
    const url = queue.shift();
    if (seen.has(url.toLowerCase())) continue;
    seen.add(url.toLowerCase());
    onEvent({ type: 'web', url, webs: out.webs, queued: queue.length });

    let web;
    try {
      web = await get(url, `/_api/web?$select=${WEB_SELECT}`);
    } catch (e) {
      fail(url, e);
      continue;
    }
    out.webs++;
    const here = { webTitle: web.Title || url, webUrl: url };

    if (web.HasUniqueRoleAssignments) {
      try {
        const roles = groupRoles(await get(url, `/_api/web${RA_PATH}`), groupId);
        if (roles.length) out.places.push({ type: 'web', title: here.webTitle, url, ...here, roles: roles.map((r) => r.name), limited: limitedOnly(roles) });
      } catch (e) {
        fail(`${here.webTitle} (web permissions)`, e);
      }
    }

    try {
      const subs = results(await get(url, `/_api/web/webs?$select=Url&$top=5000`));
      for (const s of subs) if (s.Url) queue.push(s.Url.replace(/\/+$/, ''));
    } catch (e) {
      fail(`${here.webTitle} (subsites)`, e);
    }

    let lists = [];
    try {
      lists = results(await get(url, `/_api/web/lists?$select=${LIST_SELECT}&$expand=RootFolder&$top=5000`));
    } catch (e) {
      fail(`${here.webTitle} (lists)`, e);
    }
    out.lists += lists.length;
    const unique = lists.filter((l) => l.HasUniqueRoleAssignments);
    const found = await pool(unique, concurrency, async (l) => {
      if (shouldStop()) return null;
      try {
        const roles = groupRoles(await get(url, `/_api/web/lists(guid'${l.Id}')${RA_PATH}`), groupId);
        if (!roles.length) return null;
        const lib = l.BaseType === 1;
        const path = (l.RootFolder && l.RootFolder.ServerRelativeUrl) || '';
        return { type: lib ? 'library' : 'list', title: l.Title, url: path ? new URL(path, url).href : url, ...here, roles: roles.map((r) => r.name), limited: limitedOnly(roles), hidden: !!l.Hidden || !!l.IsSystemList };
      } catch (e) {
        fail(`${here.webTitle} / ${l.Title}`, e);
        return null;
      }
    });
    out.places.push(...found.filter(Boolean));
  }
  return out;
}
