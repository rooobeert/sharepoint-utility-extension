// "Permissions": set permissions for several principals on several lists.
// Pure functions: no DOM, no chrome.* APIs.
//
// Principal: { key, id, login, title, kind, detail, email }. `kind` is the label
// shown in the UI: 'SPO Group', 'Entra Group', 'User' or 'Everyone'; `detail`
// tells Entra group types apart. SharePoint groups carry
// their `id`; users and Entra groups carry their `login` (claims login name)
// and get an id from web/ensureuser only when the plan is applied, so Preview
// stays read-only.
//
// Plan entry: { listId, principal, roleId }.
// Item-level setting per list: { read: 1 | 2 | null, write: 1 | 2 | 4 | null }
// (SP.List ReadSecurity / WriteSecurity; null = don't change).

import { CRLF, MAX_OPERATIONS_PER_BATCH } from './batch.js';

const results = (json) => (json && (json.value || (json.d && json.d.results))) || [];
const items = (v) => (Array.isArray(v) ? v : (v && (v.Items || v.results)) || []);

/** RoleTypeKind of "Full Control" (Administrator). */
export const FULL_CONTROL = 5;
/** RoleTypeKind of "Read" (Reader): the default level for a new principal. */
export const READER = 2;

/** The list settings page's "Item-level Permissions" options. */
export const READ_SECURITY = { 1: 'Read all items', 2: 'Read items that were created by the user' };
export const WRITE_SECURITY = { 1: 'Create and edit all items', 2: 'Create items and edit items that were created by the user', 4: 'None' };

/** Libraries have no item-level permission settings. */
export function isLibrary(list) {
  return list.baseType === 1 || [101, 109, 119, 851].includes(list.template);
}

/** Short text of an item-level setting, e.g. "Read: own items · Create and edit: none". */
export function itemLevelText({ read, write }) {
  const r = { 1: 'all items', 2: 'own items' };
  const w = { 1: 'all items', 2: 'create, edit own', 4: 'none' };
  return [read ? `Read: ${r[read]}` : null, write ? `Create and edit: ${w[write]}` : null].filter(Boolean).join(' · ');
}

// ---------------------------------------------------------------- principals

const lower = (s) => String(s || '').toLowerCase();

export function principalKey(p) {
  return p.id != null && !p.login ? `g:${p.id}` : `l:${lower(p.login)}`;
}

function principal(props) {
  const p = { id: null, login: null, email: null, ...props };
  return { ...p, key: principalKey(p) };
}

/** A SharePoint group from web/sitegroups. */
export function groupPrincipal(g) {
  return principal({ id: g.Id, title: g.Title, kind: SPO_GROUP });
}

export const SPO_GROUP = 'SPO Group';
export const ENTRA_GROUP = 'Entra Group';

/** Label and group type of a claims login name: { kind, detail }. */
export function loginKind(login, entityType) {
  const l = lower(login);
  if (/federateddirectoryclaimprovider/.test(l)) return { kind: ENTRA_GROUP, detail: 'Microsoft 365 group' };
  if (/spo-grid-all-users|^c:0\(\.s\|true$/.test(l)) return { kind: 'Everyone', detail: null };
  if (entityType === 'SecGroup' || l.startsWith('c:0t.c|')) return { kind: ENTRA_GROUP, detail: 'Security group' };
  return { kind: 'User', detail: null };
}

/**
 * Results of clientPeoplePickerSearchUser. The endpoint returns the entities
 * as a JSON string inside the OData response. SharePoint groups of the site
 * are left out: the search is for tenant users and groups only.
 */
export function pickerPrincipals(json) {
  let raw = json && (json.d ? json.d.ClientPeoplePickerSearchUser : json.value);
  let list;
  try {
    list = typeof raw === 'string' ? JSON.parse(raw) : raw;
  } catch {
    list = null;
  }
  if (!Array.isArray(list)) return [];
  const out = [];
  for (const e of list) {
    const data = e.EntityData || {};
    if (e.EntityType === 'SPGroup') continue;
    if (e.Key) {
      out.push(principal({
        login: e.Key,
        title: e.DisplayText || e.Key,
        ...loginKind(e.Key, e.EntityType),
        email: data.Email || null,
      }));
    }
  }
  return out;
}

/** Request body for clientPeoplePickerSearchUser (users and all group kinds). */
export function pickerQuery(text, max = 50) {
  return JSON.stringify({
    queryParams: {
      __metadata: { type: 'SP.UI.ApplicationPages.ClientPeoplePickerQueryParameters' },
      QueryString: text,
      MaximumEntitySuggestions: max,
      AllowEmailAddresses: false,
      AllowMultipleEntities: false,
      AllUrlZones: false,
      PrincipalSource: 15,
      PrincipalType: 7, // users, distribution lists, security and Microsoft 365 groups; no SharePoint groups
    },
  });
}

// ---------------------------------------------------------------- site data

/** Permission levels a user can grant: not hidden (no "Limited Access"), in SharePoint's order. */
export function roleLevels(json) {
  return results(json)
    .filter((r) => !r.Hidden)
    .map((r) => ({ id: r.Id, name: r.Name, kind: r.RoleTypeKind, description: r.Description || '', order: r.Order ?? 0 }))
    .sort((a, b) => a.order - b.order || a.name.localeCompare(b.name));
}

/** The role picked for a new principal: Read if it exists, else the last level. */
export function defaultRoleId(roles) {
  const read = roles.find((r) => r.kind === READER);
  return (read || roles[roles.length - 1] || {}).id ?? null;
}

/**
 * Site groups as principals, the default groups first (Owners, Members,
 * Visitors) and flagged `isDefault`, then the others by title.
 */
export function siteGroups(groupsJson, web = {}) {
  const defaults = [web.AssociatedOwnerGroup, web.AssociatedMemberGroup, web.AssociatedVisitorGroup]
    .map((g) => g && g.Id)
    .filter((id) => id != null);
  const all = results(groupsJson).map((g) => ({ ...groupPrincipal(g), isDefault: defaults.includes(g.Id) }));
  const rank = (p) => (p.isDefault ? defaults.indexOf(p.id) : defaults.length);
  return all.sort((a, b) => rank(a) - rank(b) || a.title.localeCompare(b.title, undefined, { sensitivity: 'base' }));
}

/** Ids of the default groups: { ids, owners, members, visitors } (null if the site has none). */
export function defaultGroupIds(web = {}) {
  const id = (g) => (g && g.Id) ?? null;
  const owners = id(web.AssociatedOwnerGroup);
  const members = id(web.AssociatedMemberGroup);
  const visitors = id(web.AssociatedVisitorGroup);
  return { ids: [owners, members, visitors].filter((x) => x != null), owners, members, visitors };
}

/** Role assignments of a list: [{ principalId, title, login, kind, roles: [{ id, name, kind, hidden }] }]. */
export function assignments(json) {
  return results(json).map((ra) => {
    const m = ra.Member || {};
    return {
      principalId: m.Id ?? ra.PrincipalId,
      title: m.Title || m.LoginName || String(ra.PrincipalId),
      login: m.LoginName || null,
      group: m.PrincipalType === 8,
      roles: items(ra.RoleDefinitionBindings).map((r) => ({ id: r.Id, name: r.Name, kind: r.RoleTypeKind, hidden: !!r.Hidden })),
    };
  });
}

// ---------------------------------------------------------------- plan

export const entryKey = (e) => `${e.listId}|${e.principal.key}|${e.roleId}`;

/** Add the cross product of lists x principals (each with its role); duplicates are skipped. Returns the number added. */
export function addEntries(plan, listIds, picks) {
  let added = 0;
  for (const listId of listIds) {
    for (const { principal: p, roleId } of picks) {
      const e = { listId, principal: p, roleId };
      const key = entryKey(e);
      if (plan.has(key)) continue;
      plan.set(key, e);
      added++;
    }
  }
  return added;
}

/** Does a role assignment belong to this principal? Groups by id, others by login name. */
export function matches(p, a) {
  if (p.id != null && a.principalId === p.id) return true;
  return !!p.login && !!a.login && lower(p.login) === lower(a.login);
}

/**
 * What to do on one list.
 *   list:      { id, title }
 *   current:   { unique, assignments, readSecurity, writeSecurity } as read by Preview
 *   entries:   plan entries of this list (none: its permissions are left alone,
 *              also in replace mode)
 *   options:   { mode: 'add' | 'replace', keepGroups: ids of default groups kept in replace mode }
 *   site:      { roles, defaults: { ids, owners } }
 *   itemLevel: { read, write } or null
 *
 * Returns { list, breaks, actions, warnings }. Actions in order:
 *   { type: 'add' | 'remove', principal|assignment, role, title }
 *   { type: 'keep' | 'skip', title, roles, why }  (informational, nothing is sent)
 * `breaks` is true when the list inherits and a permission changes: inheritance
 * is broken with a copy of the current assignments first. `settings` are the
 * item-level changes: [{ field: 'ReadSecurity' | 'WriteSecurity', label, from, to, change }].
 */
export function planList(list, current, entries, options, site, itemLevel = null) {
  const settings = itemLevelSettings(current, itemLevel);
  if (!entries.length) return { list, breaks: false, actions: [], warnings: [], settings };
  const roleById = new Map(site.roles.map((r) => [r.id, r]));
  const replace = options.mode === 'replace';
  const actions = [];
  const warnings = [];
  const roleName = (id) => (roleById.get(id) || { name: `role ${id}` }).name;

  // Desired roles per principal, in entry order.
  const desired = new Map();
  for (const e of entries) {
    if (!desired.has(e.principal.key)) desired.set(e.principal.key, { principal: e.principal, roleIds: [] });
    const d = desired.get(e.principal.key);
    if (!d.roleIds.includes(e.roleId)) d.roleIds.push(e.roleId);
  }

  const handled = new Set(); // desired principals matched to an existing assignment
  const finalRoles = []; // [{ title, kinds }] what the list ends up with, for the Full Control check

  for (const a of current.assignments) {
    const visible = a.roles.filter((r) => !r.hidden);
    const want = [...desired.values()].find((d) => matches(d.principal, a));
    const has = (id) => visible.some((r) => r.id === id);
    if (want) {
      handled.add(want.principal.key);
      const stays = replace ? visible.filter((r) => want.roleIds.includes(r.id)) : visible;
      if (replace) {
        for (const r of visible) {
          if (!want.roleIds.includes(r.id)) actions.push({ type: 'remove', assignment: a, role: r, title: a.title });
        }
      }
      for (const id of want.roleIds) {
        if (has(id)) actions.push({ type: 'skip', title: a.title, roles: [roleName(id)], why: 'already has this level' });
        else actions.push({ type: 'add', principal: { ...want.principal, id: want.principal.id ?? a.principalId }, role: roleById.get(id) || { id, name: roleName(id) }, title: want.principal.title });
      }
      finalRoles.push({ title: a.title, kinds: [...stays.map((r) => r.kind), ...want.roleIds.map((id) => (roleById.get(id) || {}).kind)] });
      continue;
    }
    if (!visible.length) {
      // Only "Limited Access": SharePoint manages it (sharing links, item permissions).
      if (replace) actions.push({ type: 'keep', title: a.title, roles: a.roles.map((r) => r.name), why: 'Limited Access only, managed by SharePoint' });
      continue;
    }
    const kept = site.defaults.ids.includes(a.principalId) && (options.keepGroups || []).includes(a.principalId);
    if (!replace || kept) {
      if (replace) actions.push({ type: 'keep', title: a.title, roles: visible.map((r) => r.name), why: 'default group' });
      finalRoles.push({ title: a.title, kinds: visible.map((r) => r.kind) });
      continue;
    }
    for (const r of visible) actions.push({ type: 'remove', assignment: a, role: r, title: a.title });
    if (a.principalId === site.defaults.owners) {
      warnings.push(`The site's Owners group "${a.title}" loses its access to this list.`);
    }
  }

  for (const d of desired.values()) {
    if (handled.has(d.principal.key)) continue;
    for (const id of d.roleIds) {
      actions.push({ type: 'add', principal: d.principal, role: roleById.get(id) || { id, name: roleName(id) }, title: d.principal.title });
    }
    finalRoles.push({ title: d.principal.title, kinds: d.roleIds.map((id) => (roleById.get(id) || {}).kind) });
  }

  const changes = actions.some((x) => x.type === 'add' || x.type === 'remove');
  if (replace && changes && !finalRoles.some((f) => f.kinds.includes(FULL_CONTROL))) {
    warnings.push('Afterwards nobody has Full Control on this list. Only site collection administrators can then change its permissions.');
  }
  return { list, breaks: changes && !current.unique, actions, warnings, settings };
}

function itemLevelSettings(current, itemLevel) {
  if (!itemLevel) return [];
  const out = [];
  const add = (field, label, names, from, to) => {
    if (to) out.push({ field, label, from: names[from] || (from == null ? 'unknown' : String(from)), to: names[to], value: to, change: from !== to });
  };
  add('ReadSecurity', 'Read access', READ_SECURITY, current.readSecurity, itemLevel.read);
  add('WriteSecurity', 'Create and edit access', WRITE_SECURITY, current.writeSecurity, itemLevel.write);
  return out;
}

/** Counts for a list plan: { add, remove, keep, skip, set }. */
export function countActions(plan) {
  const c = { add: 0, remove: 0, keep: 0, skip: 0, set: 0 };
  for (const a of plan.actions) c[a.type]++;
  c.set = (plan.settings || []).filter((s) => s.change).length;
  return c;
}

/** Does applying this list plan send anything? */
export function hasChanges(plan) {
  const c = countActions(plan);
  return c.add + c.remove + c.set > 0;
}

// ---------------------------------------------------------------- batches

/**
 * REST operations of a list plan. Add operations need `principal.id`, which
 * `idOf(principal)` supplies for principals resolved with ensureuser.
 * Returns [{ method, url, label, listTitle }].
 */
export function listOperations(site, plan, idOf = (p) => p.id) {
  const base = `${site}/_api/web/lists(guid'${plan.list.id}')`;
  const ops = [];
  const listTitle = plan.list.title;
  const set = (plan.settings || []).filter((s) => s.change);
  if (set.length) {
    const body = { __metadata: { type: 'SP.List' } };
    for (const s of set) body[s.field] = s.value;
    ops.push({
      method: 'MERGE',
      url: base,
      headers: { 'Content-Type': 'application/json;odata=verbose', 'IF-MATCH': '*' },
      body: JSON.stringify(body),
      label: `Item-level permissions: ${set.map((s) => `${s.label} "${s.to}"`).join(', ')}`,
      listTitle,
    });
  }
  if (plan.breaks) {
    ops.push({ method: 'POST', url: `${base}/breakroleinheritance(copyRoleAssignments=true,clearSubscopes=false)`, label: 'Break inheritance (copy current permissions)', listTitle });
  }
  for (const a of plan.actions) {
    if (a.type === 'remove') {
      ops.push({ method: 'POST', url: `${base}/roleassignments/removeroleassignment(principalid=${a.assignment.principalId},roledefid=${a.role.id})`, label: `Remove ${a.role.name}: ${a.title}`, listTitle });
    } else if (a.type === 'add') {
      const id = idOf(a.principal);
      if (id == null) throw new Error(`No principal id for "${a.title}".`);
      ops.push({ method: 'POST', url: `${base}/roleassignments/addroleassignment(principalid=${id},roledefid=${a.role.id})`, label: `Add ${a.role.name}: ${a.title}`, listTitle });
    }
  }
  return ops;
}

/**
 * Pack operations into batches of at most `max`. A list's operations stay in
 * one batch unless the list alone has more than `max`.
 */
export function packBatches(opLists, max = MAX_OPERATIONS_PER_BATCH) {
  const batches = [];
  let current = [];
  for (const ops of opLists) {
    if (!ops.length) continue;
    if (current.length && current.length + ops.length > max) {
      batches.push(current);
      current = [];
    }
    for (const op of ops) {
      if (current.length >= max) {
        batches.push(current);
        current = [];
      }
      current.push(op);
    }
  }
  if (current.length) batches.push(current);
  return batches;
}

/** A $batch body with all operations in one changeset. */
export function buildPermissionBatch(ops, id = 'perm') {
  const boundary = `batch_${id}`;
  const changeset = `changeset_${id}`;
  let out = `--${boundary}${CRLF}Content-Type: multipart/mixed; boundary=${changeset}${CRLF}${CRLF}`;
  for (const op of ops) {
    out += `--${changeset}${CRLF}Content-Type: application/http${CRLF}Content-Transfer-Encoding: binary${CRLF}${CRLF}`;
    const headers = { ...(op.headers || {}), Accept: 'application/json;odata=nometadata' };
    out += `${op.method} ${op.url} HTTP/1.1${CRLF}`;
    for (const [k, v] of Object.entries(headers)) out += `${k}: ${v}${CRLF}`;
    out += `${CRLF}${op.body || ''}${CRLF}`;
  }
  out += `--${changeset}--${CRLF}--${boundary}--${CRLF}`;
  return out;
}
