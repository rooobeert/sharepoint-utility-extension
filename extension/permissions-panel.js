// "Permissions" tab: give several users and groups permission levels on
// several lists at once. Plan, Preview (reads only), then Apply in $batch
// requests.

import { el, activeTab, tabRequester, ask, siteOfTab } from './shared.js';
import { isSharePointUrl } from './lib/batch.js';
import { createPermissionClient } from './lib/permission-client.js';
import { SPO_GROUP, ENTRA_GROUP, addEntries, defaultRoleId, planList, countActions, hasChanges, isLibrary, itemLevelText } from './lib/permissions.js';

const $ = (id) => document.getElementById(id);

const ui = {
  site: $('perm-site'),
  reload: $('perm-reload'),
  reloadRoles: $('perm-reload-roles'),
  listCount: $('perm-list-count'),
  filter: $('perm-filter'),
  system: $('perm-system'),
  toggle: $('perm-toggle'),
  rows: $('perm-rows'),
  search: $('perm-search'),
  results: $('perm-results'),
  picked: $('perm-picked'),
  read: $('perm-read'),
  write: $('perm-write'),
  applyAll: $('perm-apply-all'),
  add: $('perm-add'),
  addHint: $('perm-add-hint'),
  clear: $('perm-clear'),
  plan: $('perm-plan'),
  planEmpty: $('perm-plan-empty'),
  modes: [...document.querySelectorAll('input[name="perm-mode"]')],
  keep: [...document.querySelectorAll('#perm-keep input')],
  preview: $('perm-preview'),
  apply: $('perm-apply'),
  applyHint: $('perm-apply-hint'),
  progress: $('perm-progress'),
  bar: $('perm-bar'),
  status: $('perm-status'),
  output: $('perm-output'),
  outputHead: $('perm-output-h'),
  outputBody: $('perm-output-body'),
  copy: $('perm-copy'),
  usageGroup: $('usage-group'),
  usageWith: $('usage-with'),
  usageActions: $('usage-actions'),
  usageRemove: $('usage-remove'),
  usageReplace: $('usage-replace'),
  usagePlanHint: $('usage-plan-hint'),
  usageRun: $('usage-run'),
  usageStop: $('usage-stop'),
  usageStatus: $('usage-status'),
  usageResult: $('usage-result'),
};

const perm = {
  busy: false,
  loaded: false,
  site: null, // web URL
  client: null,
  data: null, // { web, lists, roles, groups, defaults }
  selected: new Set(), // list ids for the next "Add to plan"
  picked: new Map(), // principal key -> { principal, roleId }
  plan: new Map(), // entry key -> { listId, principal, roleId }
  removals: new Map(), // `${listId}|${principalId}` -> { listId, principalId, title }: a principal whose permissions on the list are removed
  settings: new Map(), // list id -> { mode: 'add' | 'replace', keep: Set of default group keys, read, write }: how the list is changed
  preview: null, // { key, plans, errorCount }
  log: '',
};

/** The "Settings for lists you add next" box: what a list starts with. */
const defaultSettings = () => ({
  mode: ui.modes.find((m) => m.checked).value,
  keep: new Set(ui.keep.filter((k) => k.checked).map((k) => k.dataset.group)),
  read: Number(ui.read.value) || null,
  write: Number(ui.write.value) || null,
});
const copySettings = (s, list) => ({ mode: s.mode, keep: new Set(s.keep), read: list && isLibrary(list) ? null : s.read, write: list && isLibrary(list) ? null : s.write });
const hasItems = (s) => !!(s.read || s.write);
/** Options for planList: the ids of the default groups to keep come from the site. */
const listOptions = (s) => ({ mode: s.mode, keepGroups: perm.data ? [...s.keep].map((g) => perm.data.defaults[g]).filter((id) => id != null) : [] });
const planKey = () => JSON.stringify([[...perm.plan.keys()].sort(), [...perm.removals.keys()].sort(), [...perm.settings.entries()].map(([id, s]) => [id, s.mode, [...s.keep].sort(), s.read, s.write]).sort()]);
const planSize = () => perm.plan.size + perm.removals.size + [...perm.settings.values()].filter(hasItems).length;
const removalsOf = (listId) => [...perm.removals.values()].filter((r) => r.listId === listId);
const listById = (id) => perm.data.lists.find((l) => l.id === id);
const roleById = (id) => perm.data.roles.find((r) => r.id === id);
const plural = (n, word, many = `${word}s`) => `${n} ${n === 1 ? word : many}`;

// ---------------------------------------------------------------- state

function setBusy(busy) {
  perm.busy = busy;
  for (const n of document.querySelectorAll('#view-perm input, #view-perm select, #view-perm button')) {
    if (![ui.preview, ui.apply, ui.add, ui.clear, ui.applyAll, ui.usageGroup, ui.usageWith, ui.usageRun, ui.usageStop, ui.usageRemove, ui.usageReplace].includes(n)) n.disabled = busy;
  }
  if (!busy) syncKeep();
  updateButtons();
}

function setStatus(text) {
  ui.progress.hidden = false;
  ui.status.textContent = text;
}

function updateButtons() {
  const ready = perm.loaded && !perm.busy;
  const lists = perm.selected.size;
  const who = perm.picked.size;
  const items = hasItems(defaultSettings());
  ui.add.disabled = !ready || !lists || !(who || items);
  ui.add.textContent = lists && who ? `Add to plan (${plural(lists, 'list')} × ${who})` : lists && items ? `Add to plan (${plural(lists, 'list')})` : 'Add to plan';
  ui.addHint.textContent = !perm.loaded ? '' : !lists && !who && !items ? 'Select lists above and add users or groups, or set item-level permissions.'
    : !lists ? 'Select at least one list above.' : !who && !items ? 'Add at least one user or group, or set item-level permissions.' : '';
  ui.clear.disabled = !ready || !planSize();
  ui.applyAll.disabled = !ready || !perm.settings.size;
  ui.reloadRoles.disabled = !ready;
  ui.preview.disabled = !ready || !planSize();
  updateUsage();

  let text;
  let cls = '';
  let enabled = false;
  const p = perm.preview;
  if (!planSize()) text = 'Add something to the plan, then run Preview.';
  else if (!p) text = 'Run Preview first. It reads the current permissions and changes nothing.';
  else if (p.key !== planKey()) text = 'The plan or the options changed since the last Preview. Run Preview again.';
  else if (p.errorCount) {
    text = `Preview could not read ${plural(p.errorCount, 'list')}. Remove them from the plan or run Preview again.`;
    cls = 'error';
  } else if (!p.changes) text = 'Preview found nothing to change.';
  else {
    enabled = true;
    cls = 'ok';
    text = p.warningCount ? `Preview done with ${plural(p.warningCount, 'warning')}. Check them before you apply.` : 'Preview done. Ready to apply.';
  }
  ui.apply.disabled = !ready || !enabled;
  ui.applyHint.textContent = text;
  ui.applyHint.className = `hint ${cls}`.trim();
}

/** Default-group checkboxes: named after the site's groups, only usable in replace mode. */
function syncKeep() {
  const replace = defaultSettings().mode === 'replace';
  for (const k of ui.keep) {
    const id = perm.data && perm.data.defaults[k.dataset.group];
    const group = id != null && perm.data.groups.find((g) => g.id === id);
    k.closest('label').hidden = !!perm.data && !group;
    if (group) k.nextElementSibling.textContent = group.title;
    k.disabled = perm.busy || !replace;
  }
}

const keptNames = (s) => defaultGroups().filter((g) => s.keep.has(g.key)).map((g) => g.title).join(', ');

// ---------------------------------------------------------------- loading

async function load() {
  setBusy(true);
  perm.loaded = false;
  perm.data = null;
  perm.selected.clear();
  perm.picked.clear();
  perm.plan.clear();
  perm.removals.clear();
  perm.settings.clear();
  perm.preview = null;
  ui.output.hidden = ui.progress.hidden = true;
  renderAll();
  try {
    const tab = await activeTab();
    if (!tab || !tab.url || !isSharePointUrl(tab.url)) {
      ui.site.textContent = 'The active tab is not a SharePoint page. Open any page of the site, then click Reload.';
      ui.site.classList.add('error');
      return;
    }
    const site = await siteOfTab(tab);
    perm.site = site;
    perm.client = createPermissionClient({ request: tabRequester(tab.id, new URL(site).origin), onEvent: onClientEvent });
    ui.site.classList.remove('error');
    ui.site.textContent = `Reading ${site}…`;
    perm.data = await perm.client.loadSite(site);
    perm.loaded = true;
    ui.site.replaceChildren(el('strong', { text: perm.data.web.Title }), el('span', { class: 'exp-url', text: perm.data.web.Url }));
  } catch (e) {
    ui.site.textContent = `Could not read the site: ${e.message}`;
    ui.site.classList.add('error');
  } finally {
    renderAll();
    setBusy(false);
  }
}

function renderAll() {
  renderUsageGroups();
  renderLists();
  renderPicked();
  renderPlan();
}

/**
 * Read the permission levels again and keep the plan. Picked principals whose
 * level is gone get the default level; plan entries with a deleted level are
 * dropped.
 */
async function reloadRoles() {
  if (perm.busy || !perm.loaded) return;
  setBusy(true);
  setStatus('Reading the permission levels…');
  try {
    const roles = await perm.client.loadRoles(perm.site);
    const known = new Set(roles.map((r) => r.id));
    const before = new Set(perm.data.roles.map((r) => r.id));
    perm.data.roles = roles;
    for (const p of perm.picked.values()) if (!known.has(p.roleId)) p.roleId = defaultRoleId(roles);
    let dropped = 0;
    for (const [k, e] of perm.plan) {
      if (!known.has(e.roleId)) {
        perm.plan.delete(k);
        dropped++;
      }
    }
    const added = roles.filter((r) => !before.has(r.id)).length;
    const removed = [...before].filter((id) => !known.has(id)).length;
    perm.preview = null;
    renderPicked();
    renderPlan();
    setStatus(`${plural(roles.length, 'permission level')} read${added ? `, ${added} new` : ''}${removed ? `, ${removed} no longer on the site` : ''}.`
      + (dropped ? ` ${plural(dropped, 'plan entry', 'plan entries')} used a level that no longer exists and ${dropped === 1 ? 'was' : 'were'} removed.` : ''));
  } catch (e) {
    setStatus(`Could not read the permission levels: ${e.message}`);
  } finally {
    setBusy(false);
  }
}

// ---------------------------------------------------------------- 1. lists

function visibleLists() {
  if (!perm.data) return [];
  const q = ui.filter.value.trim().toLowerCase();
  return perm.data.lists.filter((l) => (ui.system.checked || !l.system) && (!q || l.title.toLowerCase().includes(q) || l.url.toLowerCase().includes(q)));
}

function renderLists() {
  const shown = visibleLists();
  ui.rows.replaceChildren(...shown.map((l) => {
    const box = el('input', { type: 'checkbox', 'aria-label': `Select ${l.title}` });
    box.checked = perm.selected.has(l.id);
    box.disabled = perm.busy;
    box.addEventListener('change', () => {
      if (box.checked) perm.selected.add(l.id);
      else perm.selected.delete(l.id);
      syncToggle();
      updateButtons();
    });
    const row = el('tr', { class: l.system ? 'system' : null },
      el('td', { class: 'sel' }, box),
      el('td', {}, el('span', { class: 'exp-title', text: l.title }), el('span', { class: 'exp-url', text: l.url })),
      el('td', { class: 'perm', text: l.uniquePermissions ? 'Unique' : 'Inherited' }));
    row.addEventListener('click', (e) => {
      if (e.target !== box && !perm.busy) box.click();
    });
    return row;
  }));
  syncToggle();
  updateButtons();
}

function syncToggle() {
  const shown = visibleLists();
  const n = shown.filter((l) => perm.selected.has(l.id)).length;
  ui.toggle.checked = shown.length > 0 && n === shown.length;
  ui.toggle.indeterminate = n > 0 && n < shown.length;
  ui.listCount.textContent = perm.data ? `${perm.selected.size} selected` : '';
}

// ---------------------------------------------------------------- 2. users and groups

/** "SPO Group" / "Entra Group" / "User" label next to a name. */
function kindBadge(p) {
  const cls = p.kind === SPO_GROUP ? 'spo' : p.kind === ENTRA_GROUP ? 'entra' : 'other';
  return el('span', { class: `perm-badge ${cls}`, text: p.kind });
}

/**
 * A search field with a result list, used by step 2 and by "Where is a group
 * used?": a click shows the site's SharePoint groups, typing filters them and
 * also searches the tenant's people and Entra groups with SharePoint's people
 * picker.
 *   local(q):  the site's groups to offer for the text `q`
 *   chosen():  the current pick; a click then shows all groups again
 *   isPicked(p): marks results that are taken already
 *   clearOnPick: empty the field after a pick (step 2 collects several)
 *   emptyNote: shown when there are no site groups and nothing typed
 */
function combo({ input, list, local, chosen = () => null, isPicked = () => false, clearOnPick = false, emptyNote = null, onPick }) {
  let seq = 0;
  let timer;
  let shown = [];
  let active = -1;
  const close = () => {
    list.hidden = true;
    input.setAttribute('aria-expanded', 'false');
    input.removeAttribute('aria-activedescendant');
    active = -1;
  };
  const choose = (p) => {
    close();
    input.value = clearOnPick ? '' : p.title;
    onPick(p);
  };
  const show = (items, note) => {
    shown = items;
    active = -1;
    list.replaceChildren(
      ...items.map((p, i) => {
        const li = el('li', { role: 'option', id: `${list.id}-${i}`, class: isPicked(p) ? 'picked' : null },
          el('span', { class: 'perm-name' }, p.title, kindBadge(p)),
          p.email || p.detail ? el('span', { class: 'perm-kind', text: p.email || p.detail }) : null);
        li.addEventListener('mousedown', (e) => {
          e.preventDefault(); // keep focus in the field
          choose(p);
        });
        return li;
      }),
      ...(note ? [el('li', { class: 'perm-note', text: note })] : [])
    );
    list.hidden = false;
    input.setAttribute('aria-expanded', 'true');
  };
  const run = () => {
    if (!perm.loaded) return;
    const c = chosen();
    const q = c && input.value === c.title ? '' : input.value.trim();
    const mine = ++seq;
    clearTimeout(timer);
    const mineGroups = local(q);
    if (q.length < 2) {
      show(mineGroups, q ? 'Type at least 2 characters to search the tenant.' : mineGroups.length ? null : emptyNote);
      return;
    }
    show(mineGroups, 'Searching the tenant…');
    timer = setTimeout(async () => {
      let found = [];
      let note = null;
      try {
        found = await perm.client.search(perm.site, q);
      } catch (e) {
        note = `Search failed: ${e.message}`;
      }
      if (mine !== seq) return;
      const all = [...mineGroups, ...found];
      show(all, note || (all.length ? null : 'No matches.'));
    }, 300);
  };
  const move = (d) => {
    if (list.hidden || !shown.length) return;
    active = (active + d + shown.length) % shown.length;
    list.querySelectorAll('[role="option"]').forEach((n, i) => n.classList.toggle('active', i === active));
    const node = document.getElementById(`${list.id}-${active}`);
    input.setAttribute('aria-activedescendant', node.id);
    node.scrollIntoView({ block: 'nearest' });
  };
  input.addEventListener('input', run);
  input.addEventListener('focus', () => { input.select(); run(); });
  input.addEventListener('blur', () => setTimeout(close, 100));
  input.addEventListener('keydown', (e) => {
    if (e.key === 'ArrowDown') move(1);
    else if (e.key === 'ArrowUp') move(-1);
    else if (e.key === 'Enter' && active >= 0) choose(shown[active]);
    else if (e.key === 'Escape' && !list.hidden) close();
    else return;
    e.preventDefault();
  });
}

/** The site's groups matching `q`; default groups only if `withDefaults` (step 2 handles those with "Keep these default groups"). */
function groupsMatching(q, withDefaults) {
  const s = q.toLowerCase();
  return perm.data.groups.filter((g) => (withDefaults || !g.isDefault) && (!s || g.title.toLowerCase().includes(s)));
}

function pick(p) {
  if (!perm.picked.has(p.key)) perm.picked.set(p.key, { principal: p, roleId: defaultRoleId(perm.data.roles) });
  renderPicked();
}

function renderPicked() {
  ui.picked.replaceChildren(...[...perm.picked.values()].map((entry) => {
    const { principal: p } = entry;
    const select = el('select', { 'aria-label': `Permission level for ${p.title}` },
      perm.data.roles.map((r) => el('option', { value: r.id, text: r.name, title: r.description || null, selected: r.id === entry.roleId })));
    select.disabled = perm.busy;
    select.addEventListener('change', () => {
      entry.roleId = Number(select.value);
    });
    const remove = el('button', { type: 'button', class: 'icon danger', title: 'Remove', 'aria-label': `Remove ${p.title}`, text: '✕' });
    remove.disabled = perm.busy;
    remove.addEventListener('click', () => {
      perm.picked.delete(p.key);
      renderPicked();
    });
    return el('li', {},
      el('span', { class: 'perm-who' }, el('span', { class: 'perm-name' }, p.title, kindBadge(p)), p.email || p.detail ? el('span', { class: 'perm-kind', text: p.email || p.detail }) : null),
      select, remove);
  }));
  updateButtons();
}

function addToPlan() {
  const lists = [...perm.selected];
  const picks = [...perm.picked.values()];
  const added = addEntries(perm.plan, lists, picks);
  const dup = lists.length * picks.length - added;
  const defaults = defaultSettings();
  const parts = [];
  if (picks.length) parts.push(`Added ${plural(added, 'permission')} to the plan.${dup ? ` ${dup} were already in it.` : ''}`);
  if (picks.length || hasItems(defaults)) {
    for (const id of lists) {
      const list = listById(id);
      const prev = perm.settings.get(id);
      if (!prev) perm.settings.set(id, copySettings(defaults, list));
      else if (list && !isLibrary(list)) {
        // a list that is already in the plan keeps its settings; only item-level values that are set here are added
        prev.read = defaults.read || prev.read;
        prev.write = defaults.write || prev.write;
      }
    }
  }
  if (hasItems(defaults)) {
    const target = lists.map(listById).filter((l) => l && !isLibrary(l));
    const libs = lists.length - target.length;
    parts.push(`Item-level permissions planned for ${plural(target.length, 'list')}.${libs ? ` ${plural(libs, 'library', 'libraries')} skipped: libraries don't have this setting.` : ''}`);
  }
  perm.selected.clear();
  perm.picked.clear();
  renderAll();
  ui.addHint.textContent = parts.join(' ');
}

// ---------------------------------------------------------------- 3. plan

/** Lists in the plan: those with planned users or groups, or with item-level settings. In list title order. */
function planByList() {
  const groups = new Map();
  for (const e of perm.plan.values()) {
    if (!groups.has(e.listId)) groups.set(e.listId, []);
    groups.get(e.listId).push(e);
  }
  for (const r of perm.removals.values()) if (!groups.has(r.listId)) groups.set(r.listId, []);
  for (const [listId, s] of perm.settings) if (hasItems(s) && !groups.has(listId)) groups.set(listId, []);
  for (const id of [...perm.settings.keys()]) if (!groups.has(id)) perm.settings.delete(id); // a list that left the plan
  for (const id of groups.keys()) if (!perm.settings.has(id)) perm.settings.set(id, copySettings(defaultSettings(), listById(id)));
  return [...groups.entries()]
    .map(([listId, entries]) => ({ list: listById(listId), entries, settings: perm.settings.get(listId) }))
    .filter((g) => g.list)
    .sort((a, b) => a.list.title.localeCompare(b.list.title, undefined, { sensitivity: 'base' }));
}

/** The default groups this site has: [{ key, title }] (owners, members, visitors). */
function defaultGroups() {
  return ui.keep.map((k) => {
    const id = perm.data && perm.data.defaults[k.dataset.group];
    const group = id != null && perm.data.groups.find((g) => g.id === id);
    return group ? { key: k.dataset.group, title: group.title } : null;
  }).filter(Boolean);
}

const openPlan = new Set(); // lists whose card is open, so a change doesn't collapse it

/** The settings of one list: current permissions, default groups to keep, item-level. */
function listSettings(list, s) {
  const lib = isLibrary(list);
  const name = `perm-mode-${list.id}`;
  const mode = (value, text) => {
    const r = el('input', { type: 'radio', name, value });
    r.checked = s.mode === value;
    r.disabled = perm.busy;
    r.addEventListener('change', () => { s.mode = value; renderPlan(); });
    return el('label', { class: 'check-label' }, r, ` ${text}`);
  };
  const keep = s.mode === 'replace' ? el('div', { class: 'perm-keep perm-indent' },
    el('span', { class: 'perm-keep-label', text: 'Keep these default groups:' }),
    ...defaultGroups().map((g) => {
      const c = el('input', { type: 'checkbox' });
      c.checked = s.keep.has(g.key);
      c.disabled = perm.busy;
      c.addEventListener('change', () => { if (c.checked) s.keep.add(g.key); else s.keep.delete(g.key); renderPlan(); });
      return el('label', { class: 'check-label' }, c, ' ', el('span', { text: g.title }));
    })) : null;
  const item = (field, src, text) => {
    const select = src.cloneNode(true);
    select.removeAttribute('id');
    select.value = s[field] ? String(s[field]) : '';
    select.disabled = perm.busy;
    select.addEventListener('change', () => { s[field] = Number(select.value) || null; renderPlan(); });
    return el('label', { class: 'perm-field' }, text, select);
  };
  return el('div', { class: 'perm-list-settings' },
    el('div', { class: 'perm-sub', text: 'This list: current permissions' }),
    mode('add', 'Keep them and add the planned ones'),
    mode('replace', 'Replace them with the planned ones'),
    keep,
    lib ? el('p', { class: 'hint', text: "Item-level permissions don't exist for libraries." }) : item('read', ui.read, 'Item-level read access'),
    lib ? null : item('write', ui.write, 'Item-level create and edit access'));
}

function renderPlan() {
  const groups = perm.data ? planByList() : [];
  const named = new Map(defaultGroups().map((g) => [g.key, g.title]));
  ui.plan.replaceChildren(...groups.map(({ list, entries, settings: s }) => {
    const rows = entries.map((e) => {
      const remove = el('button', { type: 'button', class: 'icon danger', title: 'Remove from plan', 'aria-label': `Remove ${e.principal.title} from the plan`, text: '✕' });
      remove.disabled = perm.busy;
      remove.addEventListener('click', () => {
        perm.plan.forEach((v, k) => v === e && perm.plan.delete(k));
        renderPlan();
      });
      return el('tr', {},
        el('td', {}, el('span', { class: 'perm-name', text: e.principal.title }, kindBadge(e.principal))),
        el('td', { class: 'perm-level', text: (roleById(e.roleId) || { name: e.roleId }).name }),
        el('td', { class: 'perm-x' }, remove));
    });
    const gone = removalsOf(list.id).map((r) => {
      const undo = el('button', { type: 'button', class: 'icon danger', title: 'Keep its permissions (remove from plan)', 'aria-label': `Keep ${r.title} on ${list.title}`, text: '✕' });
      undo.disabled = perm.busy;
      undo.addEventListener('click', () => { perm.removals.delete(`${r.listId}|${r.principalId}`); renderPlan(); });
      return el('tr', {},
        el('td', {}, el('span', { class: 'perm-name', text: r.title })),
        el('td', { class: 'perm-level', text: 'Remove all' }),
        el('td', { class: 'perm-x' }, undo));
    });
    const removeList = el('button', { type: 'button', class: 'secondary small', text: 'Remove list from plan' });
    removeList.disabled = perm.busy;
    removeList.addEventListener('click', () => {
      for (const e of entries) perm.plan.forEach((v, k) => v === e && perm.plan.delete(k));
      for (const r of removalsOf(list.id)) perm.removals.delete(`${r.listId}|${r.principalId}`);
      perm.settings.delete(list.id);
      renderPlan();
    });
    const kept = [...s.keep].map((k) => named.get(k)).filter(Boolean).join(', ');
    const chips = el('span', { class: 'perm-list-chips' },
      gone.length ? el('span', { class: 'chip warn', text: `remove ${gone.length}` }) : null,
      entries.length ? el('span', { class: `chip ${s.mode === 'replace' ? 'warn' : 'muted'}`, text: s.mode === 'replace' ? `replace · keeps ${kept || 'none'}` : 'add' }) : null,
      hasItems(s) ? el('span', { class: 'chip ok', title: itemLevelText(s), text: 'items: set' }) : null);
    const d = el('details', { class: 'check' },
      el('summary', {}, el('span', { class: 'chip muted', text: String(entries.length + gone.length) }), list.title,
        el('span', { class: 'summary-text', text: `${isLibrary(list) ? 'library · ' : ''}${list.uniquePermissions ? 'unique' : 'inherits'}` }), chips),
      el('div', { class: 'check-body' },
        rows.length || gone.length ? el('table', { class: 'perm-table' }, el('tbody', {}, gone, rows)) : null,
        entries.length || gone.length ? null : el('p', { class: 'hint', text: 'No users or groups planned for this list. Replace needs at least one.' }),
        listSettings(list, s), removeList));
    d.open = openPlan.has(list.id);
    d.addEventListener('toggle', () => (d.open ? openPlan.add(list.id) : openPlan.delete(list.id)));
    return d;
  }));
  ui.planEmpty.hidden = groups.length > 0;
  updateButtons();
}

// ---------------------------------------------------------------- preview

async function runPreview() {
  if (perm.busy || !planSize()) return;
  setBusy(true);
  const key = planKey();
  const groups = planByList();
  const plans = [];
  ui.bar.value = 0;
  try {
    for (const [i, { list, entries, settings: s }] of groups.entries()) {
      setStatus(`Reading the permissions of "${list.title}" (${i + 1} of ${groups.length})…`);
      ui.bar.value = i / groups.length;
      try {
        const current = await perm.client.readList(perm.site, list.id);
        list.uniquePermissions = current.unique;
        plans.push(planList(list, current, entries, { ...listOptions(s), removals: removalsOf(list.id) }, perm.data, hasItems(s) && !isLibrary(list) ? { read: s.read, write: s.write } : null));
      } catch (e) {
        plans.push({ list, error: e.message, actions: [], warnings: [] });
      }
    }
    ui.bar.value = 1;
    const ok = plans.filter((p) => !p.error);
    perm.preview = {
      key,
      plans,
      errorCount: plans.length - ok.length,
      warningCount: ok.reduce((n, p) => n + p.warnings.length, 0),
      changes: ok.some(hasChanges),
    };
    setStatus(`Read ${plural(groups.length, 'list')}. Nothing was changed.`);
    renderPreview(perm.preview);
    renderLists();
  } finally {
    setBusy(false);
  }
}

const ACTION = { add: 'Add', remove: 'Remove', keep: 'Keep', skip: 'Skip' };

function previewTable(plan) {
  const rows = [];
  for (const s of plan.settings || []) {
    rows.push(el('tr', { class: s.change ? 'act-set' : 'act-skip' },
      el('td', { class: 'perm-act', text: s.change ? 'Set' : 'Skip' }),
      el('td', {}, el('span', { class: 'perm-name', text: s.to }), el('span', { class: 'perm-kind', text: s.change ? `${s.label}, now: ${s.from}` : `${s.label}, already set` })),
      el('td', { class: 'perm-level', text: 'Item-level' })));
  }
  if (plan.breaks) rows.push(el('tr', { class: 'warn' }, el('td', { class: 'perm-act', text: 'Break' }), el('td', { colspan: 2, text: 'Stop inheriting; start from a copy of the inherited permissions' })));
  for (const a of plan.actions) {
    const roles = a.role ? a.role.name : (a.roles || []).join(', ');
    rows.push(el('tr', { class: `act-${a.type}` },
      el('td', { class: 'perm-act', text: ACTION[a.type] }),
      el('td', {}, el('span', { class: 'perm-name', text: a.title }), a.why ? el('span', { class: 'perm-kind', text: a.why }) : null),
      el('td', { class: 'perm-level', text: roles })));
  }
  return el('table', { class: 'perm-table' }, el('tbody', {}, rows));
}

function renderPreview(p) {
  ui.output.hidden = false;
  ui.copy.hidden = true;
  ui.outputHead.textContent = 'Preview (nothing changed)';
  const sections = p.plans.map((plan) => {
    if (plan.error) {
      return el('details', { class: 'stage-out', open: true },
        el('summary', {}, el('span', { class: 'chip fail', text: 'error' }), el('h3', { text: plan.list.title })),
        el('div', { class: 'inner' }, el('div', { class: 'msg fail', text: `Could not read the permissions: ${plan.error}` })));
    }
    const c = countActions(plan);
    const changes = c.add + c.remove + c.set;
    const chip = !changes ? el('span', { class: 'chip muted', text: 'no change' })
      : plan.warnings.length ? el('span', { class: 'chip warn', text: 'check' })
      : el('span', { class: 'chip ok', text: 'ready' });
    const already = c.skip + (plan.settings || []).filter((s) => !s.change).length;
    const text = [plan.breaks ? 'breaks inheritance' : null, c.set ? 'item-level' : null, c.add ? `+${c.add}` : null, c.remove ? `−${c.remove}` : null,
      already ? `${already} already set` : null].filter(Boolean).join(' · ');
    return el('details', { class: 'stage-out' },
      el('summary', {}, chip, el('h3', { text: plan.list.title }), el('span', { class: 'summary-text', text })),
      el('div', { class: 'inner' },
        plan.warnings.length ? el('div', { class: 'msg warn' }, el('ul', {}, plan.warnings.map((w) => el('li', { text: w })))) : null,
        previewTable(plan)));
  });
  const ok = p.plans.filter((x) => !x.error);
  const total = (type) => ok.reduce((n, x) => n + countActions(x)[type], 0);
  const breaks = ok.filter((x) => x.breaks).length;
  const sets = ok.filter((x) => countActions(x).set).length;
  const summary = p.errorCount
    ? el('p', { class: 'preview-summary fail', text: `${plural(p.errorCount, 'list')} could not be read.` })
    : el('p', { class: 'preview-summary ok', text: `${total('add')} to add, ${total('remove')} to remove${sets ? `, item-level permissions on ${plural(sets, 'list')}` : ''}${breaks ? `, ${plural(breaks, 'list')} ${breaks === 1 ? 'stops' : 'stop'} inheriting` : ''}.${p.warningCount ? ` ${plural(p.warningCount, 'warning')}.` : ''}` });
  ui.outputBody.replaceChildren(summary, ...sections);
  updateButtons();
}

// ---------------------------------------------------------------- apply

let countdown;
function onClientEvent(e) {
  clearInterval(countdown);
  if (!perm.busy) return;
  if (e.type === 'ensure') {
    setStatus(`Resolving "${e.title}" (${e.index + 1} of ${e.total})…`);
  } else if (e.type === 'stage-start') {
    ui.bar.value = e.stage / e.total;
    setStatus(`Batch ${e.stage + 1} of ${e.total}: preparing…`);
  } else if (e.type === 'send') {
    setStatus(`Batch ${e.stage + 1} of ${e.total}: sending ${plural(e.operations, 'operation')}${e.attempt ? ` (retry ${e.attempt})` : ''}…`);
  } else if (e.type === 'stage-done') {
    ui.bar.value = (e.stage + 1) / e.total;
  } else if (e.type === 'wait') {
    const until = Date.now() + e.ms;
    const tick = () => setStatus(`SharePoint is throttling. Waiting ${Math.max(0, Math.ceil((until - Date.now()) / 1000))} s (retry ${e.retry} of ${e.of})…`);
    tick();
    countdown = setInterval(tick, 500);
  }
}

async function runApply() {
  const p = perm.preview;
  if (perm.busy || !p || p.key !== planKey() || p.errorCount) return;
  const plans = p.plans.filter(hasChanges);
  const total = (type) => plans.reduce((n, x) => n + countActions(x)[type], 0);
  const breaks = plans.filter((x) => x.breaks).length;
  const warnings = plans.flatMap((x) => x.warnings.map((w) => `${x.list.title}: ${w}`));
  const go = await ask({
    title: 'Apply permissions?',
    wide: warnings.length > 0,
    body: [
      el('p', { class: 'target-title', text: perm.data.web.Title }),
      el('p', { class: 'target', text: perm.data.web.Url }),
      el('ul', {},
        el('li', { text: `${plural(plans.length, 'list')} ${plans.length === 1 ? 'changes' : 'change'}` }),
        breaks ? el('li', { text: `${plural(breaks, 'list')} ${breaks === 1 ? 'stops' : 'stop'} inheriting permissions` }) : null,
        total('set') ? el('li', { text: `item-level permissions change on ${plural(plans.filter((x) => countActions(x).set).length, 'list')}` }) : null,
        total('add') ? el('li', { text: `${total('add')} permission(s) added` }) : null,
        total('remove') ? el('li', { text: `${total('remove')} permission(s) removed` }) : null),
      warnings.length ? el('div', { class: 'msg warn' }, 'Warnings:', el('ul', {}, warnings.map((w) => el('li', { text: w })))) : null,
      el('p', { class: 'question', text: 'SharePoint has no undo for this. Apply now?' }),
    ].filter(Boolean),
    ok: 'Apply',
  });
  if (!go) return;

  setBusy(true);
  ui.bar.value = 0;
  setStatus('Starting…');
  try {
    const result = await perm.client.apply(perm.site, plans);
    const ops = result.lists.flatMap((l) => l.ops);
    const failed = ops.filter((o) => o.outcome === 'failed').length;
    for (const plan of plans) {
      const r = result.lists.find((l) => l.id === plan.list.id);
      const broke = r && r.ops.find((o) => /breakroleinheritance/.test(o.url));
      if (broke && broke.outcome === 'success') plan.list.uniquePermissions = true;
    }
    ui.bar.value = 1;
    setStatus(failed ? `Done with ${plural(failed, 'failure')}: ${ops.length - failed} of ${ops.length} operations succeeded.` : `Done: ${ops.length === 1 ? 'the operation' : `all ${ops.length} operations`} succeeded in ${plural(result.batches, 'batch request')}.`);
    perm.preview = null;
    perm.log = logText(result);
    renderResult(result);
    renderLists();
    renderPlan();
  } catch (e) {
    setStatus(`Apply failed: ${e.message}`);
  } finally {
    clearInterval(countdown);
    setBusy(false);
  }
}

function renderResult(result) {
  ui.output.hidden = false;
  ui.copy.hidden = false;
  ui.outputHead.textContent = 'Results';
  ui.outputBody.replaceChildren(...result.lists.map((l) => {
    const failed = l.ops.filter((o) => o.outcome === 'failed');
    const chip = failed.length ? el('span', { class: 'chip fail', text: 'failures' }) : el('span', { class: 'chip ok', text: 'done' });
    return el('details', { class: 'stage-out', open: failed.length > 0 },
      el('summary', {}, chip, el('h3', { text: l.title }),
        el('span', { class: 'summary-text', text: `${l.ops.length - failed.length} of ${l.ops.length} succeeded` })),
      el('div', { class: 'inner' }, el('table', { class: 'perm-table' }, el('tbody', {}, l.ops.map((o) =>
        el('tr', { class: o.outcome === 'failed' ? 'fail' : '' },
          el('td', { class: 'code', text: o.status == null ? '—' : o.status }),
          el('td', {}, o.label, o.outcome === 'failed' && o.errorMessage ? el('span', { class: 'err', text: o.errorMessage }) : null)))))));
  }));
}

function logText(result) {
  const lines = ['SharePoint permissions', `Site:     ${perm.site}`, `Finished: ${new Date().toISOString()}`];
  for (const l of result.lists) {
    const s = perm.settings.get(l.id);
    lines.push('', `== ${l.title}`);
    if (s) lines.push(`Mode: ${s.mode}${s.mode === 'replace' ? `, default groups kept: ${keptNames(s) || 'none'}` : ''}`);
    for (const o of l.ops) {
      lines.push(`[${o.outcome === 'failed' ? 'FAIL' : 'OK  '}] ${o.status ?? '---'} ${o.label}`);
      if (o.outcome === 'failed' && o.errorMessage) lines.push(`        ${o.errorMessage}`);
    }
  }
  return lines.join('\n') + '\n';
}

// ---------------------------------------------------------------- where is a group used?

const usage = {
  running: false,
  stop: false,
  target: null, // principal to look for
  with: null, // principal that replaces it
  last: null, // { target, principalId, res } of the last scan
  selected: new Set(), // list ids ticked in the result
};

combo({
  input: ui.usageGroup,
  list: $('usage-results'),
  local: (q) => groupsMatching(q, true),
  chosen: () => usage.target,
  onPick: (p) => { usage.target = p; clearUsageResult(); updateUsage(); },
});
ui.usageGroup.addEventListener('input', () => { usage.target = null; clearUsageResult(); updateUsage(); });
combo({
  input: ui.usageWith,
  list: $('usage-with-results'),
  local: (q) => groupsMatching(q, true),
  chosen: () => usage.with,
  onPick: (p) => { usage.with = p; updateUsage(); },
});
ui.usageWith.addEventListener('input', () => { usage.with = null; updateUsage(); });

function clearUsageResult() {
  usage.last = null;
  usage.selected.clear();
  ui.usageResult.replaceChildren();
  ui.usageActions.hidden = true;
  ui.usageStatus.textContent = '';
}

const sameUrl = (a, b) => String(a).replace(/\/+$/, '').toLowerCase() === String(b).replace(/\/+$/, '').toLowerCase();

/** Can this result row be changed here? Lists and libraries of the open site where the group itself (not one of its SharePoint groups) holds more than Limited Access. */
const changeable = (p) => !!usage.last && !p.via && (p.type === 'list' || p.type === 'library') && !p.limited && sameUrl(p.webUrl, perm.data.web.Url) && !!listById(p.id);

function updateUsage() {
  const ready = perm.loaded && !usage.running;
  ui.usageRun.disabled = !ready || perm.busy || !usage.target;
  ui.usageGroup.disabled = ui.usageWith.disabled = !ready;
  ui.usageStop.hidden = !usage.running;
  const n = usage.selected.size;
  ui.usageRemove.disabled = !ready || perm.busy || !n;
  ui.usageRemove.textContent = n ? `Remove from ${plural(n, 'selected list')}` : 'Remove from selected';
  ui.usageReplace.disabled = !ready || perm.busy || !n || !usage.with;
  ui.usageReplace.textContent = n ? `Replace in ${plural(n, 'selected list')}` : 'Replace in selected';
}

function renderUsageGroups() {
  usage.target = usage.with = null;
  ui.usageGroup.value = ui.usageWith.value = '';
  ui.usagePlanHint.textContent = '';
  clearUsageResult();
}

const TYPE_LABEL = { web: 'Site', list: 'List', library: 'Library' };

function usageTable(places) {
  const rows = places.map((p) => {
    let box = null;
    if (changeable(p)) {
      box = el('input', { type: 'checkbox', 'aria-label': `Select ${p.title}` });
      box.checked = usage.selected.has(p.id);
      box.addEventListener('change', () => {
        if (box.checked) usage.selected.add(p.id);
        else usage.selected.delete(p.id);
        syncUsageToggle();
        updateUsage();
      });
    }
    const row = el('tr', { class: box ? null : 'system' },
      el('td', { class: 'sel' }, box),
      el('td', { class: 'kind', text: TYPE_LABEL[p.type] }),
      el('td', {},
        el('span', { class: 'exp-title', text: p.title }),
        el('span', { class: 'exp-url', text: p.url }),
        p.type !== 'web' && !sameUrl(p.webUrl, p.url) ? el('span', { class: 'exp-url', text: `in ${p.webTitle}` }) : null),
      el('td', {}, ...[p.via ? el('span', { class: 'chip ok', title: `Not given to the group itself: it is a member of the SharePoint group "${p.via}", which has these permissions. Change that group's permissions or its members instead.`, text: `via ${p.via}` }) : null, ...p.roles.map((r) => el('span', { class: 'chip muted', text: r })), p.limited ? el('span', { class: 'chip warn', title: 'SharePoint adds Limited Access itself so that a person can reach the item below; it is not a permission that was granted.', text: 'only Limited Access' }) : null].filter(Boolean)));
    if (box) row.addEventListener('click', (e) => { if (e.target !== box) box.click(); });
    return row;
  });
  const toggle = el('input', { type: 'checkbox', id: 'usage-toggle', 'aria-label': 'Select all lists and libraries of this site' });
  toggle.addEventListener('change', () => {
    for (const p of places.filter(changeable)) toggle.checked ? usage.selected.add(p.id) : usage.selected.delete(p.id);
    ui.usageResult.replaceChildren(usageTable(places), ...[...ui.usageResult.children].slice(1));
    syncUsageToggle();
    updateUsage();
  });
  return el('div', { class: 'exp-table-wrap' },
    el('table', { class: 'exp-table' },
      el('thead', {}, el('tr', {}, el('th', { class: 'sel' }, toggle), el('th', { class: 'kind', text: 'Type' }), el('th', { text: 'Where' }), el('th', { text: 'Permission levels' }))),
      el('tbody', {}, ...rows)));
}

function syncUsageToggle() {
  const t = $('usage-toggle');
  if (!t || !usage.last) return;
  const all = usage.last.res.places.filter(changeable);
  const n = all.filter((p) => usage.selected.has(p.id)).length;
  t.checked = all.length > 0 && n === all.length;
  t.indeterminate = n > 0 && n < all.length;
}

async function runUsage() {
  if (usage.running || !perm.loaded || !usage.target) return;
  const t = usage.target;
  const who = t.id != null ? t.id : { login: t.login };
  const name = t.title;
  usage.running = true;
  usage.stop = false;
  clearUsageResult();
  ui.usagePlanHint.textContent = '';
  updateButtons();
  ui.usageStatus.classList.remove('error');
  ui.usageStatus.textContent = 'Reading the site collection…';
  try {
    const res = await perm.client.findGroupUsage(perm.site, who, {
      shouldStop: () => usage.stop,
      onProgress: (e) => {
        if (e.type === 'web') ui.usageStatus.textContent = `Looking at ${e.url} (${plural(e.webs, 'site')} done, ${e.queued} waiting)…`;
      },
    });
    if (res.notInSite) {
      ui.usageStatus.textContent = `"${name}" has no permissions in this site collection: it is not one of its users or groups, so nothing was crawled.`;
      return;
    }
    usage.last = { target: t, principalId: res.principalId, res };
    const real = res.places.filter((p) => !p.limited);
    const limited = res.places.length - real.length;
    const viaCount = real.filter((p) => p.via).length;
    ui.usageStatus.textContent = `${res.stopped ? 'Stopped. ' : ''}"${name}": ${real.length ? `permissions in ${plural(real.length, 'place')}` : 'no permissions found'}`
      + `${viaCount ? ` (${viaCount} via ${res.viaGroups.length === 1 ? 'a SharePoint group' : 'its SharePoint groups'})` : ''}${limited ? `, ${limited} more with only Limited Access` : ''}. Looked at ${plural(res.webs, 'site')} and ${plural(res.lists, 'list')}${res.stopped ? ' so far' : ''}.`;
    ui.usageResult.replaceChildren(
      ...(res.places.length ? [usageTable(res.places)] : []),
      el('p', { class: 'hint', text: 'Sites, lists and libraries that inherit their permissions are not listed: the group shows at the parent that defines them. A group with no result can still be used on items or columns, in sharing links, alerts or flows.' }),
      ...(res.errors.length ? [el('p', { class: 'hint error', text: `${plural(res.errors.length, 'part')} could not be read (missing access?):` }), el('ul', { class: 'hint' }, ...res.errors.map((x) => el('li', { text: `${x.where}: ${x.message}` })))] : [])
    );
    ui.usageActions.hidden = !res.places.some(changeable);
  } catch (e) {
    ui.usageStatus.textContent = `Could not scan: ${e.message}`;
    ui.usageStatus.classList.add('error');
  } finally {
    usage.running = false;
    updateButtons();
  }
}

/**
 * Put the group's removal (and the replacement with the same permission levels)
 * for the ticked lists into the plan. Lists new to the plan start in "add"
 * mode without item-level changes, whatever the settings box says, so nothing
 * but the group is touched.
 */
function planUsage(replace) {
  const { principalId, res } = usage.last;
  const w = replace ? usage.with : null;
  const rows = res.places.filter((p) => changeable(p) && usage.selected.has(p.id));
  let skipped = 0;
  for (const p of rows) {
    perm.removals.set(`${p.id}|${principalId}`, { listId: p.id, principalId, title: usage.last.target.title });
    if (w) {
      const roleIds = p.roleList.filter((r) => !r.hidden && roleById(r.id)).map((r) => r.id);
      if (roleIds.length < p.roleList.filter((r) => !r.hidden).length) skipped++;
      addEntries(perm.plan, [p.id], roleIds.map((roleId) => ({ principal: w, roleId })));
    }
    if (!perm.settings.has(p.id)) perm.settings.set(p.id, { mode: 'add', keep: new Set(), read: null, write: null });
  }
  perm.preview = null;
  ui.output.hidden = true;
  usage.selected.clear();
  ui.usageResult.replaceChildren(usageTable(res.places), ...[...ui.usageResult.children].slice(1));
  syncUsageToggle();
  renderPlan();
  updateUsage();
  ui.usagePlanHint.textContent = `${replace ? 'Replacing' : 'Removing'} "${usage.last.target.title}"${w ? ` with "${w.title}"` : ''} is in the plan for ${plural(rows.length, 'list')}. Run Preview in section 3 to check it.`
    + (skipped ? ` ${plural(skipped, 'list')} had a permission level that no longer exists; it was not copied.` : '');
}

// ---------------------------------------------------------------- events

ui.usageRun.addEventListener('click', runUsage);
ui.usageStop.addEventListener('click', () => { usage.stop = true; ui.usageStatus.textContent = 'Stopping…'; });
ui.usageRemove.addEventListener('click', () => planUsage(false));
ui.usageReplace.addEventListener('click', () => planUsage(true));

document.addEventListener('sp-utility:tab-shown', (e) => {
  if (e.detail.id === 'perm' && !perm.loaded && !perm.busy) load();
});

ui.reloadRoles.addEventListener('click', reloadRoles);
ui.reload.addEventListener('click', async () => {
  if (planSize() && !(await ask({ title: 'Reload the site?', body: 'This clears the plan.', ok: 'Reload' }))) return;
  load();
});
ui.filter.addEventListener('input', renderLists);
ui.system.addEventListener('change', () => {
  if (!ui.system.checked && perm.data) for (const l of perm.data.lists) if (l.system) perm.selected.delete(l.id);
  renderLists();
});
ui.toggle.addEventListener('change', () => {
  for (const l of visibleLists()) ui.toggle.checked ? perm.selected.add(l.id) : perm.selected.delete(l.id);
  renderLists();
});

combo({
  input: ui.search,
  list: ui.results,
  local: (q) => groupsMatching(q, false),
  isPicked: (p) => perm.picked.has(p.key),
  clearOnPick: true,
  emptyNote: 'No site groups besides the default groups. Type to search the tenant.',
  onPick: pick,
});

ui.add.addEventListener('click', addToPlan);
ui.clear.addEventListener('click', async () => {
  if (!(await ask({ title: 'Clear the plan?', body: `Everything planned (${plural(planSize(), 'entry', 'entries')}) is removed from the plan. Nothing changes in SharePoint.`, ok: 'Clear' }))) return;
  perm.plan.clear();
  perm.removals.clear();
  perm.settings.clear();
  perm.preview = null;
  ui.output.hidden = true;
  renderPlan();
});
for (const m of ui.modes) m.addEventListener('change', syncKeep);
for (const n of [...ui.modes, ...ui.keep, ui.read, ui.write]) n.addEventListener('change', updateButtons);
ui.applyAll.addEventListener('click', () => {
  const defaults = defaultSettings();
  for (const [id, s] of perm.settings) perm.settings.set(id, copySettings(defaults, listById(id)));
  perm.preview = null;
  renderPlan();
  setStatus(`Applied the settings to ${plural(perm.settings.size, 'list')} in the plan.`);
});
ui.preview.addEventListener('click', runPreview);
ui.apply.addEventListener('click', runApply);
ui.copy.addEventListener('click', () => navigator.clipboard.writeText(perm.log).then(
  () => { ui.copy.textContent = 'Copied'; setTimeout(() => { ui.copy.textContent = 'Copy log'; }, 1500); },
  () => {}
));
