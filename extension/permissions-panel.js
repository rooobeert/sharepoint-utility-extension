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
  itemsChip: $('perm-items-chip'),
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
  items: new Map(), // list id -> { read, write } item-level settings
  preview: null, // { key, plans, errorCount }
  log: '',
};

const options = () => ({
  mode: ui.modes.find((m) => m.checked).value,
  keepGroups: perm.data ? ui.keep.filter((k) => k.checked).map((k) => perm.data.defaults[k.dataset.group]).filter((id) => id != null) : [],
});
const planKey = () => JSON.stringify([[...perm.plan.keys()].sort(), [...perm.items.entries()].sort(), options()]);
const itemLevel = () => ({ read: Number(ui.read.value) || null, write: Number(ui.write.value) || null });
const planSize = () => perm.plan.size + perm.items.size;
const listById = (id) => perm.data.lists.find((l) => l.id === id);
const roleById = (id) => perm.data.roles.find((r) => r.id === id);
const plural = (n, word, many = `${word}s`) => `${n} ${n === 1 ? word : many}`;

// ---------------------------------------------------------------- state

function setBusy(busy) {
  perm.busy = busy;
  for (const n of document.querySelectorAll('#view-perm input, #view-perm select, #view-perm button')) {
    if (n !== ui.preview && n !== ui.apply && n !== ui.add && n !== ui.clear) n.disabled = busy;
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
  const il = itemLevel();
  const items = !!(il.read || il.write);
  ui.itemsChip.textContent = items ? 'set' : 'off';
  ui.itemsChip.className = `chip ${items ? 'ok' : 'muted'}`;
  ui.add.disabled = !ready || !lists || !(who || items);
  ui.add.textContent = lists && who ? `Add to plan (${plural(lists, 'list')} × ${who})` : lists && items ? `Add to plan (${plural(lists, 'list')})` : 'Add to plan';
  ui.addHint.textContent = !perm.loaded ? '' : !lists && !who && !items ? 'Select lists above and add users or groups, or set item-level permissions.'
    : !lists ? 'Select at least one list above.' : !who && !items ? 'Add at least one user or group, or set item-level permissions.' : '';
  ui.clear.disabled = !ready || !planSize();
  ui.reloadRoles.disabled = !ready;
  ui.preview.disabled = !ready || !planSize();

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
  const replace = options().mode === 'replace';
  for (const k of ui.keep) {
    const id = perm.data && perm.data.defaults[k.dataset.group];
    const group = id != null && perm.data.groups.find((g) => g.id === id);
    k.closest('label').hidden = !!perm.data && !group;
    if (group) k.nextElementSibling.textContent = group.title;
    k.disabled = perm.busy || !replace;
  }
}

const keptNames = () => ui.keep.filter((k) => k.checked && !k.closest('label').hidden).map((k) => k.nextElementSibling.textContent).join(', ');

// ---------------------------------------------------------------- loading

async function load() {
  setBusy(true);
  perm.loaded = false;
  perm.data = null;
  perm.selected.clear();
  perm.picked.clear();
  perm.plan.clear();
  perm.items.clear();
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

let searchSeq = 0;
let searchTimer;
let options_ = []; // principals currently shown in the result list
let active = -1;

function closeResults() {
  ui.results.hidden = true;
  ui.search.setAttribute('aria-expanded', 'false');
  ui.search.removeAttribute('aria-activedescendant');
  active = -1;
}

function showResults(principals, note) {
  options_ = principals;
  active = -1;
  ui.results.replaceChildren(
    ...principals.map((p, i) => {
      const li = el('li', { role: 'option', id: `perm-opt-${i}`, class: perm.picked.has(p.key) ? 'picked' : null },
        el('span', { class: 'perm-name' }, p.title, kindBadge(p)),
        p.email || p.detail ? el('span', { class: 'perm-kind', text: p.email || p.detail }) : null);
      li.addEventListener('mousedown', (e) => {
        e.preventDefault(); // keep focus in the search field
        pick(p);
      });
      return li;
    }),
    ...(note ? [el('li', { class: 'perm-note', text: note })] : [])
  );
  ui.results.hidden = false;
  ui.search.setAttribute('aria-expanded', 'true');
}

/** "SPO Group" / "Entra Group" / "User" label next to a name. */
function kindBadge(p) {
  const cls = p.kind === SPO_GROUP ? 'spo' : p.kind === ENTRA_GROUP ? 'entra' : 'other';
  return el('span', { class: `perm-badge ${cls}`, text: p.kind });
}

/** The site's SharePoint groups matching `q`, without the default groups (those are handled by "Keep these default groups"). */
function siteGroups(q) {
  const s = q.toLowerCase();
  return perm.data.groups.filter((g) => !g.isDefault && (!s || g.title.toLowerCase().includes(s)));
}

function runSearch() {
  if (!perm.loaded) return;
  const q = ui.search.value.trim();
  const seq = ++searchSeq;
  clearTimeout(searchTimer);
  const local = siteGroups(q);
  if (q.length < 2) {
    showResults(local, q ? 'Type at least 2 characters to search the tenant.' : local.length ? null : 'No site groups besides the default groups. Type to search the tenant.');
    return;
  }
  showResults(local, 'Searching the tenant…');
  searchTimer = setTimeout(async () => {
    let found = [];
    let note = null;
    try {
      found = await perm.client.search(perm.site, q);
    } catch (e) {
      note = `Search failed: ${e.message}`;
    }
    if (seq !== searchSeq) return;
    const all = [...local, ...found];
    showResults(all, note || (all.length ? null : 'No matches.'));
  }, 300);
}

function pick(p) {
  if (!perm.picked.has(p.key)) perm.picked.set(p.key, { principal: p, roleId: defaultRoleId(perm.data.roles) });
  ui.search.value = '';
  closeResults();
  renderPicked();
}

function moveActive(delta) {
  if (ui.results.hidden || !options_.length) return;
  active = (active + delta + options_.length) % options_.length;
  ui.results.querySelectorAll('[role="option"]').forEach((n, i) => n.classList.toggle('active', i === active));
  const node = $(`perm-opt-${active}`);
  ui.search.setAttribute('aria-activedescendant', node.id);
  node.scrollIntoView({ block: 'nearest' });
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
  const il = itemLevel();
  const parts = [];
  if (picks.length) parts.push(`Added ${plural(added, 'permission')} to the plan.${dup ? ` ${dup} were already in it.` : ''}`);
  if (il.read || il.write) {
    const target = lists.map(listById).filter((l) => !isLibrary(l));
    for (const l of target) {
      const prev = perm.items.get(l.id) || {};
      perm.items.set(l.id, { read: il.read || prev.read || null, write: il.write || prev.write || null });
    }
    const libs = lists.length - target.length;
    parts.push(`Item-level permissions planned for ${plural(target.length, 'list')}.${libs ? ` ${plural(libs, 'library', 'libraries')} skipped: libraries don't have this setting.` : ''}`);
    ui.read.value = ui.write.value = '';
  }
  perm.selected.clear();
  perm.picked.clear();
  renderAll();
  ui.addHint.textContent = parts.join(' ');
}

// ---------------------------------------------------------------- 3. plan

/** Plan entries grouped by list, in list title order. */
function planByList() {
  const groups = new Map();
  for (const e of perm.plan.values()) {
    if (!groups.has(e.listId)) groups.set(e.listId, []);
    groups.get(e.listId).push(e);
  }
  for (const listId of perm.items.keys()) if (!groups.has(listId)) groups.set(listId, []);
  return [...groups.entries()]
    .map(([listId, entries]) => ({ list: listById(listId), entries, itemLevel: perm.items.get(listId) || null }))
    .filter((g) => g.list)
    .sort((a, b) => a.list.title.localeCompare(b.list.title, undefined, { sensitivity: 'base' }));
}

function renderPlan() {
  const groups = perm.data ? planByList() : [];
  ui.plan.replaceChildren(...groups.map(({ list, entries, itemLevel: il }) => {
    const rows = entries.map((e) => {
      const remove = el('button', { type: 'button', class: 'icon danger', title: 'Remove from plan', 'aria-label': `Remove ${e.principal.title} from the plan`, text: '✕' });
      remove.disabled = perm.busy;
      remove.addEventListener('click', () => {
        perm.plan.forEach((v, k) => v === e && perm.plan.delete(k));
        renderPlan();
      });
      return el('tr', {},
        el('td', {}, el('span', { class: 'perm-name' }, e.principal.title, kindBadge(e.principal))),
        el('td', { class: 'perm-level', text: (roleById(e.roleId) || { name: e.roleId }).name }),
        el('td', { class: 'perm-x' }, remove));
    });
    if (il) {
      const remove = el('button', { type: 'button', class: 'icon danger', title: 'Remove from plan', 'aria-label': 'Remove the item-level permissions from the plan', text: '✕' });
      remove.disabled = perm.busy;
      remove.addEventListener('click', () => {
        perm.items.delete(list.id);
        renderPlan();
      });
      rows.unshift(el('tr', {},
        el('td', {}, el('span', { class: 'perm-name', text: 'Item-level permissions' }), el('span', { class: 'perm-kind', text: itemLevelText(il) })),
        el('td', { class: 'perm-level', text: '' }),
        el('td', { class: 'perm-x' }, remove)));
    }
    const removeList = el('button', { type: 'button', class: 'secondary small', text: 'Remove list from plan' });
    removeList.disabled = perm.busy;
    removeList.addEventListener('click', () => {
      for (const e of entries) perm.plan.forEach((v, k) => v === e && perm.plan.delete(k));
      perm.items.delete(list.id);
      renderPlan();
    });
    return el('details', { class: 'check' },
      el('summary', {}, el('span', { class: 'chip muted', text: String(entries.length + (il ? 1 : 0)) }), list.title,
        el('span', { class: 'summary-text', text: list.uniquePermissions ? 'unique' : 'inherits' })),
      el('div', { class: 'check-body' }, el('table', { class: 'perm-table' }, el('tbody', {}, rows)), removeList));
  }));
  ui.planEmpty.hidden = groups.length > 0;
  updateButtons();
}

// ---------------------------------------------------------------- preview

async function runPreview() {
  if (perm.busy || !planSize()) return;
  setBusy(true);
  const key = planKey();
  const opts = options();
  const groups = planByList();
  const plans = [];
  ui.bar.value = 0;
  try {
    for (const [i, { list, entries, itemLevel: il }] of groups.entries()) {
      setStatus(`Reading the permissions of "${list.title}" (${i + 1} of ${groups.length})…`);
      ui.bar.value = i / groups.length;
      try {
        const current = await perm.client.readList(perm.site, list.id);
        list.uniquePermissions = current.unique;
        plans.push(planList(list, current, entries, opts, perm.data, il));
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
  const lines = ['SharePoint permissions', `Site:     ${perm.site}`, `Finished: ${new Date().toISOString()}`, `Mode:     ${options().mode}${options().mode === 'replace' ? `, default groups kept: ${keptNames() || 'none'}` : ''}`];
  for (const l of result.lists) {
    lines.push('', `== ${l.title}`);
    for (const o of l.ops) {
      lines.push(`[${o.outcome === 'failed' ? 'FAIL' : 'OK  '}] ${o.status ?? '---'} ${o.label}`);
      if (o.outcome === 'failed' && o.errorMessage) lines.push(`        ${o.errorMessage}`);
    }
  }
  return lines.join('\n') + '\n';
}

// ---------------------------------------------------------------- events

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

ui.search.addEventListener('input', runSearch);
ui.search.addEventListener('focus', runSearch);
ui.search.addEventListener('blur', () => setTimeout(closeResults, 100));
ui.search.addEventListener('keydown', (e) => {
  if (e.key === 'ArrowDown') moveActive(1);
  else if (e.key === 'ArrowUp') moveActive(-1);
  else if (e.key === 'Enter' && active >= 0) pick(options_[active]);
  else if (e.key === 'Escape' && !ui.results.hidden) closeResults();
  else return;
  e.preventDefault();
});

ui.add.addEventListener('click', addToPlan);
ui.clear.addEventListener('click', async () => {
  if (!(await ask({ title: 'Clear the plan?', body: `Everything planned (${plural(planSize(), 'entry', 'entries')}) is removed from the plan. Nothing changes in SharePoint.`, ok: 'Clear' }))) return;
  perm.plan.clear();
  perm.items.clear();
  perm.preview = null;
  ui.output.hidden = true;
  renderPlan();
});
for (const m of ui.modes) {
  m.addEventListener('change', () => {
    syncKeep();
    updateButtons();
  });
}
for (const k of ui.keep) k.addEventListener('change', updateButtons);
ui.read.addEventListener('change', updateButtons);
ui.write.addEventListener('change', updateButtons);
ui.preview.addEventListener('click', runPreview);
ui.apply.addEventListener('click', runApply);
ui.copy.addEventListener('click', () => navigator.clipboard.writeText(perm.log).then(
  () => { ui.copy.textContent = 'Copied'; setTimeout(() => { ui.copy.textContent = 'Copy log'; }, 1500); },
  () => {}
));
