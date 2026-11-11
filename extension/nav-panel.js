// "Navigation" tab: build the Quick Launch or top navigation from the site's
// lists, grouped by their solution tag (hrpLocations -> HRP), with drag and drop.

import { el, activeTab, tabRequester, ask, siteOfTab } from './shared.js';
import { isSharePointUrl } from './lib/batch.js';
import { createExporter } from './lib/exporter.js';
import { createNavigator } from './lib/navigator.js';
import {
  NAV_TARGETS,
  MAX_DEPTH,
  detectNavigation,
  groupListsByPrefix,
  linksToList,
  treeFromNodes,
  cloneTree,
  treeSignature,
  walkTree,
  subtreeHeight,
  findNode,
  validateTree,
  planSave,
  planIsEmpty,
  newKey,
  editorUrl,
  exportNavigation,
  parseNavigationExport,
  backupKey,
  snapshotTree,
  countNodes,
  addBackup,
  treeFromSnapshot,
} from './lib/navigation.js';

const $ = (id) => document.getElementById(id);

const ui = {
  site: $('nav-site'),
  reload: $('nav-reload'),
  targetName: $('nav-target-name'),
  switchTarget: $('nav-switch'),
  targetHint: $('nav-target-hint'),
  filter: $('nav-filter'),
  listCount: $('nav-list-count'),
  groups: $('nav-groups'),
  tree: $('nav-tree'),
  treeEmpty: $('nav-empty'),
  dirty: $('nav-dirty'),
  addHeading: $('nav-add-heading'),
  addLink: $('nav-add-link'),
  copy: $('nav-copy'),
  paste: $('nav-paste'),
  restore: $('nav-restore'),
  discard: $('nav-discard'),
  save: $('nav-save'),
  saveHint: $('nav-save-hint'),
  progress: $('nav-progress'),
  bar: $('nav-bar'),
  status: $('nav-status'),
  result: $('nav-result'),
};

const nav = {
  busy: false,
  loaded: false,
  site: null, // web URL
  origin: null,
  request: null, // runs requests in the SharePoint tab
  web: null,
  lists: [],
  target: null, // 'quicklaunch' | 'topnav'
  detected: null, // detectNavigation() of the site
  original: [], // tree as read from SharePoint
  tree: [], // tree being edited
  open: new Set(), // expanded list groups (by label)
  backups: [], // earlier navigations of this site and target, newest first: { at, tree }
};

const isDirty = () => nav.loaded && treeSignature(nav.tree) !== treeSignature(nav.original);
const targetLabel = () => NAV_TARGETS[nav.target]?.label || 'navigation';

// ---------------------------------------------------------------- state

function setBusy(busy) {
  nav.busy = busy;
  for (const n of [ui.reload, ui.filter, ui.addHeading, ui.addLink, ui.discard, ui.switchTarget, ui.copy, ui.paste, ui.restore]) n.disabled = busy;
  for (const n of document.querySelectorAll('#view-nav .nav-group button, #nav-tree input, #nav-tree button')) n.disabled = busy;
  updateSaveState();
}

function setStatus(text) {
  ui.progress.hidden = false;
  ui.status.textContent = text;
}

function updateSaveState() {
  const dirty = isDirty();
  ui.save.disabled = nav.busy || !nav.loaded || !dirty;
  ui.discard.disabled = nav.busy || !dirty;
  ui.addHeading.disabled = ui.addLink.disabled = ui.paste.disabled = nav.busy || !nav.loaded;
  ui.copy.disabled = nav.busy || !nav.loaded || !nav.tree.length;
  ui.restore.disabled = nav.busy || !nav.loaded || !nav.backups.length;
  ui.dirty.textContent = dirty ? 'Unsaved changes' : '';
  const problems = nav.loaded ? validateTree(nav.tree) : [];
  ui.saveHint.className = problems.length ? 'hint error' : 'hint';
  ui.saveHint.textContent = !nav.loaded
    ? ''
    : problems.length
      ? problems[0] + (problems.length > 1 ? ` (and ${problems.length - 1} more)` : '')
      : dirty
        ? `Save writes this tree to the ${targetLabel()} of the site. You see what will change before anything is sent.`
        : `This is the current ${targetLabel()} of the site.`;
}

/** Call after any change to the tree. */
function changed({ structure = true } = {}) {
  if (structure) renderTree();
  renderGroups();
  updateSaveState();
}

// ---------------------------------------------------------------- backups

const BACKUPS_KEY = 'navBackups'; // { "<site>|<target>": [{ at, tree }] }, kept in this browser's extension storage

async function readBackups() {
  try {
    return (await chrome.storage.local.get(BACKUPS_KEY))[BACKUPS_KEY] || {};
  } catch {
    return {};
  }
}

async function refreshBackups() {
  nav.backups = nav.site && nav.target ? (await readBackups())[backupKey(nav.site, nav.target)] || [] : [];
  ui.restore.textContent = nav.backups.length ? `Restore… (${nav.backups.length})` : 'Restore…';
  updateSaveState();
}

/** Keep the navigation the site has now, before a Save changes it. Returns false if it could not be stored. */
async function backupCurrent() {
  if (!nav.original.length) return true; // nothing on the site to lose
  try {
    const all = await readBackups();
    await chrome.storage.local.set({ [BACKUPS_KEY]: addBackup(all, backupKey(nav.site, nav.target), { at: new Date().toISOString(), tree: snapshotTree(nav.original) }) });
    return true;
  } catch {
    return false;
  }
}

// ---------------------------------------------------------------- loading

async function loadAll() {
  setBusy(true);
  nav.loaded = false;
  ui.result.replaceChildren();
  try {
    const tab = await activeTab();
    if (!tab || !tab.url || !isSharePointUrl(tab.url)) {
      ui.site.textContent = 'The active tab is not a SharePoint page. Open any page of the site, then click Reload.';
      ui.site.classList.add('error');
      return;
    }
    const site = await siteOfTab(tab);
    if (site !== nav.site) nav.target = null; // a new site gets its own default
    nav.site = site;
    nav.origin = new URL(site).origin;
    nav.request = tabRequester(tab.id, nav.origin);
    ui.site.classList.remove('error');
    ui.site.textContent = `Reading lists of ${site}…`;
    const data = await createExporter({ request: nav.request }).loadLists(site);
    nav.web = data.web;
    nav.lists = data.lists;
    ui.site.replaceChildren(el('strong', { text: data.web.Title }), el('span', { class: 'exp-url', text: data.web.Url }));
    nav.detected = detectNavigation(await createNavigator({ request: nav.request }).settings(site));
    if (!nav.target) nav.target = 'quicklaunch'; // almost always what is edited, whatever the site detects
    await loadNav();
    await refreshBackups();
  } catch (e) {
    ui.site.textContent = `Could not read the site: ${e.message}`;
    ui.site.classList.add('error');
  } finally {
    setBusy(false);
  }
}

async function loadNav() {
  const d = nav.detected;
  const other = nav.target === 'topnav' ? 'quicklaunch' : 'topnav';
  ui.targetName.textContent = NAV_TARGETS[nav.target].label + (nav.target === d.target ? '' : ' (chosen instead of the detected one)');
  ui.targetHint.textContent = `${d.text} Hub navigation is not changed here.`;
  ui.switchTarget.hidden = false;
  ui.switchTarget.textContent = other === d.target ? `Back to the ${NAV_TARGETS[other].label}` : `Edit the ${NAV_TARGETS[other].label} instead`;
  setStatus(`Reading the ${targetLabel()}…`);
  ui.bar.removeAttribute('value');
  const nodes = await createNavigator({ request: nav.request }).load(nav.site, nav.target);
  nav.original = treeFromNodes(nodes);
  nav.tree = cloneTree(nav.original);
  nav.loaded = true;
  ui.bar.value = 1;
  let count = 0;
  walkTree(nav.tree, () => count++);
  setStatus(`Read ${count} node(s) from the ${targetLabel()}.`);
  changed();
}

/** Ask before throwing away edits. Resolves true when it is fine to go on. */
async function confirmDiscard(what) {
  if (!isDirty()) return true;
  return ask({ title: 'Discard your changes?', body: `The ${targetLabel()} has unsaved changes. ${what} discards them.`, ok: 'Discard', cancel: 'Keep editing' });
}

// ---------------------------------------------------------------- lists palette

const isLinked = (list) => {
  let hit = false;
  walkTree(nav.tree, (n) => (hit = hit || linksToList(n.url, list.url, nav.origin)));
  return hit;
};
const isLibrary = (list) => [101, 109, 119, 851].includes(list.template);

function renderGroups() {
  if (!nav.loaded) return ui.groups.replaceChildren();
  const q = ui.filter.value.trim().toLowerCase();
  const groups = groupListsByPrefix(nav.lists)
    .map((g) => ({ ...g, shown: g.lists.filter((l) => !q || l.title.toLowerCase().includes(q) || l.url.toLowerCase().includes(q) || g.label.toLowerCase().includes(q)) }))
    .filter((g) => g.shown.length);
  const total = nav.lists.filter((l) => !l.system).length;
  ui.listCount.textContent = `${total} list(s)`;
  ui.groups.replaceChildren(
    ...groups.map((g) => {
      const linked = g.lists.filter(isLinked).length;
      const details = el('details', { class: 'check nav-group', open: nav.open.has(g.label) || !!q },
        el('summary', { draggable: g.prefix ? 'true' : null, title: g.prefix ? 'Drag into the navigation to add a heading with these lists' : null },
          el('span', { class: 'nav-group-label', text: g.label }),
          el('span', { class: `chip ${linked === g.lists.length ? 'ok' : 'muted'}`, text: `${linked}/${g.lists.length}`, title: `${linked} of ${g.lists.length} already in the navigation` }),
          g.prefix ? groupButton(g) : null),
        el('ul', { class: 'nav-lists' }, g.shown.map((l) => listRow(l, g))));
      details.addEventListener('toggle', () => (details.open ? nav.open.add(g.label) : nav.open.delete(g.label)));
      if (g.prefix) dragSource(details.querySelector('summary'), () => ({ kind: 'group', group: g }));
      return details;
    })
  );
  if (!groups.length) ui.groups.append(el('p', { class: 'empty', text: q ? 'No lists match.' : 'The site has no lists or libraries.' }));
  for (const b of ui.groups.querySelectorAll('button')) b.disabled = nav.busy;
}

function groupButton(g) {
  const b = el('button', { type: 'button', class: 'secondary small nav-add', text: '+ Add as menu', title: `Add a "${g.label}" heading with all its lists (missing ones only)` });
  b.addEventListener('click', (e) => {
    e.preventDefault(); // don't toggle the group
    addGroup(g);
  });
  return b;
}

function listRow(l, g) {
  const linked = isLinked(l);
  const add = el('button', {
    type: 'button',
    class: 'icon',
    text: '+',
    'aria-label': `Add ${l.title}`,
    title: g.prefix ? `Add under the "${g.label}" heading` : 'Add at the end of the navigation',
  });
  add.addEventListener('click', () => addList(l, g));
  const row = el('li', { class: linked ? 'linked' : null, draggable: 'true', title: 'Drag into the navigation' },
    el('span', { class: 'nav-list-text' },
      el('span', { class: 'exp-title', text: l.title }),
      el('span', { class: 'exp-url', text: l.url.split('/').pop() + (isLibrary(l) ? ' · library' : '') })),
    linked ? el('span', { class: 'chip ok', text: 'in nav' }) : null,
    add);
  dragSource(row, () => ({ kind: 'list', list: l }));
  return row;
}

const listNode = (l) => ({ key: newKey(), id: null, title: l.title, url: l.url, children: [] });

/** Root-level heading with this title (case-insensitive), created at the end if missing. */
function headingFor(label) {
  let h = nav.tree.find((n) => n.title.trim().toLowerCase() === label.toLowerCase());
  if (!h) {
    h = { key: newKey(), id: null, title: label, url: '', children: [] };
    nav.tree.push(h);
  }
  return h;
}

function addGroup(g) {
  const missing = g.lists.filter((l) => !isLinked(l));
  if (!missing.length) return setStatus(`All ${g.label} lists are already in the navigation.`);
  headingFor(g.label).children.push(...missing.map(listNode));
  setStatus(`Added ${missing.length} ${g.label} list(s). Not saved yet.`);
  changed();
}

function addList(l, g) {
  (g.prefix ? headingFor(g.label).children : nav.tree).push(listNode(l));
  setStatus(`Added "${l.title}". Not saved yet.`);
  changed();
}

// ---------------------------------------------------------------- tree

function renderTree() {
  const build = (nodes, depth) =>
    nodes.map((n) => {
      const title = el('input', { type: 'text', class: 'nav-title', 'aria-label': 'Title', spellcheck: 'false', placeholder: 'Title' });
      const url = el('input', {
        type: 'text',
        class: 'nav-url',
        'aria-label': 'URL',
        spellcheck: 'false',
        placeholder: 'No link (heading only)',
        title: 'Server-relative (/sites/…) or https://… Leave empty for a heading without a link.',
      });
      title.value = n.title;
      url.value = n.url;
      title.addEventListener('input', () => {
        n.title = title.value;
        changed({ structure: false });
      });
      url.addEventListener('input', () => {
        n.url = url.value;
        changed({ structure: false });
      });
      url.addEventListener('change', () => (url.value = n.url = editorUrl(url.value)));
      const remove = el('button', { type: 'button', class: 'icon danger', text: '✕', 'aria-label': `Remove ${n.title}`, title: n.children.length ? 'Remove with everything below it' : 'Remove' });
      remove.addEventListener('click', () => removeNode(n.key));
      const handle = el('span', {
        class: 'drag-handle',
        draggable: 'true',
        tabindex: '0',
        role: 'button',
        title: 'Drag to move (or focus and press ↑/↓ to reorder, →/← to nest under the item above or move out)',
        'aria-label': `Move ${n.title}: drag, or press arrow keys`,
      }, '⠿');
      dragSource(handle, () => ({ kind: 'node', key: n.key }), (e) => e.dataTransfer.setDragImage(handle.closest('.nav-node'), 16, 16));
      handle.addEventListener('keydown', (e) => keyMove(e, n.key));
      const isNew = n.id == null;
      return el('li', { class: `nav-node${isNew ? ' new' : ''}`, 'data-key': n.key, 'data-depth': depth },
        el('div', { class: 'nav-row' },
          handle,
          el('div', { class: 'nav-fields' }, title, url),
          isNew ? el('span', { class: 'chip ok', text: 'new', title: 'Not saved yet' }) : null,
          remove),
        n.children.length ? el('ol', { class: 'nav-children' }, build(n.children, depth + 1)) : null);
    });
  ui.tree.replaceChildren(...build(nav.tree, 1));
  ui.treeEmpty.hidden = !nav.loaded || nav.tree.length > 0;
  for (const n of ui.tree.querySelectorAll('input, button')) n.disabled = nav.busy;
}

function removeNode(key) {
  const hit = findNode(nav.tree, key);
  if (!hit) return;
  hit.siblings.splice(hit.index, 1);
  changed();
}

/** Arrow keys on a focused handle: ↑/↓ reorder, → nest under the node above, ← move out. */
function keyMove(e, key) {
  const hit = findNode(nav.tree, key);
  if (!hit || nav.busy) return;
  const { node, siblings, index, depth } = hit;
  if (e.key === 'ArrowUp' || e.key === 'ArrowDown') {
    const to = index + (e.key === 'ArrowUp' ? -1 : 1);
    if (to < 0 || to >= siblings.length) return;
    siblings.splice(to, 0, ...siblings.splice(index, 1));
  } else if (e.key === 'ArrowRight') {
    const above = siblings[index - 1];
    if (!above || depth + subtreeHeight(node) > MAX_DEPTH) return setStatus(!above ? 'Nothing above to nest under.' : `SharePoint shows at most ${MAX_DEPTH} levels.`);
    siblings.splice(index, 1);
    above.children.push(node);
  } else if (e.key === 'ArrowLeft') {
    if (!hit.parent) return;
    const up = findNode(nav.tree, hit.parent.key);
    siblings.splice(index, 1);
    up.siblings.splice(up.index + 1, 0, node);
  } else return;
  e.preventDefault();
  changed();
  ui.tree.querySelector(`.nav-node[data-key="${key}"] .drag-handle`)?.focus();
}

// ---------------------------------------------------------------- drag and drop

const drag = { item: null, target: null }; // what is dragged, and where it would land

/** Make `node` start a drag of `item()`. */
function dragSource(node, item, extra) {
  node.addEventListener('dragstart', (e) => {
    if (nav.busy) return e.preventDefault();
    e.stopPropagation();
    drag.item = item();
    e.dataTransfer.effectAllowed = drag.item.kind === 'node' ? 'move' : 'copy';
    e.dataTransfer.setData('text/plain', drag.item.kind);
    if (extra) extra(e);
    (node.closest('.nav-node') || node).classList.add('dragging');
  });
  node.addEventListener('dragend', () => {
    (node.closest('.nav-node') || node).classList.remove('dragging');
    clearDropMarks();
    drag.item = null;
  });
}

/** Levels the dragged item takes up. */
function dragHeight(item) {
  if (item.kind === 'node') return subtreeHeight(findNode(nav.tree, item.key).node);
  return item.kind === 'group' ? 2 : 1;
}

/** Where a drop at clientY over `li` would land, or null if it isn't allowed. */
function dropPosition(li, clientY) {
  const key = li.dataset.key;
  const target = findNode(nav.tree, key);
  if (!target) return null;
  if (drag.item.kind === 'node') {
    const moving = findNode(nav.tree, drag.item.key);
    if (!moving || findNode([moving.node], key)) return null; // not onto itself or below itself
  }
  const height = dragHeight(drag.item);
  const box = li.querySelector('.nav-row').getBoundingClientRect();
  const y = (clientY - box.top) / box.height;
  const canNest = target.depth + height <= MAX_DEPTH;
  const canSide = target.depth + height - 1 <= MAX_DEPTH;
  let where = y < 0.3 ? 'before' : y > 0.7 ? 'after' : 'inside';
  if (where === 'inside' && !canNest) where = y < 0.5 ? 'before' : 'after';
  if (where !== 'inside' && !canSide) return null;
  return { key, where };
}

function clearDropMarks() {
  drag.target = null;
  ui.tree.classList.remove('drop-end');
  for (const n of ui.tree.querySelectorAll('.drop-before, .drop-after, .drop-inside')) n.classList.remove('drop-before', 'drop-after', 'drop-inside');
}

function onDragOver(e) {
  if (!drag.item || nav.busy || !nav.loaded) return;
  const row = e.target.closest('.nav-row');
  let pos;
  if (row) pos = dropPosition(row.closest('.nav-node'), e.clientY);
  else if (e.target === ui.tree || e.target === ui.treeEmpty) pos = { key: null, where: 'end' };
  else {
    if (drag.target) e.preventDefault(); // over a child list's padding: keep the last mark
    return;
  }
  e.preventDefault();
  e.dataTransfer.dropEffect = pos ? (drag.item.kind === 'node' ? 'move' : 'copy') : 'none';
  if (drag.target && pos && drag.target.key === pos.key && drag.target.where === pos.where) return;
  clearDropMarks();
  if (!pos) return;
  drag.target = pos;
  if (pos.where === 'end') ui.tree.classList.add('drop-end');
  else ui.tree.querySelector(`.nav-node[data-key="${pos.key}"]`).classList.add(`drop-${pos.where}`);
}

function onDrop(e) {
  if (!drag.item) return;
  e.preventDefault();
  const pos = drag.target;
  const item = drag.item;
  clearDropMarks();
  if (!pos) return;
  let node;
  if (item.kind === 'node') {
    const hit = findNode(nav.tree, item.key);
    hit.siblings.splice(hit.index, 1);
    node = hit.node;
  } else if (item.kind === 'list') {
    node = listNode(item.list);
  } else {
    const missing = item.group.lists.filter((l) => !isLinked(l));
    node = { key: newKey(), id: null, title: item.group.label, url: '', children: (missing.length ? missing : item.group.lists).map(listNode) };
  }
  if (pos.where === 'end') nav.tree.push(node);
  else {
    const t = findNode(nav.tree, pos.key);
    if (pos.where === 'inside') t.node.children.push(node);
    else t.siblings.splice(t.index + (pos.where === 'after' ? 1 : 0), 0, node);
  }
  changed();
}

for (const zone of [ui.tree, ui.treeEmpty]) {
  zone.addEventListener('dragover', onDragOver);
  zone.addEventListener('drop', onDrop);
}
ui.tree.addEventListener('dragleave', (e) => {
  if (drag.item && !ui.tree.contains(e.relatedTarget)) clearDropMarks();
});

// ---------------------------------------------------------------- saving

const titles = (items, max = 12) => {
  const list = items.map((x) => x.title || '(no title)');
  return el('ul', {}, list.slice(0, max).map((t) => el('li', { text: t })), list.length > max ? el('li', { text: `… and ${list.length - max} more` }) : null);
};

async function saveNav() {
  if (!isDirty() || nav.busy) return;
  const problems = validateTree(nav.tree);
  if (problems.length) {
    return ask({ title: 'Fix these first', body: el('ul', {}, problems.map((p) => el('li', { text: p }))), cancel: null });
  }
  const plan = planSave(nav.original, nav.tree);
  if (planIsEmpty(plan)) return setStatus('Nothing to save.');
  const added = plan.creates.filter((c) => !c.recreated);
  const recreated = plan.creates.filter((c) => c.recreated);
  const section = (label, items, body) => (items.length ? el('details', {}, el('summary', { text: `${label}: ${items.length}` }), body || titles(items)) : null);
  const ok = await ask({
    title: `Save the ${targetLabel()}?`,
    wide: true,
    body: [
      el('p', { class: 'target-title', text: nav.web.Title }),
      el('p', { class: 'target', text: nav.site }),
      el('p', { text: `${plan.deletes.length + plan.updates.length + plan.creates.length} change(s), sent as you (the signed-in user):` }),
      section('Add', added),
      section('Rename or change link', plan.updates, el('ul', {}, plan.updates.map((u) => el('li', { text: u.from.title === u.title ? `${u.title}: ${u.from.url || '(no link)'} → ${u.url || '(no link)'}` : `${u.from.title} → ${u.title}` })))),
      section('Remove', plan.removed),
      section('Recreate (moved)', recreated),
      recreated.length
        ? el('p', { class: 'hint', text: 'SharePoint can only add nodes at the end and cannot move them, so moved nodes and everything after them on their level are deleted and created again. Audience targeting on those nodes is lost.' })
        : null,
    ].filter(Boolean),
    ok: 'Save',
  });
  if (!ok) return setStatus('Save cancelled. Nothing was sent.');
  const backedUp = await backupCurrent();

  setBusy(true);
  ui.result.replaceChildren();
  ui.bar.value = 0;
  let countdown;
  const navigator = createNavigator({
    request: nav.request,
    onEvent: (e) => {
      clearInterval(countdown);
      if (e.type === 'step') {
        ui.bar.value = e.index / Math.max(1, e.total);
        setStatus(`${{ delete: 'Removing', update: 'Updating', create: 'Adding' }[e.action]} "${e.title}" (${e.index + 1} of ${e.total})…`);
      } else if (e.type === 'wait') {
        const until = Date.now() + e.ms;
        const tick = () => setStatus(`SharePoint is throttling. Waiting ${Math.max(0, Math.ceil((until - Date.now()) / 1000))} s (retry ${e.retry} of ${e.of})…`);
        tick();
        countdown = setInterval(tick, 500);
      }
    },
  });
  try {
    const results = await navigator.save(nav.site, nav.target, plan);
    clearInterval(countdown);
    const failed = results.filter((r) => !r.ok);
    await loadNav(); // show what SharePoint has now
    await refreshBackups();
    setStatus((backedUp ? '' : 'The previous navigation could not be backed up. ') + (failed.length
      ? `Saved with ${failed.length} failure(s) of ${results.length} change(s). The tree shows the ${targetLabel()} as it is now.`
      : `Saved ${results.length} change(s). Reload the SharePoint page to see the new ${targetLabel()}. The previous one can be brought back with Restore….`));
    ui.result.replaceChildren(
      ...(failed.length ? [el('details', { class: 'check', open: true },
            el('summary', {}, el('span', { class: 'chip fail', text: String(failed.length) }), 'Failed'),
            el('div', { class: 'check-body' }, el('div', { class: 'msg fail' }, el('ul', {}, failed.map((f) => el('li', { text: `${f.action} "${f.title}": ${f.error}` }))))))] : [])
    );
  } catch (e) {
    clearInterval(countdown);
    setStatus(`Save stopped: ${e.message}. Click Reload to see the ${targetLabel()} as it is now.`);
  } finally {
    setBusy(false);
  }
}

// ---------------------------------------------------------------- wiring

document.addEventListener('sp-utility:tab-shown', (e) => {
  if (e.detail.id === 'nav' && !nav.loaded && !nav.busy) loadAll();
});
ui.reload.addEventListener('click', async () => {
  if (await confirmDiscard('Reloading')) loadAll();
});
ui.switchTarget.addEventListener('click', async () => {
  const next = nav.target === 'topnav' ? 'quicklaunch' : 'topnav';
  if (!nav.loaded || !(await confirmDiscard(`Switching to the ${NAV_TARGETS[next].label}`))) return;
  nav.target = next;
  setBusy(true);
  try {
    await loadNav();
    await refreshBackups();
  } catch (e) {
    nav.loaded = false;
    setStatus(`Could not read the ${targetLabel()}: ${e.message}`);
    changed();
  } finally {
    setBusy(false);
  }
});
ui.filter.addEventListener('input', renderGroups);
ui.addHeading.addEventListener('click', () => {
  nav.tree.push({ key: newKey(), id: null, title: '', url: '', children: [] });
  changed();
  ui.tree.lastElementChild?.querySelector('.nav-title')?.focus();
});
ui.addLink.addEventListener('click', () => {
  nav.tree.push({ key: newKey(), id: null, title: '', url: '/', children: [] });
  changed();
  ui.tree.lastElementChild?.querySelector('.nav-title')?.focus();
});
ui.copy.addEventListener('click', async () => {
  const text = exportNavigation(nav.tree, {
    target: nav.target,
    site: nav.site,
    origin: nav.origin,
    webServerRelativeUrl: nav.web.ServerRelativeUrl,
  });
  const count = (list) => list.reduce((n, x) => n + 1 + count(x.children), 0);
  try {
    await navigator.clipboard.writeText(text);
    setStatus(`Copied ${count(nav.tree)} item(s) of the ${targetLabel()}. Open the other site, go to Navigation and click Paste….`);
  } catch {
    // no clipboard access: show the text to copy by hand
    const box = el('textarea', { class: 'nav-paste-box', readonly: true, rows: 12, spellcheck: 'false' });
    box.value = text;
    await ask({ title: 'Copy the navigation', body: [el('p', { text: 'The browser did not allow copying. Select all and copy:' }), box], ok: 'Close', cancel: null, wide: true });
  }
});

ui.paste.addEventListener('click', async () => {
  const box = el('textarea', { class: 'nav-paste-box', rows: 12, spellcheck: 'false', placeholder: 'Paste the copied navigation here' });
  const ok = await ask({
    title: 'Paste a navigation',
    body: [el('p', { text: `The items are added to the editor below; nothing is written to ${nav.site} until you Save.` }), box],
    ok: 'Paste',
    wide: true,
  });
  if (!ok) return;
  let imported;
  try {
    imported = parseNavigationExport(box.value, { site: nav.site, webServerRelativeUrl: nav.web.ServerRelativeUrl });
  } catch (e) {
    setStatus(`Could not paste: ${e.message}`);
    return;
  }
  if (!imported.count) {
    setStatus('Nothing to paste: the copied navigation is empty.');
    return;
  }
  let replace = false;
  if (nav.tree.length) {
    replace = await ask({
      title: 'Paste a navigation',
      body: `Replace the ${nav.tree.length} top-level item(s) in the editor, or add the ${imported.nodes.length} pasted item(s) at the end?`,
      ok: 'Replace',
      cancel: 'Add to end',
    });
  }
  nav.tree = replace ? imported.nodes : [...nav.tree, ...imported.nodes];
  const from = imported.source ? ` from ${imported.source}` : '';
  const other = imported.target && imported.target !== nav.target ? ` It was the ${NAV_TARGETS[imported.target].label}, you are editing the ${targetLabel()}.` : '';
  setStatus(`Pasted ${imported.count} item(s)${from}.${other} Check the tree, then Save.`);
  changed();
});

ui.restore.addEventListener('click', async () => {
  await refreshBackups();
  if (!nav.backups.length) return setStatus('There is no earlier navigation stored for this site.');
  const when = (iso) => new Date(iso).toLocaleString();
  const pick = el('select', { 'aria-label': 'Earlier navigation' },
    nav.backups.map((b, i) => el('option', { value: String(i), text: `${when(b.at)} · ${countNodes(b.tree)} item(s)` })));
  const ok = await ask({
    title: `Restore an earlier ${targetLabel()}`,
    body: [
      el('p', { text: `These are the ${targetLabel()} versions ${nav.site} had before a Save from this extension (last ${nav.backups.length}, newest first). The chosen one replaces the tree in the editor; nothing is written until you Save.` }),
      pick,
    ],
    ok: 'Load into editor',
  });
  if (!ok) return;
  const entry = nav.backups[Number(pick.value)];
  nav.tree = treeFromSnapshot(entry.tree);
  setStatus(`Loaded the ${targetLabel()} from ${when(entry.at)} into the editor. Check the tree, then Save.`);
  changed();
});

ui.discard.addEventListener('click', async () => {
  const ok = await ask({ title: 'Discard your changes?', body: `The tree goes back to the ${targetLabel()} as it is on the site.`, ok: 'Discard', cancel: 'Keep editing' });
  if (!ok) return;
  nav.tree = cloneTree(nav.original);
  setStatus('Changes discarded.');
  changed();
});
ui.save.addEventListener('click', saveNav);
