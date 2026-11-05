// "Compare" tab: the differences between the lists and columns of two open
// SharePoint sites. Source and target: any open SharePoint tab (the target
// defaults to the active tab). Read only: every request is a GET, run in its own tab.

import { el, activeTab, tabRequester, siteOfTab } from './shared.js';
import { isSharePointUrl } from './lib/batch.js';
import { webRelative } from './lib/provision.js';
import { createExporter } from './lib/exporter.js';
import {
  compareSites, countStatuses, columnSummary, listStatus, filterView, diffRows, diffCsv, diffFileName, STATUSES, STATUS_LABELS,
} from './lib/diff.js';
import { settingsStage, settingsTransferable, columnStage, columnTransferable, listStages } from './lib/transfer.js';

const $ = (id) => document.getElementById(id);

const ui = {
  reload: $('cmp-reload'),
  source: $('cmp-source'),
  target: $('cmp-target'),
  addSource: $('cmp-add-source'),
  addTarget: $('cmp-add-target'),
  siteMsg: $('cmp-site-msg'),
  system: $('cmp-system'),
  run: $('cmp-run'),
  progress: $('cmp-progress'),
  bar: $('cmp-bar'),
  status: $('cmp-status'),
  result: $('cmp-result'),
  counts: $('cmp-counts'),
  filter: $('cmp-filter'),
  custom: $('cmp-custom'),
  lists: $('cmp-lists'),
  empty: $('cmp-empty'),
  download: $('cmp-download'),
  scopeSum: $('cmp-scope-sum'),
  scopeLoad: $('cmp-scope-load'),
  scopeFilter: $('cmp-scope-filter'),
  scopeAll: $('cmp-scope-all'),
  scopeNone: $('cmp-scope-none'),
  scopeRows: $('cmp-scope-rows'),
};

const cmp = {
  shown: false, // the Compare tab is the visible tab
  busy: false,
  loading: null, // the running loadSites(), so calls don't overlap
  sites: [], // [{ tabId, site, title, manual }]: open SharePoint sites and the ones added by hand
  result: null, // { lists, source, target }
  scope: null, // null = all lists; else { lists: [{ key, title, rel, system }], picked: Set of keys } of the chosen source
  statuses: new Set(['missing', 'changed', 'only']),
};

const CHIP = { missing: 'add', changed: 'warn', only: 'muted', same: 'ok', error: 'fail' };
const SHORT = { missing: 'Missing', changed: 'Changed', only: 'Only in target', same: 'Same', error: 'Not compared' };
const COUNT_TEXT = { missing: 'missing in target', changed: 'changed', only: 'only in target', same: 'same' };
const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;
const bare = (url) => url.replace(/^https:\/\//, '');

function setBusy(busy) {
  cmp.busy = busy;
  for (const n of [ui.scopeLoad, ui.reload, ui.source, ui.target, ui.addSource, ui.addTarget, ui.system]) n.disabled = busy;
  updateRunButton();
}

function updateRunButton() {
  ui.run.disabled = cmp.busy || !ui.target.value || !ui.source.value || ui.source.value === ui.target.value;
}

function setStatus(text) {
  ui.progress.hidden = false;
  ui.status.textContent = text;
}

/** The web title, read in the site's own tab; null if it can't be read. */
async function webTitle(request, site) {
  try {
    const res = await request({ url: `${site}/_api/web?$select=Title`, method: 'GET', headers: { Accept: 'application/json;odata=nometadata' } });
    return res.status === 200 ? JSON.parse(res.text).Title : null;
  } catch {
    return null;
  }
}

// ---------------------------------------------------------------- sites

const withTimeout = (promise, ms) => Promise.race([promise, new Promise((_, reject) => setTimeout(() => reject(new Error('timed out')), ms))]);

function loadSites() {
  if (!cmp.loading) {
    cmp.loading = readSites()
      .catch((e) => {
        ui.siteMsg.textContent = `Could not read the open tabs: ${e.message}`;
        ui.siteMsg.hidden = false;
      })
      .finally(() => { cmp.loading = null; });
  }
  return cmp.loading;
}

/** Read the site of one tab: { tabId, site, title }, or null if it can't be read (discarded, still loading). */
const tabCache = new Map(); // tabId -> { url, entry }: a tab is only read again after it navigated

async function readTab(t) {
  const hit = tabCache.get(t.id);
  if (hit && hit.url === t.url) return { ...hit.entry };
  const entry = await readTabNow(t);
  if (entry) tabCache.set(t.id, { url: t.url, entry });
  return entry;
}

async function readTabNow(t) {
  try {
    const site = await withTimeout(siteOfTab(t), 5000);
    if (!site) return null;
    const title = (await withTimeout(webTitle(tabRequester(t.id, new URL(site).origin), site), 5000).catch(() => null)) || t.title || site;
    return { tabId: t.id, site, title };
  } catch {
    return null;
  }
}

const byTitle = (a, b) => a.title.localeCompare(b.title, undefined, { sensitivity: 'base' });

/** Add or refresh a site in the lists. Sites added by hand stay when their tab is closed. */
function remember(entry, manual) {
  const known = cmp.sites.find((s) => s.site === entry.site);
  if (known) Object.assign(known, entry, { manual: known.manual || manual });
  else cmp.sites.push({ ...entry, manual });
  cmp.sites.sort(byTitle);
}

/** Every open SharePoint site goes into both dropdowns. Nothing is selected for you. */
async function readSites() {
  const tabs = (await chrome.tabs.query({})).filter((t) => t.url && isSharePointUrl(t.url));
  for (const id of [...tabCache.keys()]) if (!tabs.some((t) => t.id === id)) tabCache.delete(id);
  const fresh = tabs.some((t) => (tabCache.get(t.id) || {}).url !== t.url);
  if (fresh) showLoading('Reading the open SharePoint tabs…');
  else ui.siteMsg.hidden = true;
  const found = (await Promise.all(tabs.map(readTab))).filter(Boolean);
  ui.siteMsg.classList.remove('sch-loading');
  ui.siteMsg.hidden = true;
  for (const s of cmp.sites) s.tabId = null; // set again below for the sites still open
  for (const f of found) remember(f, false);
  cmp.sites = cmp.sites.filter((s) => s.manual || s.tabId != null);
  renderSites();
  if (!cmp.sites.length) {
    ui.siteMsg.textContent = 'No SharePoint site found. Open a site in the browser, then click "Add active tab".';
    ui.siteMsg.hidden = false;
  }
}

function showLoading(text) {
  ui.siteMsg.textContent = text;
  ui.siteMsg.classList.add('sch-loading');
  ui.siteMsg.hidden = false;
}

/** The site in the active tab goes into one dropdown and is selected there. */
async function addActive(select) {
  ui.siteMsg.classList.remove('sch-loading');
  ui.siteMsg.hidden = true;
  const tab = await activeTab();
  if (tab && tab.url && isSharePointUrl(tab.url) && !tabCache.has(tab.id)) showLoading('Reading the active tab…');
  const entry = tab && tab.url && isSharePointUrl(tab.url) ? await readTab(tab) : null;
  ui.siteMsg.classList.remove('sch-loading');
  ui.siteMsg.hidden = true;
  if (!entry) {
    ui.siteMsg.textContent = "The active tab is not a SharePoint page (or can't be read). Switch to a page of the site in the browser first.";
    ui.siteMsg.hidden = false;
    return;
  }
  remember(entry, true);
  renderSites();
  select.value = entry.site;
  updateRunButton();
  if (select === ui.source) await loadScope();
}

function renderSites() {
  const keep = { source: ui.source.value, target: ui.target.value };
  for (const [name, select] of [['source', ui.source], ['target', ui.target]]) {
    select.replaceChildren(
      el('option', { value: '', text: `Choose the ${name} site…` }),
      ...cmp.sites.map((s) => el('option', { value: s.site, text: `${s.title} — ${bare(s.site)}` })));
    if (cmp.sites.some((s) => s.site === keep[name])) select.value = keep[name];
  }
  updateRunButton();
}

// ---------------------------------------------------------------- scope

const listKey = (l, web) => webRelative(l.url, web.ServerRelativeUrl).toLowerCase();

/** Back to "all lists" (the chosen sites changed). */
function clearScope() {
  cmp.scope = null;
  ui.scopeRows.replaceChildren();
  ui.scopeSum.textContent = 'Lists: all';
}

/** Read the lists of the chosen source site for the picker (all ticked at first). */
async function loadScope() {
  if (cmp.busy) return;
  await loadSites();
  const source = cmp.sites.find((s) => s.site === ui.source.value);
  if (!source || source.tabId == null) {
    clearScope();
    return;
  }
  const sides = [source];
  setBusy(true);
  ui.scopeSum.textContent = 'Lists: loading…';
  showLoading('Loading the lists of the source…');
  try {
    const byKey = new Map();
    for (const side of sides) {
      const { web, lists } = await createExporter({ request: tabRequester(side.tabId, new URL(side.site).origin) }).loadLists(side.site);
      for (const l of lists) {
        const key = listKey(l, web);
        const known = byKey.get(key);
        if (known) known.system = known.system && l.system;
        else byKey.set(key, { key, title: l.title, rel: webRelative(l.url, web.ServerRelativeUrl), system: !!l.system });
      }
    }
    const lists = [...byKey.values()].sort((a, b) => a.title.localeCompare(b.title, undefined, { sensitivity: 'base' }));
    cmp.scope = { lists, picked: new Set(lists.filter((l) => !l.system || ui.system.checked).map((l) => l.key)) };
    ui.scopeFilter.value = '';
    renderScope();
  } catch (e) {
    clearScope();
    ui.siteMsg.textContent = `Could not read the lists: ${e.message}`;
    ui.siteMsg.hidden = false;
  } finally {
    ui.siteMsg.classList.remove('sch-loading');
    if (ui.siteMsg.textContent.startsWith('Loading')) ui.siteMsg.hidden = true;
    setBusy(false);
  }
}

const scopeShown = () => {
  const q = ui.scopeFilter.value.trim().toLowerCase();
  return cmp.scope.lists.filter((l) => (ui.system.checked || !l.system || cmp.scope.picked.has(l.key)) && (!q || `${l.title} ${l.rel}`.toLowerCase().includes(q)));
};

function setPicked(on) {
  if (!cmp.scope) return;
  for (const l of scopeShown()) (on ? cmp.scope.picked.add(l.key) : cmp.scope.picked.delete(l.key));
  renderScope();
}

function renderScope() {
  if (!cmp.scope) return;
  ui.scopeRows.replaceChildren(...scopeShown().map((l) => {
    const box = el('input', { type: 'checkbox' });
    box.checked = cmp.scope.picked.has(l.key);
    box.addEventListener('change', () => {
      if (box.checked) cmp.scope.picked.add(l.key);
      else cmp.scope.picked.delete(l.key);
      renderScope();
    });
    return el('label', {}, box, l.title, el('span', { class: 'exp-url', text: l.rel }));
  }));
  ui.scopeSum.textContent = `Lists: ${cmp.scope.picked.size} of ${cmp.scope.lists.length} chosen`;
}

// ---------------------------------------------------------------- compare

/** Read one side: lists and their columns. `progress(fraction, text)` reports. */
async function readSide(side, label, progress) {
  let countdown;
  const exporter = createExporter({
    request: tabRequester(side.tabId, new URL(side.site).origin),
    onEvent: (e) => {
      clearInterval(countdown);
      if (e.type === 'list') progress(e.index / e.total, `Reading ${label}: "${e.title}" (${e.index + 1} of ${e.total})…`);
      else if (e.type === 'wait') {
        const until = Date.now() + e.ms;
        const tick = () => setStatus(`SharePoint is throttling the ${label}. Waiting ${Math.max(0, Math.ceil((until - Date.now()) / 1000))} s (retry ${e.retry} of ${e.of})…`);
        tick();
        countdown = setInterval(tick, 500);
      }
    },
  });
  try {
    progress(0, `Reading the lists of the ${label}…`);
    const loaded = await exporter.loadLists(side.site);
    const ids = loaded.lists
      .filter((l) => (cmp.scope ? cmp.scope.picked.has(listKey(l, loaded.web)) : ui.system.checked || !l.system))
      .map((l) => l.id);
    const data = await exporter.collect(side.site, loaded, ids, { permissions: false, views: false, contentTypes: false });
    return { web: loaded.web, allLists: loaded.lists, lists: data.lists };
  } finally {
    clearInterval(countdown);
  }
}

async function runCompare() {
  if (cmp.busy) return;
  setBusy(true);
  cmp.result = null;
  ui.result.hidden = true;
  ui.bar.value = 0;
  try {
    await loadSites(); // the tabs may have changed since the sites were picked
    if (cmp.scope && !cmp.scope.picked.size) throw new Error('No list is chosen. Tick at least one.');
    const source = cmp.sites.find((s) => s.site === ui.source.value);
    const target = cmp.sites.find((s) => s.site === ui.target.value);
    if (!source || !target) throw new Error('Choose a source and a target site.');
    if (source.site === target.site) throw new Error('Source and target are the same site.');
    for (const side of [source, target]) if (side.tabId == null) throw new Error(`${side.title} is not open in a tab. Open it, then run again.`);
    const src = await readSide(source, 'source', (f, text) => { ui.bar.value = f / 2; setStatus(text); });
    const tgt = await readSide(target, 'target', (f, text) => { ui.bar.value = 0.5 + f / 2; setStatus(text); });
    ui.bar.value = 1;
    cmp.result = { lists: compareSites(src, tgt), ctx: { web: src.web, allLists: src.allLists, targetSite: target.site }, source: { ...source, title: src.web.Title || source.title }, target: { ...target, title: tgt.web.Title || target.title } };
    setStatus(`Compared ${plural(src.lists.length, 'list')} from ${cmp.result.source.title} with ${cmp.result.target.title}.`);
    ui.result.hidden = false;
    renderResult();
  } catch (e) {
    setStatus(`Compare failed: ${e.message}`);
  } finally {
    setBusy(false);
  }
}

// ---------------------------------------------------------------- results

const opts = () => ({ customOnly: ui.custom.checked });

function renderCounts() {
  const n = countStatuses(cmp.result.lists, opts());
  const chips = STATUSES.map((s) => {
    const box = el('input', { type: 'checkbox', 'aria-label': `Show ${STATUS_LABELS[s].toLowerCase()}` });
    box.checked = cmp.statuses.has(s);
    box.addEventListener('change', () => {
      if (box.checked) cmp.statuses.add(s);
      else cmp.statuses.delete(s);
      renderLists();
    });
    return el('label', {}, box, el('span', { class: `chip ${CHIP[s]}`, text: `${n[s]} ${COUNT_TEXT[s]}` }));
  });
  if (n.error) chips.push(el('span', { class: 'chip fail', text: `${n.error} not compared` }));
  ui.counts.replaceChildren(...chips);
}

/** "Required: No → Yes", choices as + / −; the target's value first, then what the source has. */
function diffItem(d) {
  const shown = (v) => (v === '' ? '(empty)' : v);
  const parts = [`${d.prop}: `];
  if (d.added) {
    const changes = [
      ...d.added.map((v) => el('span', { class: 'to', text: `+ "${v}"` })),
      ...d.removed.map((v) => el('span', { class: 'from', text: `− "${v}"` })),
    ];
    if (!changes.length) changes.push(el('span', { class: 'to', text: 'order changed' }));
    changes.forEach((c, i) => parts.push(...(i ? [', ', c] : [c])));
  } else {
    parts.push(el('span', { class: 'from', text: shown(d.target) }), ' → ', el('span', { class: 'to', text: shown(d.source) }));
  }
  if (d.blocked) parts.push(el('span', { class: 'chip fail', text: "can't be changed in place" }));
  return el('li', {}, ...parts);
}

function listSubtitle(list) {
  if (list.status === 'error') return list.error;
  if (list.status === 'missing' || list.status === 'only') return plural(list.columnCount, 'column');
  const cols = columnSummary(list, opts());
  const settings = list.diffs.length ? 'settings changed' : '';
  return [settings, cols].filter(Boolean).join(' · ') || 'no differences';
}

/** A button that hands the stages from `build()` to the Batch Runner (appended to its stages). */
function sendButton(what, build) {
  const b = el('button', { type: 'button', class: 'secondary small', text: 'To Batch Runner', title: `Send ${what} to the Batch Runner as a stage. Nothing runs yet.` });
  b.addEventListener('click', () => {
    const stages = [build()].flat().filter(Boolean);
    if (!stages.length) {
      setStatus(`Nothing to send for ${what}.`);
      return;
    }
    document.getElementById('tab-batch').click();
    document.dispatchEvent(new CustomEvent('sp-utility:load-stages', {
      detail: { stages, append: true, label: `${what} from Compare` },
    }));
  });
  return b;
}

function listBody({ list, columns, showSettings }) {
  const body = [];
  if (list.status === 'missing') {
    body.push(el('p', { class: 'hint', text: 'The whole list is missing in the target.' }));
    if (list.item && list.item.list) body.push(el('div', { class: 'row' }, sendButton(`creating "${list.title}"`, () => listStages(list, cmp.result.ctx))));
  }
  if (list.status === 'only') body.push(el('p', { class: 'hint', text: 'This list exists only in the target.' }));
  if (list.status === 'error') body.push(el('div', { class: 'msg fail', text: `Not compared: ${list.error}.` }));
  if (list.note) body.push(el('div', { class: 'msg info cmp-note', text: list.note }));
  if (showSettings && list.diffs.length) {
    body.push(el('p', { class: 'summary-text cmp-settings-h', text: 'List settings' }), el('ul', { class: 'cmp-diff cmp-props' }, list.diffs.map(diffItem)));
    if (settingsTransferable(list)) body.push(el('div', { class: 'row' }, sendButton(`the settings of "${list.title}"`, () => settingsStage(list, cmp.result.ctx.targetSite))));
  }
  if (columns.length) {
    body.push(el('table', { class: 'cmp-table' },
      el('thead', {}, el('tr', {}, el('th', { text: 'Column' }), el('th', { class: 'kind', text: 'Type' }), el('th', { class: 'st', text: 'Status' })))),
      el('tbody', {}, columns.map((c) => el('tr', {},
        el('td', {},
          c.title, el('span', { class: 'iname', text: c.name }),
          c.diffs && c.diffs.length ? el('ul', { class: 'cmp-diff' }, c.diffs.map(diffItem)) : null,
          c.note ? el('ul', { class: 'cmp-diff' }, el('li', { text: c.note })) : null,
          columnTransferable(c) ? sendButton(`column "${c.title}"`, () => columnStage(list, c, cmp.result.ctx)) : null),
        el('td', { class: 'kind', text: c.type }),
        el('td', { class: 'st' }, el('span', { class: `chip ${CHIP[c.status]}`, text: SHORT[c.status] }))))));
  }
  return el('div', { class: 'check-body' }, body);
}

/** Which lists are expanded, so a filter change doesn't collapse them. */
const openLists = new Set();

function currentView() {
  return filterView(cmp.result.lists, { statuses: [...cmp.statuses], query: ui.filter.value, ...opts() });
}

function renderLists() {
  const view = currentView();
  ui.lists.replaceChildren(...view.map((v) => {
    const { list } = v;
    const shownStatus = listStatus(list, opts());
    const d = el('details', { class: 'check cmp-list' },
      el('summary', {},
        el('span', { class: `chip ${CHIP[shownStatus]}`, text: SHORT[shownStatus] }),
        el('span', { class: 't' }, list.title, el('span', { class: 'exp-url', text: `${list.rel} · ${listSubtitle(list)}` }))),
      listBody(v));
    d.open = openLists.has(list.key);
    d.addEventListener('toggle', () => (d.open ? openLists.add(list.key) : openLists.delete(list.key)));
    return d;
  }));
  ui.empty.hidden = view.length > 0;
  ui.download.disabled = view.length === 0;
}

function renderResult() {
  openLists.clear();
  renderCounts();
  renderLists();
}

function download() {
  if (!cmp.result) return;
  const text = diffCsv(diffRows(currentView()));
  const url = URL.createObjectURL(new Blob([text], { type: 'text/csv;charset=utf-8' }));
  const a = el('a', { href: url, download: diffFileName(cmp.result.target.title) });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ---------------------------------------------------------------- events

document.addEventListener('sp-utility:tab-shown', (e) => {
  cmp.shown = e.detail.id === 'cmp';
  if (cmp.shown && !cmp.busy) loadSites();
});

// The target follows the active tab; refresh while the Compare tab is visible.
chrome.tabs.onActivated.addListener(() => {
  if (cmp.shown && !cmp.busy) loadSites();
});
chrome.tabs.onUpdated.addListener((_id, info) => {
  if (cmp.shown && !cmp.busy && info.status === 'complete') loadSites();
});
chrome.tabs.onRemoved.addListener(() => {
  if (cmp.shown && !cmp.busy) loadSites();
});

ui.reload.addEventListener('click', loadSites);
ui.source.addEventListener('change', () => (clearScope(), updateRunButton(), loadScope()));
ui.target.addEventListener('change', updateRunButton);
ui.scopeLoad.addEventListener('click', loadScope);
ui.scopeFilter.addEventListener('input', renderScope);
ui.system.addEventListener('change', () => cmp.scope && renderScope());
ui.scopeAll.addEventListener('click', () => setPicked(true));
ui.scopeNone.addEventListener('click', () => setPicked(false));
ui.addSource.addEventListener('click', () => addActive(ui.source));
ui.addTarget.addEventListener('click', () => addActive(ui.target));
ui.run.addEventListener('click', runCompare);
ui.filter.addEventListener('input', () => cmp.result && renderLists());
ui.custom.addEventListener('change', () => cmp.result && (renderCounts(), renderLists()));
ui.download.addEventListener('click', download);
