// "Compare" tab: the differences between the lists and columns of two open
// SharePoint sites. Source: another open SharePoint tab. Target: the site in
// the active tab. Read only: every request is a GET, run in its own tab.

import { el, activeTab, tabRequester, siteOfTab } from './shared.js';
import { isSharePointUrl } from './lib/batch.js';
import { createExporter } from './lib/exporter.js';
import {
  compareSites, countStatuses, columnSummary, listStatus, filterView, diffRows, diffCsv, diffFileName, STATUSES, STATUS_LABELS,
} from './lib/diff.js';

const $ = (id) => document.getElementById(id);

const ui = {
  reload: $('cmp-reload'),
  source: $('cmp-source'),
  target: $('cmp-target'),
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
};

const cmp = {
  shown: false, // the Compare tab is the visible tab
  busy: false,
  loading: null, // the running loadSites(), so calls don't overlap
  target: null, // { tabId, site, title }
  sources: [], // [{ tabId, site, title }]
  result: null, // { lists, source, target }
  statuses: new Set(['missing', 'changed', 'only']),
};

const CHIP = { missing: 'add', changed: 'warn', only: 'muted', same: 'ok', error: 'fail' };
const SHORT = { missing: 'Missing', changed: 'Changed', only: 'Only in target', same: 'Same', error: 'Not compared' };
const COUNT_TEXT = { missing: 'missing in target', changed: 'changed', only: 'only in target', same: 'same' };
const plural = (n, one, many = `${one}s`) => `${n} ${n === 1 ? one : many}`;
const bare = (url) => url.replace(/^https:\/\//, '');

function setBusy(busy) {
  cmp.busy = busy;
  for (const n of [ui.reload, ui.source, ui.system]) n.disabled = busy;
  updateRunButton();
}

function updateRunButton() {
  ui.run.disabled = cmp.busy || !cmp.target || !ui.source.value;
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

function loadSites() {
  if (!cmp.loading) cmp.loading = readSites().finally(() => { cmp.loading = null; });
  return cmp.loading;
}

async function readSites() {
  const keep = ui.source.value;
  ui.siteMsg.hidden = true;
  cmp.target = null;
  updateRunButton();

  const tab = await activeTab();
  if (!tab || !tab.url || !isSharePointUrl(tab.url)) {
    ui.target.className = 'cmp-target error';
    ui.target.textContent = 'The active tab is not a SharePoint page. Open any page of the target site in it.';
    renderSources([]);
    return;
  }
  const site = await siteOfTab(tab);
  const title = (await webTitle(tabRequester(tab.id, new URL(site).origin), site)) || tab.title || site;
  cmp.target = { tabId: tab.id, site, title };
  ui.target.className = 'cmp-target';
  ui.target.replaceChildren(el('strong', { text: title }), el('span', { class: 'exp-url', text: site }));

  // Every other open SharePoint tab, one entry per site, without the target site.
  const tabs = (await chrome.tabs.query({})).filter((t) => t.id !== tab.id && t.url && isSharePointUrl(t.url));
  const sources = [];
  for (const t of tabs) {
    let s;
    try {
      s = await siteOfTab(t);
    } catch {
      continue; // the tab can't be read (discarded, still loading)
    }
    if (!s || s === site || sources.some((x) => x.site === s)) continue;
    const name = (await webTitle(tabRequester(t.id, new URL(s).origin), s)) || t.title || s;
    sources.push({ tabId: t.id, site: s, title: name });
  }
  sources.sort((a, b) => a.title.localeCompare(b.title, undefined, { sensitivity: 'base' }));
  renderSources(sources, keep);
}

function renderSources(sources, keep) {
  cmp.sources = sources;
  if (!sources.length) {
    ui.source.replaceChildren(el('option', { value: '', text: 'No other SharePoint tab is open' }));
    if (cmp.target) {
      ui.siteMsg.textContent = 'Open the source site (e.g. dev) in another tab, then click Reload tabs.';
      ui.siteMsg.hidden = false;
    }
  } else {
    ui.source.replaceChildren(...sources.map((s) => el('option', { value: String(s.tabId), text: `${s.title} — ${bare(s.site)}` })));
    if (keep && sources.some((s) => String(s.tabId) === keep)) ui.source.value = keep;
  }
  updateRunButton();
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
    const ids = loaded.lists.filter((l) => ui.system.checked || !l.system).map((l) => l.id);
    const data = await exporter.collect(side.site, loaded, ids, { permissions: false, views: false, contentTypes: false });
    return { web: loaded.web, allLists: loaded.lists, lists: data.lists };
  } finally {
    clearInterval(countdown);
  }
}

async function runCompare() {
  if (cmp.busy) return;
  const source = cmp.sources.find((s) => String(s.tabId) === ui.source.value);
  const target = cmp.target;
  if (!source || !target) return;
  setBusy(true);
  cmp.result = null;
  ui.result.hidden = true;
  ui.bar.value = 0;
  try {
    if (source.site === target.site) throw new Error('Source and target are the same site.');
    const src = await readSide(source, 'source', (f, text) => { ui.bar.value = f / 2; setStatus(text); });
    const tgt = await readSide(target, 'target', (f, text) => { ui.bar.value = 0.5 + f / 2; setStatus(text); });
    ui.bar.value = 1;
    cmp.result = { lists: compareSites(src, tgt), source: { ...source, title: src.web.Title || source.title }, target: { ...target, title: tgt.web.Title || target.title } };
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

function listBody({ list, columns, showSettings }) {
  const body = [];
  if (list.status === 'missing') body.push(el('p', { class: 'hint', text: 'The whole list is missing in the target.' }));
  if (list.status === 'only') body.push(el('p', { class: 'hint', text: 'This list exists only in the target.' }));
  if (list.status === 'error') body.push(el('div', { class: 'msg fail', text: `Not compared: ${list.error}.` }));
  if (list.note) body.push(el('div', { class: 'msg info cmp-note', text: list.note }));
  if (showSettings && list.diffs.length) {
    body.push(el('p', { class: 'summary-text cmp-settings-h', text: 'List settings' }), el('ul', { class: 'cmp-diff cmp-props' }, list.diffs.map(diffItem)));
  }
  if (columns.length) {
    body.push(el('table', { class: 'cmp-table' },
      el('thead', {}, el('tr', {}, el('th', { text: 'Column' }), el('th', { class: 'kind', text: 'Type' }), el('th', { class: 'st', text: 'Status' }))),
      el('tbody', {}, columns.map((c) => el('tr', {},
        el('td', {},
          c.title, el('span', { class: 'iname', text: c.name }),
          c.diffs && c.diffs.length ? el('ul', { class: 'cmp-diff' }, c.diffs.map(diffItem)) : null,
          c.note ? el('ul', { class: 'cmp-diff' }, el('li', { text: c.note })) : null),
        el('td', { class: 'kind', text: c.type }),
        el('td', { class: 'st' }, el('span', { class: `chip ${CHIP[c.status]}`, text: SHORT[c.status] })))))));
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
ui.source.addEventListener('change', updateRunButton);
ui.run.addEventListener('click', runCompare);
ui.filter.addEventListener('input', () => cmp.result && renderLists());
ui.custom.addEventListener('change', () => cmp.result && (renderCounts(), renderLists()));
ui.download.addEventListener('click', download);
