// "Export Architecture" tab, plus the tab switching for the side panel.

import { el, activeTab, tabRequester, siteOfTab } from './shared.js';
import { isSharePointUrl } from './lib/batch.js';
import { createExporter } from './lib/exporter.js';
import { toCsv, csvFileName } from './lib/export.js';
import { buildProvisioning, batchFileName, splitStageFile as splitForRunner } from './lib/provision.js';

const $ = (id) => document.getElementById(id);
const SEPARATOR = ';'; // German Excel opens semicolon CSV directly in columns

$('version').textContent = `v${chrome.runtime.getManifest().version}`;

// ---------------------------------------------------------------- tabs

const tabs = [
  { id: 'batch', tab: $('tab-batch'), view: $('view-batch') },
  { id: 'export', tab: $('tab-export'), view: $('view-export') },
  { id: 'nav', tab: $('tab-nav'), view: $('view-nav') },
  { id: 'perm', tab: $('tab-perm'), view: $('view-perm') },
  { id: 'cmp', tab: $('tab-cmp'), view: $('view-cmp') },
  { id: 'sch', tab: $('tab-sch'), view: $('view-sch') },
];

function selectTab(i) {
  tabs.forEach((t, k) => {
    t.tab.setAttribute('aria-selected', String(k === i));
    t.tab.tabIndex = k === i ? 0 : -1;
    t.view.hidden = k !== i;
  });
  if (tabs[i].id === 'export' && !exp.data && !exp.busy) loadLists();
  document.dispatchEvent(new CustomEvent('sp-utility:tab-shown', { detail: { id: tabs[i].id } }));
}

tabs.forEach((t, i) => {
  t.tab.addEventListener('click', () => selectTab(i));
  t.tab.addEventListener('keydown', (e) => {
    if (e.key !== 'ArrowLeft' && e.key !== 'ArrowRight') return;
    const next = (i + (e.key === 'ArrowRight' ? 1 : tabs.length - 1)) % tabs.length;
    selectTab(next);
    tabs[next].tab.focus();
  });
});

// ---------------------------------------------------------------- export

const ui = {
  site: $('exp-site'),
  reload: $('exp-reload'),
  count: $('exp-count'),
  filter: $('exp-filter'),
  all: $('exp-all'),
  none: $('exp-none'),
  system: $('exp-system'),
  toggle: $('exp-toggle'),
  rows: $('exp-rows'),
  empty: $('exp-empty'),
  allCols: $('exp-allcols'),
  run: $('exp-run'),
  batches: $('exp-batches'),
  progress: $('exp-progress'),
  bar: $('exp-bar'),
  status: $('exp-status'),
  result: $('exp-result'),
};

const exp = {
  busy: false,
  site: null, // web URL
  request: null, // runs requests in the SharePoint tab
  data: null, // { web, siteCollectionUrl, lists }
  selected: new Set(), // list ids
};

function setBusy(busy) {
  exp.busy = busy;
  for (const n of [ui.reload, ui.filter, ui.all, ui.none, ui.system, ui.toggle, ui.allCols]) n.disabled = busy;
  for (const n of ui.rows.querySelectorAll('input')) n.disabled = busy;
  updateRunButton();
}

function setStatus(text) {
  ui.progress.hidden = false;
  ui.status.textContent = text;
}

/** Lists shown with the current filter and system-list switch. */
function visibleLists() {
  if (!exp.data) return [];
  const q = ui.filter.value.trim().toLowerCase();
  return exp.data.lists.filter(
    (l) => (ui.system.checked || !l.system) && (!q || l.title.toLowerCase().includes(q) || l.url.toLowerCase().includes(q))
  );
}

function updateRunButton() {
  const n = exp.selected.size;
  ui.run.disabled = ui.batches.disabled = exp.busy || !exp.data || n === 0;
  ui.run.textContent = n ? `Export CSV (${n} list${n === 1 ? '' : 's'})` : 'Export CSV';
  if (exp.data) {
    const total = exp.data.lists.filter((l) => ui.system.checked || !l.system).length;
    ui.count.textContent = `${n} of ${total} selected`;
  }
}

function renderTable() {
  const shown = visibleLists();
  ui.rows.replaceChildren(
    ...shown.map((l) => {
      const box = el('input', { type: 'checkbox', 'aria-label': `Export ${l.title}` });
      box.checked = exp.selected.has(l.id);
      box.disabled = exp.busy;
      box.addEventListener('change', () => {
        if (box.checked) exp.selected.add(l.id);
        else exp.selected.delete(l.id);
        syncToggle();
        updateRunButton();
      });
      const row = el('tr', { class: l.system ? 'system' : null },
        el('td', { class: 'sel' }, box),
        el('td', {}, el('span', { class: 'exp-title', text: l.title }), el('span', { class: 'exp-url', text: l.url })),
        el('td', { class: 'kind', text: l.kind.replace(/ \(\d+\)$/, ''), title: l.kind }),
        el('td', { class: 'num', text: l.itemCount ?? '' }),
        el('td', { class: 'perm', text: l.uniquePermissions ? 'Unique' : 'Inherited' }));
      row.addEventListener('click', (e) => {
        if (e.target !== box && !exp.busy) box.click();
      });
      return row;
    })
  );
  ui.empty.hidden = shown.length > 0;
  ui.empty.textContent = exp.data ? 'No lists match.' : '';
  syncToggle();
  updateRunButton();
}

function syncToggle() {
  const shown = visibleLists();
  const picked = shown.filter((l) => exp.selected.has(l.id)).length;
  ui.toggle.checked = shown.length > 0 && picked === shown.length;
  ui.toggle.indeterminate = picked > 0 && picked < shown.length;
}

async function loadLists() {
  setBusy(true);
  exp.data = null;
  exp.selected.clear();
  ui.result.replaceChildren();
  renderTable();
  try {
    const tab = await activeTab();
    if (!tab || !tab.url || !isSharePointUrl(tab.url)) {
      ui.site.textContent = 'The active tab is not a SharePoint page. Open any page of the site, then click Reload lists.';
      ui.site.classList.add('error');
      return;
    }
    const site = await siteOfTab(tab);
    exp.site = site;
    exp.request = tabRequester(tab.id, new URL(site).origin);
    ui.site.classList.remove('error');
    ui.site.textContent = `Reading lists of ${site}…`;
    const data = await createExporter({ request: exp.request }).loadLists(site);
    exp.data = data;
    ui.site.replaceChildren(el('strong', { text: data.web.Title }), el('span', { class: 'exp-url', text: data.web.Url }));
    renderTable();
  } catch (e) {
    ui.site.textContent = `Could not read the lists: ${e.message}`;
    ui.site.classList.add('error');
  } finally {
    setBusy(false);
  }
}

function download(text, name, type = 'text/csv;charset=utf-8') {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = el('a', { href: url, download: name });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

/** A small button that downloads the export (again). */
function downloadButton(text, name, type, label = 'Download') {
  const b = el('button', { type: 'button', class: 'secondary small', text: label, title: name });
  b.addEventListener('click', () => download(text, name, type));
  return b;
}

/** An exporter whose progress events drive the bar and the status line. */
function progressExporter() {
  let countdown;
  const exporter = createExporter({
    request: exp.request,
    onEvent: (e) => {
      clearInterval(countdown);
      if (e.type === 'list') {
        ui.bar.value = e.index / e.total;
        setStatus(`Reading "${e.title}" (${e.index + 1} of ${e.total})…`);
      } else if (e.type === 'wait') {
        const until = Date.now() + e.ms;
        const tick = () => setStatus(`SharePoint is throttling. Waiting ${Math.max(0, Math.ceil((until - Date.now()) / 1000))} s (retry ${e.retry} of ${e.of})…`);
        tick();
        countdown = setInterval(tick, 500);
      } else if (e.type === 'done') {
        ui.bar.value = 1;
      }
    },
  });
  return { exporter, stop: () => clearInterval(countdown) };
}

async function runExport() {
  if (!exp.data || !exp.selected.size || exp.busy) return;
  setBusy(true);
  ui.result.replaceChildren();
  ui.bar.value = 0;
  setStatus('Reading the site and its permissions…');
  const { exporter, stop } = progressExporter();
  try {
    const { rows, problems } = await exporter.exportSite(exp.site, exp.data, [...exp.selected], { allColumns: ui.allCols.checked });
    const name = csvFileName(exp.data.web.Title);
    const csv = toCsv(rows, SEPARATOR);
    download(csv, name);
    const byType = {};
    for (const r of rows) byType[r.Type] = (byType[r.Type] || 0) + 1;
    setStatus(`Exported ${rows.length} rows to ${name}.`);
    ui.result.replaceChildren(
      el('ul', { class: 'exp-summary' }, Object.entries(byType).map(([t, n]) => el('li', {}, el('span', { text: t }), el('span', { text: String(n) })))),
      problems.length
        ? el('details', { class: 'check' },
            el('summary', {}, el('span', { class: 'chip warn', text: String(problems.length) }), 'Parts that could not be read (marked in the CSV)'),
            el('div', { class: 'check-body' }, el('div', { class: 'msg warn' }, el('ul', {}, problems.map((p) => el('li', { text: p }))))))
        : null,
      el('div', { class: 'row' }, downloadButton(csv, name, 'text/csv;charset=utf-8', 'Download again'))
    );
  } catch (e) {
    setStatus(`Export failed: ${e.message}`);
  } finally {
    stop();
    setBusy(false);
  }
}

async function runBatchExport() {
  if (!exp.data || !exp.selected.size || exp.busy) return;
  setBusy(true);
  ui.result.replaceChildren();
  ui.bar.value = 0;
  setStatus('Reading the selected lists…');
  const { exporter, stop } = progressExporter();
  try {
    const data = await exporter.collect(exp.site, exp.data, [...exp.selected], { permissions: false, contentTypes: false });
    const { text, stages, notes } = buildProvisioning({
      web: exp.data.web,
      allLists: exp.data.lists,
      lists: data.lists,
      version: chrome.runtime.getManifest().version,
    });
    const name = batchFileName(exp.data.web.Title);
    const ops = stages.reduce((n, st) => n + st.ops.length, 0);
    setStatus(`Built ${stages.length} stage(s) with ${ops} operation(s). Load them into the Batch Runner or download ${name}.`);
    const load = el('button', { type: 'button', class: 'secondary small', text: 'Load into Batch Runner' });
    load.addEventListener('click', () => {
      selectTab(0);
      document.dispatchEvent(new CustomEvent('sp-utility:load-stages', {
        detail: { stages: (splitForRunner(text)), label: `"${name}" (${stages.length} stages)` },
      }));
    });
    ui.result.replaceChildren(
      el('ul', { class: 'exp-summary' }, stages.map((st) => el('li', {}, el('span', { text: st.name }), el('span', { text: `${st.ops.length} op(s)` })))),
      notes.length
        ? el('details', { class: 'check' },
            el('summary', {}, el('span', { class: 'chip warn', text: String(notes.length) }), 'Notes (also at the top of the file)'),
            el('div', { class: 'check-body' }, el('div', { class: 'msg warn' }, el('ul', {}, notes.map((n) => el('li', { text: n }))))))
        : null,
      el('div', { class: 'row' }, load, downloadButton(text, name, 'text/plain;charset=utf-8'))
    );
  } catch (e) {
    setStatus(`Export failed: ${e.message}`);
  } finally {
    stop();
    setBusy(false);
  }
}

ui.reload.addEventListener('click', loadLists);
ui.filter.addEventListener('input', renderTable);
ui.system.addEventListener('change', () => {
  if (!ui.system.checked && exp.data) for (const l of exp.data.lists) if (l.system) exp.selected.delete(l.id);
  renderTable();
});
const selectShown = (on) => {
  for (const l of visibleLists()) on ? exp.selected.add(l.id) : exp.selected.delete(l.id);
  renderTable();
};
ui.all.addEventListener('click', () => selectShown(true));
ui.none.addEventListener('click', () => selectShown(false));
ui.toggle.addEventListener('change', () => selectShown(ui.toggle.checked));
ui.run.addEventListener('click', runExport);
ui.batches.addEventListener('click', runBatchExport);
