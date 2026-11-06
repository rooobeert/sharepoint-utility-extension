// "Search Schema" tab: bulk-map crawled properties to managed (Refinable*) properties on the
// site collection of the active tab. Ported from the sp-schema-mapper bookmarklet.
// Unsupported: reads the classic search schema pages and replays the OK postback of
// crawledproperty.aspx. Rows and the status cache live in memory only.

import { el, activeTab, tabRequester, ask } from './shared.js';
import { pageReadContext } from './lib/runner.js';
import { isSharePointUrl, siteUrlFromPageUrl } from './lib/batch.js';
import {
  lc, sameName, has, splitCrawled, REFINABLE_SET, STATUS_FILTERS, STATUS_MAX_AGE, crawledSuggestions, managedSuggestions,
  crawledUrl, managedUrl, managedListUrl, columnsUrl, columnsFromLists, crawledStatus, patchStatus, mergeStatusLists, freeCount,
  parsePaste, rowState, otherFeeds, pendingChanges, nextMappings, verifyProblems,
  formBody, formAction, parseCrawledPage, mappingsPostback, pageError, parseManagedPage,
  managedListFilter, managedListNext, parseManagedList,
} from './lib/schema.js';

const $ = (id) => document.getElementById(id);

const ui = {
  reload: $('sch-reload'),
  site: $('sch-site'),
  statusText: $('sch-status-text'),
  refresh: $('sch-refresh'),
  clear: $('sch-clear'),
  rows: $('sch-rows'),
  save: $('sch-save'),
  foot: $('sch-foot'),
  suggest: $('sch-suggest'),
};

const sch = {
  shown: false,
  saving: false,
  loading: null, // the running loadSite()
  site: null, // site collection URL
  layouts: null,
  request: null, // runs pageRequest in the SharePoint tab
  columns: new Map(), // lc(crawled) -> { name, title, type, lists }
  status: new Map(), // lc(managed) -> { name, type, crawled: [] }
  statusAt: 0,
  statusLoading: null,
  statusPages: 0, // pages of the mapping status read so far
  columnsError: null, // why the list columns could not be read
  columnsLoading: false,
  crawledCache: new Map(),
  managedCache: new Map(),
  bySite: new Map(), // site -> { columns, status, statusAt }: kept while the panel is open
};

const rows = [];
const removals = new Map(); // lc(crawled) -> Map(lc(managed) -> managed)
const removalsFor = (c) => [...(removals.get(lc(c))?.values() || [])];

// ---------------------------------------------------------------- requests

// SharePoint takes ~1-8 s per classic page. Row checks and the bulk status load get separate
// lanes so the status load never blocks typing.
function makeLimiter(concurrency) {
  const queue = [];
  let active = 0;
  const pump = () => {
    while (active < concurrency && queue.length) {
      const { fn, resolve, reject } = queue.shift();
      active++;
      fn().then(resolve, reject).finally(() => { active--; pump(); });
    }
  };
  return (fn) => new Promise((resolve, reject) => { queue.push({ fn, resolve, reject }); pump(); });
}
const rowLane = makeLimiter(4);
const bulkLane = makeLimiter(2);

const parseHtml = (html) => new DOMParser().parseFromString(html, 'text/html');

async function send(req, lane) {
  const request = sch.request;
  const res = await lane(() => request(req));
  if (res.phase === 'network') throw new Error(res.statusText || 'network error');
  return res;
}

async function getDoc(url, lane = rowLane) {
  const res = await send({ url, method: 'GET', headers: {} }, lane);
  if (res.status !== 200) throw new Error(`HTTP ${res.status}`);
  return parseHtml(res.text);
}

async function postForm(url, body, lane, extra = {}) {
  return send({ url, method: 'POST', headers: { 'Content-Type': 'application/x-www-form-urlencoded' }, body: body.toString(), ...extra }, lane);
}

async function loadCrawled(value) {
  const url = crawledUrl(sch.layouts, value);
  return parseCrawledPage(await getDoc(url), url);
}
const loadManaged = async (name) => parseManagedPage(await getDoc(managedUrl(sch.layouts, name)));

function cached(cache, loader, key) {
  const k = lc(key);
  if (!cache.has(k)) {
    const p = loader(key);
    cache.set(k, p);
    p.catch(() => cache.delete(k));
  }
  return cache.get(k);
}
const crawledInfo = (v) => cached(sch.crawledCache, loadCrawled, v);
const managedInfo = (v) => cached(sch.managedCache, loadManaged, v);

/** Same postback as OK on crawledproperty.aspx. A redirect back to the list page means success. */
async function postMappings(page, next, deleted) {
  const { url, body } = mappingsPostback(page, next, deleted);
  const res = await postForm(url, body, rowLane, { redirect: 'manual' });
  if (res.type === 'opaqueredirect' || (res.status >= 300 && res.status < 400)) return;
  const msg = res.text ? pageError(parseHtml(res.text)) : '';
  throw new Error(`Save rejected (HTTP ${res.status})${msg ? ': ' + msg : ''}`);
}

async function postList(url, doc, overrides) {
  const form = doc.querySelector('form');
  const res = await postForm(formAction(form, url), formBody(form, overrides), bulkLane);
  if (res.status !== 200) throw new Error(`HTTP ${res.status}`);
  return parseHtml(res.text);
}

/** listmanagedproperties.aspx, filtered, 50 rows per page. */
async function loadManagedList(filter, onPage) {
  const url = managedListUrl(sch.layouts);
  let doc = await getDoc(url, bulkLane);
  doc = await postList(url, doc, managedListFilter(doc, filter));
  const all = [];
  const seen = new Set();
  for (let page = 0; page < 20; page++) {
    const items = parseManagedList(doc, url);
    // Stop on an empty page or one already seen (the last page may still offer "next").
    if (!items.length || seen.has(lc(items[0].name))) break;
    items.forEach((i) => seen.add(lc(i.name)));
    all.push(...items);
    onPage();
    const next = managedListNext(doc);
    if (!next || items.length < 50) break;
    doc = await postList(url, doc, next);
  }
  return all;
}

async function loadColumns() {
  const res = await send({ url: columnsUrl(sch.site), method: 'GET', headers: { Accept: 'application/json;odata=nometadata' } }, bulkLane);
  if (res.status !== 200) throw new Error(`HTTP ${res.status}`);
  return columnsFromLists(JSON.parse(res.text));
}

// ---------------------------------------------------------------- site

function loadSite() {
  if (!sch.loading) sch.loading = readSite().finally(() => { sch.loading = null; });
  return sch.loading;
}

function showSiteError(text) {
  ui.site.className = 'exp-site error';
  ui.site.textContent = text;
}

async function readSite() {
  const tab = await activeTab();
  if (!tab || !tab.url || !isSharePointUrl(tab.url)) {
    if (!sch.site) showSiteError('The active tab is not a SharePoint page. Open any page of the site collection in it.');
    else setFoot('The active tab is not a SharePoint page. Switch back to the site before saving.');
    return;
  }
  let site = null;
  try {
    const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: pageReadContext, world: 'MAIN' });
    site = res && res.result && res.result.siteAbsoluteUrl;
  } catch {
    /* fall back to the URL */
  }
  site = (site && isSharePointUrl(site) ? site : siteUrlFromPageUrl(tab.url)).replace(/\/$/, '');
  sch.request = tabRequester(tab.id, new URL(site).origin);
  if (site === sch.site) return;

  // Another site collection: keep the rows, check them again there.
  if (sch.site) sch.bySite.set(sch.site, { columns: sch.columns, status: sch.status, statusAt: sch.statusAt });
  const kept = sch.bySite.get(site);
  Object.assign(sch, {
    site,
    layouts: `${site}/_layouts/15`,
    columns: kept?.columns || new Map(),
    status: kept?.status || new Map(),
    statusAt: kept?.statusAt || 0,
    crawledCache: new Map(),
    managedCache: new Map(),
    columnsLoading: false,
  });
  removals.clear();
  ui.site.className = 'exp-site';
  ui.site.replaceChildren(el('strong', { text: tab.title || site }), el('span', { class: 'exp-url', text: site }));
  updateStatusText();
  renderSuggestions();
  for (const r of rows) validate(r, true);

  sch.columnsError = null;
  if (!kept) {
    sch.columnsLoading = true;
    updateSuggestHint();
    try {
      const columns = await loadColumns();
      if (sch.site === site) sch.columns = columns;
    } catch (e) {
      if (sch.site === site) sch.columnsError = e.message;
    }
    if (sch.site !== site) return;
    sch.columnsLoading = false;
    renderSuggestions();
    updateStatusText();
    rows.forEach(compose);
  }
  updateSuggestHint();
  await loadStatus(false);
}

// ---------------------------------------------------------------- mapping status

function updateStatusText(text, cls = '') {
  ui.statusText.className = cls;
  if (text) return void (ui.statusText.textContent = text);
  if (!sch.statusAt) return void (ui.statusText.textContent = `Mapping status not loaded · ${sch.columns.size} list columns`);
  const time = new Date(sch.statusAt).toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' });
  ui.statusText.textContent = `Status from ${time} · ${sch.status.size} Refinable properties, ${freeCount(sch.status)} free · ${sch.columns.size} list columns`;
}

/** Above the rows: what the dropdowns can't show yet, while it loads. */
function updateSuggestHint() {
  const { suggest } = ui;
  const show = (text, cls) => {
    suggest.hidden = false;
    suggest.className = `hint ${cls}`;
    suggest.textContent = text;
  };
  if (!sch.site) suggest.hidden = true;
  else if (sch.columnsLoading) show('Loading suggestions: reading the list columns…', 'sch-loading');
  else if (sch.columnsError) show(`Could not read the list columns (${sch.columnsError}). Type crawled property names yourself; each row is still checked.`, 'error');
  else if (sch.statusLoading && !sch.statusAt) {
    show(`Loading the mapping status${sch.statusPages ? ` (${sch.statusPages} pages)` : ''}… The dropdowns already suggest names but don't show yet what is mapped.`, 'sch-loading');
  } else suggest.hidden = true;
}

function loadStatus(force) {
  if (!sch.statusLoading) {
    sch.statusLoading = refreshStatus(force).finally(() => {
      sch.statusLoading = null;
      updateButtons();
      updateSuggestHint();
    });
  }
  updateButtons();
  updateSuggestHint();
  renderSuggestions();
  return sch.statusLoading;
}

async function refreshStatus(force) {
  const site = sch.site;
  if (!force && sch.statusAt && Date.now() - sch.statusAt < STATUS_MAX_AGE) return updateStatusText();
  sch.statusPages = 0;
  updateStatusText('Loading mapping status…');
  const onPage = () => {
    if (sch.site !== site) return;
    updateStatusText(`Loading mapping status… ${++sch.statusPages} pages`);
    updateSuggestHint();
  };
  try {
    const lists = await Promise.all(STATUS_FILTERS.map((f) => loadManagedList(f, onPage)));
    if (sch.site !== site) return;
    sch.status = new Map(mergeStatusLists(lists).map((i) => [lc(i.name), i]));
    sch.statusAt = Date.now();
    updateStatusText();
    renderSuggestions();
  } catch (e) {
    if (sch.site === site) updateStatusText(`Could not load the mapping status (${e.message}). The checks per row still work.`, 'error');
  }
}

// The managed-property list upper-cases crawled names; show the column's own casing when known.
const displayCrawled = (name) => sch.columns.get(lc(name))?.name || name;

// ---------------------------------------------------------------- suggestion dropdown
// One list, moved under the focused field. Same look and keys as the Permissions people search.

const listbox = el('ul', { id: 'sch-results', class: 'perm-results sch-results', role: 'listbox', hidden: true });
const combo = { input: null, kind: null, row: null, items: [], active: -1 };

function chipEl(cls, text, title) {
  return el('span', { class: `chip ${cls}`, text, title });
}
const firstPlus = (names) => `${names[0]}${names.length > 1 ? ` +${names.length - 1}` : ''}`;

/** "Not mapped" / "→ RefinableString07" for crawled; "Free" / "← ows_Region" for managed. Nothing until the status is loaded. */
function mappingChip(kind, item) {
  if (kind === 'crawled') {
    if (!sch.statusAt) return null;
    return item.mapped.length ? chipEl('info mono', '→ ' + firstPlus(item.mapped), item.mapped.join(', ')) : chipEl('muted', 'Not mapped');
  }
  if (item.state === 'free') return chipEl('ok', 'Free');
  if (item.state === 'missing') return chipEl('muted', 'Not on this site');
  if (item.state !== 'used') return null;
  const names = item.crawled.map(displayCrawled);
  return chipEl('warn mono', '← ' + firstPlus(names), names.join(', '));
}

function openSuggestions(row, kind) {
  const input = kind === 'crawled' ? row.cIn : row.mIn;
  Object.assign(combo, { input, kind, row });
  input.parentElement.append(listbox);
  renderSuggestions();
}

function closeSuggestions() {
  if (combo.input) {
    combo.input.setAttribute('aria-expanded', 'false');
    combo.input.removeAttribute('aria-activedescendant');
  }
  listbox.hidden = true;
  Object.assign(combo, { input: null, kind: null, row: null, items: [], active: -1 });
}

/** (Re)draw the open dropdown, e.g. after typing or when the columns or the status arrive. */
function renderSuggestions() {
  const { input, kind } = combo;
  if (!input) return;
  const q = input.value;
  const { items, more } = kind === 'crawled'
    ? crawledSuggestions(sch.columns, sch.status, q)
    : managedSuggestions(sch.status, !!sch.statusAt, q);
  const note = (text) => el('li', { class: 'perm-note', text });
  const notes = [];
  if (kind === 'crawled' && sch.columnsLoading) notes.push(note('Reading the list columns…'));
  if (sch.site && !sch.statusAt) notes.push(note(sch.statusLoading || sch.columnsLoading ? 'Mapping status still loading; shown once it\'s read.' : 'Mapping status not loaded.'));
  combo.items = items;
  combo.active = -1;
  input.removeAttribute('aria-activedescendant');
  listbox.replaceChildren(...[
    ...notes,
    ...items.map((item, i) => {
      const li = el('li', { role: 'option', id: `sch-opt-${i}`, class: 'sch-opt', 'aria-selected': 'false' },
        el('span', { class: 'perm-name', text: item.name }),
        el('span', { class: 'perm-kind', text: item.detail, title: item.detail }),
        mappingChip(kind, item));
      li.addEventListener('mousedown', (e) => {
        e.preventDefault(); // keep focus in the field
        pickSuggestion(item);
      });
      return li;
    }),
    more ? note(`${more} more · keep typing to narrow down`) : null,
    !items.length && q.trim() ? note('No suggestions. Any name can still be typed; the row checks it.') : null,
  ].filter(Boolean));
  listbox.hidden = !listbox.children.length;
  input.setAttribute('aria-expanded', String(!listbox.hidden));
}

function pickSuggestion(item) {
  const { input, row } = combo;
  input.value = item.name;
  closeSuggestions();
  ensureTrailingRow();
  schedule(row, 0);
}

function moveActive(delta) {
  if (listbox.hidden || !combo.items.length) return;
  combo.active = (combo.active + delta + combo.items.length) % combo.items.length;
  listbox.querySelectorAll('[role="option"]').forEach((n, i) => {
    n.classList.toggle('active', i === combo.active);
    n.setAttribute('aria-selected', String(i === combo.active));
  });
  const node = $(`sch-opt-${combo.active}`);
  combo.input.setAttribute('aria-activedescendant', node.id);
  node.scrollIntoView({ block: 'nearest' });
}

function comboKeydown(e, row, kind) {
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    if (combo.input !== e.target) openSuggestions(row, kind);
    moveActive(e.key === 'ArrowDown' ? 1 : -1);
  } else if (e.key === 'Enter' && combo.input === e.target && combo.active >= 0) pickSuggestion(combo.items[combo.active]);
  else if (e.key === 'Escape' && combo.input === e.target && !listbox.hidden) closeSuggestions();
  else return;
  e.preventDefault();
}

// ---------------------------------------------------------------- rows

function addRow(crawled = '', managed = '') {
  const input = (placeholder, label) => el('input', {
    type: 'text', placeholder, 'aria-label': label, autocomplete: 'off', spellcheck: 'false',
    role: 'combobox', 'aria-autocomplete': 'list', 'aria-controls': 'sch-results', 'aria-expanded': 'false',
  });
  const cIn = input('ows_Field', 'Crawled property');
  const mIn = input('RefinableString00', 'Managed property');
  cIn.value = crawled;
  mIn.value = managed;
  const num = el('span', { class: 'sch-num' });
  const chip = el('span', { class: 'chip', hidden: true });
  const note = el('span', { class: 'sch-note' });
  const del = el('button', { type: 'button', class: 'icon', title: 'Remove row', 'aria-label': 'Remove row', text: '✕' });
  const cHint = el('div', { class: 'sch-hint' });
  const mHint = el('div', { class: 'sch-hint' });
  const li = el('li', { class: 'sch-row' },
    el('div', { class: 'sch-head' }, num, chip, note, del),
    el('div', { class: 'sch-field' },
      el('label', { text: 'Crawled' }), el('div', { class: 'sch-combo' }, cIn), cHint,
      el('label', { text: 'Managed' }), el('div', { class: 'sch-combo' }, mIn), mHint));
  ui.rows.append(li);

  // c/m: values last checked; cr/mr: their results ({ loading } while in flight)
  const row = { li, cIn, mIn, cHint, mHint, num, chip, note, c: '', m: '', cr: null, mr: null, state: 'empty', timer: 0 };
  rows.push(row);
  renumber();
  // While typing, only check values that match a suggestion; anything else is checked on blur/Enter (change).
  const isKnown = () => {
    const c = lc(cIn.value.trim());
    const m = lc(mIn.value.trim());
    return (!c || sch.columns.has(c) || crawledStatus(sch.status.values()).has(c)) && (!m || REFINABLE_SET.has(m));
  };
  const onEdit = () => { ensureTrailingRow(); clearTimeout(row.timer); if (isKnown()) schedule(row, 300); };
  for (const [inp, kind] of [[cIn, 'crawled'], [mIn, 'managed']]) {
    inp.addEventListener('input', () => {
      onEdit();
      if (combo.input === inp) renderSuggestions(); else openSuggestions(row, kind);
    });
    inp.addEventListener('change', () => schedule(row, 0));
    inp.addEventListener('focus', () => openSuggestions(row, kind));
    inp.addEventListener('blur', () => setTimeout(() => { if (combo.input === inp) closeSuggestions(); }, 100));
    inp.addEventListener('keydown', (e) => comboKeydown(e, row, kind));
  }
  cIn.addEventListener('paste', (e) => handlePaste(e, row));
  del.addEventListener('click', () => {
    clearTimeout(row.timer);
    if (combo.row === row) closeSuggestions();
    rows.splice(rows.indexOf(row), 1);
    li.remove();
    ensureTrailingRow();
    renumber();
    updateButtons();
  });
  compose(row);
  if (crawled || managed) schedule(row, 0);
  return row;
}

function renumber() {
  rows.forEach((r, i) => { r.num.textContent = String(i + 1); });
}

function ensureTrailingRow() {
  const last = rows[rows.length - 1];
  if (!last || last.cIn.value.trim() || last.mIn.value.trim()) addRow();
}

function schedule(row, delay) {
  clearTimeout(row.timer);
  row.timer = setTimeout(() => validate(row), delay);
}

function handlePaste(e, row) {
  const parsed = parsePaste(e.clipboardData?.getData('text') || '');
  if (!parsed) return;
  e.preventDefault();
  closeSuggestions();
  let i = rows.indexOf(row);
  for (const [crawled, managed] of parsed) {
    const r = rows[i] || addRow();
    r.cIn.value = crawled;
    r.mIn.value = managed;
    schedule(r, 0);
    i++;
  }
  ensureTrailingRow();
}

/** Look up only the side(s) whose value changed (or both with force, e.g. after a save). */
function validate(row, force = false) {
  if (!sch.site) return Promise.resolve();
  const site = sch.site;
  const c = row.cIn.value.trim();
  const m = row.mIn.value.trim();
  const fail = (e) => ({ error: e.message });
  const current = () => sch.site === site && rows.includes(row);
  const jobs = [];
  if (force || c !== row.c) {
    row.c = c;
    row.cr = c ? { loading: true } : null;
    if (c) jobs.push(crawledInfo(c).catch(fail).then((r) => { if (current() && row.c === c) { row.cr = r; compose(row); } }));
  }
  if (force || m !== row.m) {
    row.m = m;
    row.mr = m ? { loading: true } : null;
    if (m) jobs.push(managedInfo(m).catch(fail).then((r) => { if (current() && row.m === m) { row.mr = r; compose(row); } }));
  }
  compose(row);
  return Promise.all(jobs);
}

function toggleRemoval(crawled, managed) {
  const k = lc(crawled);
  const set = removals.get(k) || new Map();
  if (set.has(lc(managed))) set.delete(lc(managed)); else set.set(lc(managed), managed);
  if (set.size) removals.set(k, set); else removals.delete(k);
  rows.filter((r) => r.c && sameName(r.c, crawled)).forEach(compose);
}

function mapChip(text, { removed = false, onClick, title } = {}) {
  const s = el('span', { class: removed ? 'sch-map removed' : 'sch-map' }, text);
  if (onClick) {
    const x = el('button', { type: 'button', title, 'aria-label': `${title}: ${text}`, text: removed ? '↶' : '×' });
    x.addEventListener('click', onClick);
    s.append(x);
  }
  return s;
}

function hint(node, parts, cls = '') {
  node.className = cls ? `sch-hint ${cls}` : 'sch-hint';
  node.replaceChildren(...parts);
}

const STATES = {
  empty: null, incomplete: null,
  checking: ['Checking…', 'muted'], ready: ['Ready', 'info'], warn: ['Ready ⚠', 'warn'], already: ['Mapped ✓', 'ok'],
  saving: ['Saving…', 'muted'], error: ['Error', 'fail'],
};

/** Render hints and the row state from the cached lookups and pending removals. */
function compose(row) {
  const { c, m, cr, mr } = row;
  const col = c && sch.columns.get(lc(c));
  const prefix = col ? `${col.title} · ${col.type} · ${[...col.lists].join(', ')} · ` : '';
  const removing = c ? removalsFor(c) : [];

  if (!cr) hint(row.cHint, []);
  else if (cr.loading) hint(row.cHint, [prefix + 'checking…']);
  else if (cr.error) hint(row.cHint, [cr.error], 'fail');
  else if (!cr.exists) hint(row.cHint, [prefix + 'not found (not crawled yet?)'], 'fail');
  else if (!cr.current.length) hint(row.cHint, [prefix + 'not mapped']);
  else {
    hint(row.cHint, [prefix + 'mapped to: ', ...cr.current.map((x) => {
      const removed = has(removing, x);
      return mapChip(x, { removed, title: removed ? 'Keep this mapping' : 'Remove this mapping on save', onClick: () => toggleRemoval(c, x) });
    })]);
  }

  if (!mr) hint(row.mHint, []);
  else if (mr.loading) hint(row.mHint, ['checking…']);
  else if (mr.error) hint(row.mHint, [mr.error], 'fail');
  else if (!mr.exists) hint(row.mHint, ['not found'], 'fail');
  else {
    const type = mr.type ? mr.type + ' · ' : '';
    const others = otherFeeds(mr, c);
    if (others.length) hint(row.mHint, [type + 'also fed by: ', ...others.map((x) => mapChip(displayCrawled(x)))], 'warn');
    else hint(row.mHint, [type + (mr.mapped.length ? 'fed only by this crawled property' : 'free')], 'ok');
  }

  const { state, detail } = rowState({ c, m, cr, mr, removing, columnType: col ? col.type : null, display: displayCrawled });
  setState(row, state, detail);
}

function setState(row, state, detail = '') {
  row.state = state;
  const look = STATES[state];
  row.chip.hidden = !look;
  if (look) {
    row.chip.textContent = look[0];
    row.chip.className = `chip ${look[1]}`;
  }
  row.note.textContent = state === 'incomplete' ? '' : detail;
  row.li.className = `sch-row${state === 'warn' ? ' warn' : state === 'error' ? ' fail' : ''}`;
  updateButtons();
}

// ---------------------------------------------------------------- save

const changes = () => pendingChanges(rows, removals);

function updateButtons() {
  const list = changes();
  const adds = list.reduce((n, g) => n + g.add.length, 0);
  const dels = list.reduce((n, g) => n + g.remove.length, 0);
  const busy = rows.some((r) => r.state === 'checking' || r.state === 'saving');
  ui.save.textContent = adds || dels ? `Save (${[adds && `+${adds}`, dels && `−${dels}`].filter(Boolean).join(' ')})` : 'Save';
  ui.save.disabled = sch.saving || busy || !(adds || dels);
  ui.clear.disabled = sch.saving;
  ui.reload.disabled = sch.saving;
  ui.refresh.disabled = sch.saving || !sch.site || !!sch.statusLoading;
}

/** While saving, the rows can't be focused or clicked (inert also covers rows and chips re-rendered meanwhile). */
function lockRows(on) {
  closeSuggestions();
  ui.rows.inert = on;
  ui.rows.classList.toggle('locked', on);
}

function setFoot(text, cls = '') {
  ui.foot.className = cls ? `status ${cls}` : 'status';
  ui.foot.textContent = text;
}

async function saveAll() {
  const list = changes();
  if (!list.length) return;
  const warned = rows.filter((r) => r.state === 'warn').length;
  const ok = await ask({
    title: 'Save search schema changes?',
    wide: true,
    ok: 'Save',
    body: [
      el('p', { class: 'target' }, sch.site),
      el('ul', {}, list.map((g) => el('li', {},
        el('code', { text: g.crawled }), ': ',
        [...g.add.map((m) => '+' + m), ...g.remove.map((m) => '−' + m)].join(', ')))),
      warned ? el('p', { class: 'msg warn', text: `${warned} row${warned === 1 ? ' has' : 's have'} warnings (see the notes on the rows).` }) : null,
    ].filter(Boolean),
  });
  if (!ok) return;

  sch.saving = true;
  lockRows(true); // no edits while the save runs: the rows are what is being saved
  const site = sch.site;
  const t0 = performance.now();
  setFoot(`Saving ${list.length} crawled propert${list.length === 1 ? 'y' : 'ies'}…`);
  let failed = 0;
  try {
    await Promise.all(list.map(async (g) => {
      const affected = rows.filter((r) => r.c && sameName(r.c, g.crawled));
      affected.forEach((r) => setState(r, 'saving'));
      try {
        sch.crawledCache.delete(lc(g.crawled));
        const info = await crawledInfo(g.crawled);
        if (!info.exists) throw new Error('crawled property not found');
        const { next, deleted } = nextMappings(info.current, g.add, g.remove);
        await postMappings(info.page, next, deleted);
        sch.crawledCache.delete(lc(g.crawled));
        [...g.add, ...g.remove].forEach((m) => sch.managedCache.delete(lc(m)));
        const problems = verifyProblems((await crawledInfo(g.crawled)).current || [], g.add, g.remove);
        if (problems.length) throw new Error(`saved, but the page shows ${problems.join(', ')}`);
        removals.delete(lc(g.crawled));
        patchStatus(sch.status, splitCrawled(g.crawled).name, g.add, g.remove);
      } catch (e) {
        failed++;
        affected.forEach((r) => setState(r, 'error', e.message));
        return;
      }
      await Promise.all(affected.map((r) => validate(r, true)));
    }));
  } finally {
    sch.saving = false;
    lockRows(false);
  }
  if (sch.site === site) {
    renderSuggestions();
    updateStatusText();
  }
  updateButtons();
  const secs = ((performance.now() - t0) / 1000).toFixed(1);
  if (failed) setFoot(`${list.length - failed} of ${list.length} crawled properties saved, ${failed} failed (${secs}s).`, 'error');
  else setFoot(`Saved in ${secs}s · ${list.length} of ${list.length} confirmed. Reindex the list or library so values appear.`, 'ok');
}

// ---------------------------------------------------------------- events

document.addEventListener('sp-utility:tab-shown', (e) => {
  sch.shown = e.detail.id === 'sch';
  if (sch.shown && !sch.saving) loadSite();
});

// The site follows the active tab while this tab is visible.
chrome.tabs.onActivated.addListener(() => {
  if (sch.shown && !sch.saving) loadSite();
});
chrome.tabs.onUpdated.addListener((_id, info) => {
  if (sch.shown && !sch.saving && info.status === 'complete') loadSite();
});

ui.reload.addEventListener('click', async () => {
  if (sch.site) sch.bySite.delete(sch.site);
  sch.site = null;
  await loadSite();
});
ui.refresh.addEventListener('click', () => loadStatus(true));
ui.clear.addEventListener('click', () => {
  closeSuggestions();
  rows.splice(0).forEach((r) => { clearTimeout(r.timer); r.li.remove(); });
  removals.clear();
  setFoot('');
  addRow();
  updateButtons();
});
ui.save.addEventListener('click', saveAll);

addRow();
updateStatusText('Open any page of the site collection in the active tab.');
