import {
  parseBatch,
  analyseBatch,
  normalizeSiteUrl,
  siteUrlFromPageUrl,
  isSharePointUrl,
  targetSiteFromStages,
} from './lib/batch.js';
import { createRunner, pageReadContext } from './lib/runner.js';
import { findPlaceholders, describePlaceholder } from './lib/placeholders.js';
import { formatRunText, formatRunJson, shortUrl, stageSummaryText } from './lib/log.js';
import { el, activeTab, tabRequester, ask } from './shared.js';
import { splitStageFile } from './lib/provision.js';

const $ = (id) => document.getElementById(id);

const ui = {
  site: $('site'),
  detect: $('detect'),
  siteHint: $('site-hint'),
  addStage: $('add-stage'),
  loadFolder: $('load-folder'),
  folderInput: $('folder-input'),
  loadFile: $('load-file'),
  fileInput: $('file-input'),
  stageList: $('stage-list'),
  noStages: $('no-stages'),
  template: $('stage-template'),
  preview: $('preview'),
  run: $('run'),
  runHint: $('run-hint'),
  progress: $('progress'),
  progressBar: $('progress-bar'),
  status: $('status'),
  output: $('output'),
  outputHead: $('output-h'),
  outputBody: $('output-body'),
  copyLog: $('copy-log'),
  exportJson: $('export-json'),
};

const state = {
  stages: [], // { id, name, body }
  lastRun: null,
  running: false,
  preview: null, // { key, errorCount, warningCount } of the last Preview
  lastStageId: null, // stage that had focus last (target for Ctrl+F)
};
let nextId = 1;

// ---------------------------------------------------------------- helpers


function debounce(fn, ms) {
  let t;
  return (...args) => {
    clearTimeout(t);
    t = setTimeout(() => fn(...args), ms);
  };
}

const sortByFileName = (a, b) => a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' });
const stripTxt = (name) => name.replace(/\.txt$/i, '');

// ---------------------------------------------------------------- storage

const save = debounce(() => {
  chrome.storage.local.set({
    siteUrl: normalizeSiteUrl(ui.site.value),
    stages: state.stages.map(({ name, body }) => ({ name, body })),
  });
}, 400);

async function load() {
  const stored = await chrome.storage.local.get(['siteUrl', 'stages']);
  if (stored.siteUrl) ui.site.value = stored.siteUrl;
  for (const s of stored.stages || []) state.stages.push({ id: nextId++, name: s.name, body: s.body });
  renderStages();
  if (!ui.site.readOnly) await detectSite({ auto: true });
}

// ---------------------------------------------------------------- site detection


function setHint(content, isError = false) {
  ui.siteHint.replaceChildren(...(Array.isArray(content) ? content : [content]));
  ui.siteHint.classList.toggle('error', isError);
}

function safeParse(body) {
  try {
    return parseBatch(body);
  } catch {
    return { operations: [] };
  }
}

/** One line per site when stages disagree: which stages use it and how often. */
function conflictLines(sites) {
  return sites.sites.map(
    (s) => `${s.site}: ` + s.uses.map((u) => `stage ${u.stage + 1} "${u.name}" (${u.operations.length} op(s))`).join(', ')
  );
}

/**
 * The batches always target exactly one site, so the site URL comes from the
 * operation URLs in the batch bodies and the field is locked. Only batches
 * with relative URLs (no host) leave the field editable.
 */
function updateSiteFromStages() {
  const sites = targetSiteFromStages(state.stages.map((s) => ({ name: s.name, parsed: safeParse(s.body) })));
  const wasLocked = ui.site.readOnly;
  ui.site.readOnly = !sites.relativeOnly;
  ui.detect.hidden = !sites.relativeOnly;
  if (sites.site) {
    if (normalizeSiteUrl(ui.site.value) !== normalizeSiteUrl(sites.site)) {
      ui.site.value = normalizeSiteUrl(sites.site);
      save();
    }
    setHint('Taken from the batch bodies: every operation targets this site. It is checked with SharePoint before Preview and Run.');
  } else if (!sites.relativeOnly) {
    ui.site.value = '';
    setHint(['The stages target different sites. All stages must target one site:', el('ul', {}, conflictLines(sites).map((l) => el('li', { text: l })))], true);
  } else if (wasLocked) {
    setHint('The batches don\'t name a site (relative URLs or {{WebUrl}}). Detect the site from the active tab or enter it.');
  }
  return sites;
}

async function detectSite({ auto = false } = {}) {
  const tab = await activeTab();
  if (!tab || !tab.url || !isSharePointUrl(tab.url)) {
    if (!auto || !ui.site.value) setHint('The active tab is not a SharePoint page (https://*.sharepoint.com). Open the site, then click Detect.', !auto);
    else setHint('Using the last site. Open a page of the site in the active tab before running.');
    return;
  }
  let detected = null;
  let source = 'the page URL';
  try {
    const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: pageReadContext, world: 'MAIN' });
    if (res && res.result && res.result.webAbsoluteUrl && isSharePointUrl(res.result.webAbsoluteUrl)) {
      detected = res.result.webAbsoluteUrl;
      source = '_spPageContextInfo';
    }
  } catch {
    /* page not ready or not scriptable: fall back to the URL */
  }
  detected = detected || siteUrlFromPageUrl(tab.url);
  ui.site.value = normalizeSiteUrl(detected);
  setHint(`Detected from ${source}. You can override it.`);
  save();
}

// ---------------------------------------------------------------- stages

function addStage(name, body = '') {
  state.stages.push({ id: nextId++, name: name || `Stage ${state.stages.length + 1}`, body });
}

function updateCount(node, body) {
  const count = node.querySelector('.stage-count');
  count.classList.remove('bad');
  if (!body.trim()) {
    count.textContent = 'empty';
    return;
  }
  try {
    const n = parseBatch(body).operations.length;
    count.textContent = `${n} op${n === 1 ? '' : 's'}`;
    if (n > 100) count.classList.add('bad');
  } catch {
    count.textContent = 'invalid';
    count.classList.add('bad');
  }
}

function renderStages() {
  ui.stageList.replaceChildren();
  state.stages.forEach((stage, i) => {
    const node = ui.template.content.firstElementChild.cloneNode(true);
    node.dataset.id = stage.id;
    node.querySelector('.stage-num').textContent = `${i + 1}.`;
    const name = node.querySelector('.stage-name');
    const body = node.querySelector('.stage-body');
    const file = node.querySelector('.stage-file');
    name.value = stage.name;
    body.value = stage.body;
    updateCount(node, stage.body);

    const recount = debounce(() => {
      updateCount(node, stage.body);
      updateSiteFromStages();
    }, 250);
    name.addEventListener('input', () => {
      stage.name = name.value;
      save();
    });
    body.addEventListener('input', () => {
      stage.body = body.value;
      recount();
      save();
      updateRunState();
    });
    node.addEventListener('focusin', () => (state.lastStageId = stage.id));
    node.find = attachFind(node);
    file.addEventListener('change', async () => {
      const f = file.files[0];
      if (!f) return;
      stage.body = await f.text();
      if (!stage.name.trim() || /^Stage \d+$/.test(stage.name)) stage.name = stripTxt(f.name);
      file.value = '';
      renderStages();
      save();
    });
    node.querySelector('.stage-tools').addEventListener('click', (e) => {
      const act = e.target.closest('button')?.dataset.act;
      if (!act) return;
      if (act === 'load') return file.click();
      if (act === 'remove') state.stages.splice(state.stages.indexOf(stage), 1);
      if (act === 'up') return moveStage(stage, -1);
      if (act === 'down') return moveStage(stage, 1);
      renderStages();
      save();
    });

    // Drag handle at the front of the row; Arrow Up/Down when it has focus.
    const handle = node.querySelector('.drag-handle');
    handle.addEventListener('dragstart', (e) => {
      drag.id = stage.id;
      e.dataTransfer.effectAllowed = 'move';
      e.dataTransfer.setData('text/plain', `stage:${stage.id}`);
      e.dataTransfer.setDragImage(node, 16, 16);
      node.classList.add('dragging');
    });
    handle.addEventListener('dragend', () => {
      node.classList.remove('dragging');
      clearDropMarks();
      drag.id = null;
    });
    handle.addEventListener('keydown', (e) => {
      if (e.key !== 'ArrowUp' && e.key !== 'ArrowDown') return;
      e.preventDefault();
      moveStage(stage, e.key === 'ArrowUp' ? -1 : 1, true);
    });
    node.querySelector('[data-act="up"]').disabled = i === 0;
    node.querySelector('[data-act="down"]').disabled = i === state.stages.length - 1;
    ui.stageList.append(node);
  });
  ui.noStages.hidden = state.stages.length > 0;
  updateSiteFromStages();
  setBusy(state.running);
}

// ---------------------------------------------------------------- reordering

const drag = { id: null, target: null }; // stage being dragged, and where it would land

/** Move a stage up (-1) or down (+1). */
function moveStage(stage, delta, keepFocus = false) {
  const from = state.stages.indexOf(stage);
  const to = from + delta;
  if (to < 0 || to >= state.stages.length) return;
  state.stages.splice(to, 0, ...state.stages.splice(from, 1));
  renderStages();
  save();
  if (keepFocus) ui.stageList.querySelector(`.stage[data-id="${stage.id}"] .drag-handle`)?.focus();
}

/** Put the dragged stage before or after the target stage. */
function dropStage(dragId, targetId, after) {
  const from = state.stages.findIndex((s) => s.id === dragId);
  if (from < 0 || dragId === targetId) return;
  const [moved] = state.stages.splice(from, 1);
  const to = state.stages.findIndex((s) => s.id === targetId);
  state.stages.splice(to + (after ? 1 : 0), 0, moved);
  renderStages();
  save();
}

function clearDropMarks() {
  drag.target = null;
  for (const n of ui.stageList.querySelectorAll('.drop-before, .drop-after')) n.classList.remove('drop-before', 'drop-after');
}

ui.stageList.addEventListener('dragover', (e) => {
  if (drag.id == null) return; // text dragged into a textarea etc.
  e.preventDefault();
  e.dataTransfer.dropEffect = 'move';
  const target = e.target.closest('.stage');
  if (!target) return;
  const box = target.getBoundingClientRect();
  const after = e.clientY > box.top + box.height / 2;
  if (drag.target && drag.target.node === target && drag.target.after === after) return;
  clearDropMarks();
  target.classList.add(after ? 'drop-after' : 'drop-before');
  drag.target = { node: target, id: Number(target.dataset.id), after };
});
ui.stageList.addEventListener('dragleave', (e) => {
  if (drag.id != null && !ui.stageList.contains(e.relatedTarget)) clearDropMarks();
});
ui.stageList.addEventListener('drop', (e) => {
  if (drag.id == null) return;
  e.preventDefault(); // never drop the drag text into a textarea
  const t = drag.target;
  clearDropMarks();
  if (t) dropStage(drag.id, t.id, t.after);
});

/** Add stages, asking whether to replace the existing ones. */
async function addStages(stages, what, { append = false } = {}) {
  if (!stages.length) return;
  let replace = false;
  if (state.stages.length && !append) {
    replace = await ask({
      title: `Load ${what}`,
      body: `Replace the ${state.stages.length} existing stage(s), or add the ${stages.length} new stage(s) after them?`,
      ok: 'Replace',
      cancel: 'Add to end',
    });
  }
  if (replace) state.stages = [];
  for (const st of stages) addStage(st.name, st.body);
  renderStages();
  save();
}

/** Stages from .txt files: a file with "### STAGE:" markers becomes several stages. */
async function stagesFromFiles(files) {
  const out = [];
  for (const f of files) out.push(...splitStageFile(await f.text(), stripTxt(f.name)));
  return out;
}

async function loadFolder(fileList) {
  const files = [...fileList].filter((f) => /\.txt$/i.test(f.name)).sort(sortByFileName);
  if (!files.length) {
    await ask({ title: 'No .txt files', body: 'The selected folder contains no .txt files.', cancel: null, ok: 'OK' });
    return;
  }
  await addStages(await stagesFromFiles(files), `${files.length} file(s)`);
}

async function loadBatchFile(file) {
  const stages = await stagesFromFiles([file]);
  await addStages(stages, `"${file.name}" (${stages.length} stage${stages.length === 1 ? '' : 's'})`);
}

// The Export tab can hand over generated stages directly.
document.addEventListener('sp-utility:load-stages', async (e) => {
  await addStages(e.detail.stages, e.detail.label || 'exported batches', { append: e.detail.append });
});

// ---------------------------------------------------------------- find / replace

/**
 * Find/replace bar of one stage. Matches are painted on a mirror layer behind
 * the (transparent) textarea, so they stay visible while typing in the bar.
 * Replacements go through execCommand('insertText') so Ctrl+Z undoes them.
 */
function attachFind(node) {
  const bar = node.querySelector('.find-bar');
  const findInput = node.querySelector('.find-input');
  const replaceInput = node.querySelector('.replace-input');
  const count = node.querySelector('.find-count');
  const caseBtn = node.querySelector('[data-find="case"]');
  const ta = node.querySelector('.stage-body');
  const mirror = node.querySelector('.stage-mirror');
  const f = { matches: [], current: -1, matchCase: false };

  const syncScroll = () => {
    mirror.scrollTop = ta.scrollTop;
    mirror.scrollLeft = ta.scrollLeft;
  };

  function search(keepCurrent = true) {
    const q = findInput.value;
    const text = ta.value;
    f.matches = [];
    if (q) {
      const hay = f.matchCase ? text : text.toLowerCase();
      const needle = f.matchCase ? q : q.toLowerCase();
      for (let i = hay.indexOf(needle); i !== -1; i = hay.indexOf(needle, i + needle.length)) f.matches.push(i);
    }
    if (!keepCurrent || f.current < 0) {
      const caret = ta.selectionStart || 0;
      f.current = f.matches.findIndex((m) => m >= caret);
      if (f.current < 0 && f.matches.length) f.current = 0;
    }
    if (f.current >= f.matches.length) f.current = f.matches.length ? 0 : -1;
    paint();
  }

  function paint(scroll = true) {
    const q = findInput.value;
    const n = f.matches.length;
    count.textContent = !q ? '' : n ? `${f.current + 1} of ${n}` : 'No results';
    count.classList.toggle('none', !!q && !n);
    mirror.replaceChildren();
    if (!q || !n) return;
    const text = ta.value;
    let pos = 0;
    f.matches.forEach((m, i) => {
      mirror.append(text.slice(pos, m), el('mark', { class: i === f.current ? 'current' : null }, text.slice(m, m + q.length)));
      pos = m + q.length;
    });
    mirror.append(text.slice(pos) + '\n');
    const cur = mirror.querySelector('mark.current');
    if (scroll && cur) {
      if (cur.offsetTop < ta.scrollTop || cur.offsetTop + cur.offsetHeight > ta.scrollTop + ta.clientHeight) {
        ta.scrollTop = cur.offsetTop - ta.clientHeight / 2;
      }
      if (cur.offsetLeft < ta.scrollLeft || cur.offsetLeft + cur.offsetWidth > ta.scrollLeft + ta.clientWidth) {
        ta.scrollLeft = cur.offsetLeft - ta.clientWidth / 3;
      }
    }
    syncScroll();
  }

  function step(dir) {
    if (!f.matches.length) return;
    f.current = (f.current + dir + f.matches.length) % f.matches.length;
    paint();
  }

  /** Replace ta[start, end) with text as one undoable edit. */
  function edit(start, end, text) {
    ta.focus();
    ta.setSelectionRange(start, end);
    const ok = text ? document.execCommand('insertText', false, text) : document.execCommand('delete');
    if (!ok) {
      ta.setRangeText(text, start, end, 'end');
      ta.dispatchEvent(new Event('input', { bubbles: true }));
    }
  }

  function replaceOne() {
    if (f.current < 0) return;
    const m = f.matches[f.current];
    edit(m, m + findInput.value.length, replaceInput.value);
    search(); // the next match now has the same index
    replaceInput.focus();
  }

  function replaceAll() {
    const n = f.matches.length;
    if (!n) return;
    const q = findInput.value;
    const escaped = q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const next = ta.value.replace(new RegExp(escaped, f.matchCase ? 'g' : 'gi'), () => replaceInput.value);
    edit(0, ta.value.length, next);
    f.current = -1;
    search();
    count.textContent = `Replaced ${n}`;
    replaceInput.focus();
  }

  function open() {
    const sel = ta.value.slice(ta.selectionStart, ta.selectionEnd);
    if (sel && !sel.includes('\n')) findInput.value = sel;
    bar.hidden = false;
    findInput.focus();
    findInput.select();
    search(false);
  }

  function close() {
    bar.hidden = true;
    mirror.replaceChildren();
    ta.focus();
    if (f.current >= 0 && findInput.value) {
      const m = f.matches[f.current];
      ta.setSelectionRange(m, m + findInput.value.length);
    }
  }

  findInput.addEventListener('input', () => search(false));
  findInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      step(e.shiftKey ? -1 : 1);
    }
  });
  replaceInput.addEventListener('keydown', (e) => {
    if (e.key === 'Enter') {
      e.preventDefault();
      replaceOne();
    }
  });
  bar.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      e.preventDefault();
      e.stopPropagation();
      close();
    }
  });
  bar.addEventListener('click', (e) => {
    const act = e.target.closest('button')?.dataset.find;
    if (act === 'next') step(1);
    else if (act === 'prev') step(-1);
    else if (act === 'replace') replaceOne();
    else if (act === 'all') replaceAll();
    else if (act === 'close') close();
    else if (act === 'case') {
      f.matchCase = !f.matchCase;
      caseBtn.setAttribute('aria-pressed', String(f.matchCase));
      search(false);
    }
  });
  ta.addEventListener('scroll', syncScroll);
  ta.addEventListener('input', () => {
    if (!bar.hidden) search();
  });
  new ResizeObserver(() => !bar.hidden && paint(false)).observe(ta);
  return { open, close };
}

// Ctrl+F / Cmd+F opens find in the stage that has (or last had) focus.
document.addEventListener('keydown', (e) => {
  if (!(e.ctrlKey || e.metaKey) || e.altKey || e.key.toLowerCase() !== 'f' || state.running) return;
  if (document.getElementById('view-batch').hidden) return; // only in the Batch Runner tab
  const node =
    document.activeElement?.closest?.('.stage') ||
    ui.stageList.querySelector(`.stage[data-id="${state.lastStageId}"]`) ||
    ui.stageList.querySelector('.stage');
  if (!node) return;
  e.preventDefault();
  node.find.open();
});

// ---------------------------------------------------------------- running

function setBusy(busy) {
  state.running = busy;
  for (const b of [ui.detect, ui.addStage, ui.loadFolder, ui.loadFile, ui.preview, ui.run, ui.copyLog, ui.exportJson]) b.disabled = busy;
  ui.site.disabled = busy;
  for (const n of ui.stageList.querySelectorAll('input, textarea, button')) n.disabled = busy;
  if (!busy) {
    // first stage can't move up, last can't move down
    ui.stageList.querySelectorAll('.stage').forEach((node, i, all) => {
      node.querySelector('[data-act="up"]').disabled = i === 0;
      node.querySelector('[data-act="down"]').disabled = i === all.length - 1;
    });
  }
  ui.run.textContent = busy ? 'Running…' : 'Run';
  if (!busy) updateRunState();
}

/** What Preview was run on: the site and the stage bodies. Any change invalidates it. */
function previewKey() {
  return JSON.stringify([normalizeSiteUrl(ui.site.value), state.stages.map((s) => s.body)]);
}

/** Run needs a Preview of the current stages without errors. Warnings don't block. */
function updateRunState() {
  if (state.running) return;
  const p = state.preview;
  let text;
  let cls = '';
  let enabled = false;
  if (!state.stages.length) text = 'Add a stage, then run Preview.';
  else if (!p) text = 'Run Preview first. Run is enabled once Preview finds no errors.';
  else if (p.key !== previewKey()) text = 'The stages changed since the last Preview. Run Preview again.';
  else if (p.errorCount) {
    text = `Preview found ${p.errorCount} error(s) that block the run. Fix them and run Preview again.`;
    cls = 'error';
  } else {
    enabled = true;
    cls = 'ok';
    text = p.warningCount
      ? `Preview done with ${p.warningCount} warning(s). Warnings don't block the run.`
      : 'Preview done, no warnings. Ready to run.';
  }
  ui.run.disabled = !enabled;
  ui.runHint.textContent = text;
  ui.runHint.className = `hint ${cls}`.trim();
}

/** Show why Run can't go ahead. Always resolves to null. */
async function block(title, lines) {
  setStatus(`${title}.`);
  await ask({ title, body: el('ul', {}, lines.map((l) => el('li', { text: l }))), ok: 'OK', cancel: null });
  return null;
}

/** Site collection of the page in the tab: _spPageContextInfo, else guessed from the URL. */
async function tabSiteCollection(tab) {
  try {
    const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: pageReadContext, world: 'MAIN' });
    const sc = res && res.result && res.result.siteAbsoluteUrl;
    if (sc && isSharePointUrl(sc)) return normalizeSiteUrl(sc);
  } catch {
    /* fall back to the URL */
  }
  return siteUrlFromPageUrl(tab.url);
}

/**
 * Site checks, in order:
 * 1. all stages target one site (taken from the batch bodies);
 * 2. the active tab is a page of that site collection;
 * 3. SharePoint confirms the site exists at exactly that URL.
 * Returns { site, problems, info?, request? }. `request` (runs requests in the
 * tab) is only set when the tab is on the right tenant.
 */
async function checkSite() {
  const sites = updateSiteFromStages();
  if (!sites.site && !sites.relativeOnly) {
    return { site: null, problems: ['The stages target different sites. A run must target exactly one site:', ...conflictLines(sites)] };
  }
  const site = normalizeSiteUrl(sites.site || ui.site.value);
  if (!isSharePointUrl(site)) {
    return { site, problems: [`"${site || '(empty)'}" is not a https://<tenant>.sharepoint.com site URL.`] };
  }
  const tab = await activeTab();
  const origin = new URL(site).origin;
  if (!tab || !tab.url || new URL(tab.url).origin !== origin) {
    return { site, problems: [`Open a page of ${site} in the active tab. Requests are sent from that tab with your sign-in.`] };
  }
  const request = tabRequester(tab.id, origin);
  let info;
  try {
    info = await createRunner({ request }).verifySite(site);
  } catch (e) {
    return { site, problems: [e.message], request };
  }
  const problems = [];
  if (info.webUrl.toLowerCase() !== site.toLowerCase()) {
    problems.push(`The batches target ${site}, but SharePoint resolves that address to ${info.webUrl}. The URL in the batch bodies is probably misspelled, or the site does not exist.`);
  }
  const tabSc = await tabSiteCollection(tab);
  if (!tabSc || tabSc.toLowerCase() !== info.siteCollectionUrl.toLowerCase()) {
    problems.push(`The active tab is in ${tabSc || 'an unknown site'}, but the batches target site collection ${info.siteCollectionUrl}. Open a page of the target site in the active tab.`);
  }
  if (!problems.length) setHint(`Verified with SharePoint: "${info.title}" at ${info.webUrl}.`);
  return { site, problems, info, request };
}


function stagesToRun() {
  const problems = [];
  const parsed = state.stages.map((s, i) => {
    if (!s.body.trim()) {
      problems.push(`Stage ${i + 1} "${s.name}" is empty.`);
      return null;
    }
    try {
      return parseBatch(s.body);
    } catch (e) {
      problems.push(`Stage ${i + 1} "${s.name}": ${e.message}`);
      return null;
    }
  });
  if (!state.stages.length) problems.push('Add at least one stage.');
  return { parsed, problems };
}

let countdown = null;
function setStatus(text, wait = false) {
  clearInterval(countdown);
  ui.status.textContent = text;
  ui.status.classList.toggle('wait', wait);
}

function onRunnerEvent(total) {
  let part = null; // current part of a stage that is sent in parts
  return (e) => {
    const label = `Stage ${e.stage + 1}/${total}`;
    if (e.type === 'stage-start') {
      part = null;
      ui.progressBar.value = e.stage / total;
      setStatus(`${label} "${e.name}": resolving placeholders…`);
    } else if (e.type === 'resolve') {
      setStatus(`${label}: looking up lists…`);
    } else if (e.type === 'part') {
      part = e;
    } else if (e.type === 'send') {
      const where = part ? `, part ${part.part} (from line ${part.from})` : '';
      setStatus(`${label}${where}: sending ${e.operations} operation(s)${e.attempt ? ` (retry ${e.attempt})` : ''}…`);
    } else if (e.type === 'wait') {
      const until = Date.now() + e.ms;
      const tick = () => {
        const s = Math.max(0, Math.ceil((until - Date.now()) / 1000));
        ui.status.textContent = `${label}: ${e.reason}. Waiting ${s} s before retry ${e.retry} of ${e.of}…`;
      };
      setStatus('', true);
      tick();
      countdown = setInterval(tick, 500);
    } else if (e.type === 'stage-done') {
      ui.progressBar.value = (e.stage + 1) / total;
      setStatus(`${label}: ${stageSummaryText(e.result)}`);
      renderRun(state.lastRun);
    }
  };
}

/** Preview one stage. With a runner, placeholders are looked up; without one, parse only. */
async function previewStage(runner, site, stage) {
  const empty = { name: stage.name, operations: [], errors: [], warnings: [], placeholders: [] };
  if (!stage.body.trim()) return { ...empty, errors: ['The stage is empty.'] };
  try {
    if (runner) return await runner.preview(site, stage);
    const parsed = parseBatch(stage.body);
    return {
      ...empty,
      operations: parsed.operations,
      errors: parsed.errors,
      warnings: [...parsed.warnings, ...analyseBatch(parsed)],
      placeholders: findPlaceholders(stage.body).map((p) => ({
        label: describePlaceholder(p),
        id: null,
        error: 'Not looked up because the site check failed.',
      })),
    };
  } catch (e) {
    return { ...empty, errors: [e.message] };
  }
}

async function runPreview() {
  if (!state.stages.length || state.running) return;
  const key = previewKey();
  setBusy(true);
  ui.progress.hidden = false;
  ui.progressBar.removeAttribute('value');
  setStatus('Preview: checking the site with SharePoint (nothing is sent)…');
  try {
    const siteCheck = await checkSite();
    const runner = siteCheck.request && !siteCheck.problems.length ? createRunner({ request: siteCheck.request }) : null;
    setStatus('Preview: parsing stages and looking up placeholders (nothing is sent)…');
    const previews = [];
    for (const [i, stage] of state.stages.entries()) {
      previews.push({ ...(await previewStage(runner, siteCheck.site, stage)), stageId: stage.id });
      ui.progressBar.value = (i + 1) / state.stages.length;
    }
    const errorCount = siteCheck.problems.length + previews.reduce((n, p) => n + p.errors.length, 0);
    const warningCount = previews.reduce((n, p) => n + p.warnings.length + p.placeholders.filter((x) => !x.id && !x.pending).length, 0);
    state.preview = { key, errorCount, warningCount };
    renderPreview(siteCheck, previews);
    const ops = previews.reduce((n, p) => n + p.operations.length, 0);
    setStatus(`Preview: ${previews.length} stage(s), ${ops} operation(s), ${errorCount} error(s), ${warningCount} warning(s). Nothing was sent.`);
  } catch (e) {
    state.preview = null;
    setStatus(`Preview failed: ${e.message}`);
  } finally {
    setBusy(false);
  }
}

async function runAll() {
  if (ui.run.disabled || state.running) return;
  setBusy(true); // no second click while checking and confirming
  // The tab may have changed since Preview, so the site is checked again.
  setStatus('Checking the site with SharePoint…');
  let t;
  try {
    t = await checkSite();
  } catch (e) {
    setBusy(false);
    return setStatus(`Site check failed: ${e.message}`);
  }
  if (t.problems.length) {
    state.preview = null;
    setBusy(false);
    return block('The site check failed', t.problems);
  }
  const parsed = state.stages.map((s) => parseBatch(s.body));
  const total = parsed.reduce((n, p) => n + p.operations.length, 0);
  const warnings = parsed.flatMap((p, i) => [...p.warnings, ...analyseBatch(p)].map((w) => `Stage ${i + 1}: ${w}`));
  const confirmed = await ask({
    title: 'Run batches against this site?',
    body: [
      el('p', { class: 'target-title', text: t.info.title }),
      el('p', { class: 'target', text: t.info.webUrl }),
      el('p', { class: 'hint', text: `Site collection: ${t.info.siteCollectionUrl}. Checked with SharePoint just now.` }),
      el('p', { text: `${state.stages.length} stage(s), ${total} operation(s), sent as you (the signed-in user).` }),
      warnings.length
        ? el('details', {}, el('summary', { text: `${warnings.length} warning(s)` }), el('ul', {}, warnings.map((w) => el('li', { text: w }))))
        : null,
    ].filter(Boolean),
    ok: `Run ${total} operation(s)`,
  });
  if (!confirmed) {
    setBusy(false);
    return setStatus('Run cancelled.');
  }

  setBusy(true);
  ui.progress.hidden = false;
  ui.progressBar.value = 0;
  const stages = state.stages.map(({ name, body }) => ({ name, body }));
  state.lastRunIds = state.stages.map((s) => s.id); // for "go to line" in the results
  state.lastRun = { site: t.site, startedAt: new Date().toISOString(), stages: [] };
  const showProgress = onRunnerEvent(stages.length);
  const runner = createRunner({
    request: t.request,
    onEvent: (e) => {
      if (e.type === 'stage-done') state.lastRun.stages.push(e.result);
      showProgress(e);
    },
  });
  try {
    const run = await runner.runStages(t.site, stages, {
      shouldContinue: async (result, i) => {
        setStatus(`Stage ${i + 1} has failures. Waiting for your decision…`);
        return ask({
          title: `Stage ${i + 1} "${result.name}" had failures`,
          body: [...stageOutcome(result, t.site), el('p', { class: 'question', text: `Continue with stage ${i + 2} "${stages[i + 1].name}"?` })],
          ok: 'Continue',
          cancel: 'Stop',
          wide: true,
        });
      },
    });
    state.lastRun = { ...run, siteTitle: t.info.title };
    renderRun(state.lastRun);
    const sum = run.stages.reduce(
      (a, s) => ({ ok: a.ok + (s.summary?.succeeded || 0), warn: a.warn + (s.summary?.warnings || 0), fail: a.fail + (s.summary?.failed || 0) + (s.status === 'error' ? 1 : 0) }),
      { ok: 0, warn: 0, fail: 0 }
    );
    ui.progressBar.value = 1;
    setStatus(`Finished: ${sum.ok} succeeded, ${sum.fail} failed, ${sum.warn} already existed.`);
  } catch (e) {
    setStatus(`Run stopped: ${e.message}`);
  } finally {
    setBusy(false);
  }
}

// ---------------------------------------------------------------- rendering

/** Select a line in a stage's editor and scroll it into view. */
function goToLine(stageId, line) {
  const node = ui.stageList.querySelector(`.stage[data-id="${stageId}"]`);
  if (!node) return setStatus('That stage no longer exists.');
  const ta = node.querySelector('.stage-body');
  const lines = ta.value.split('\n');
  const n = Math.min(Math.max(1, line), lines.length);
  const start = lines.slice(0, n - 1).reduce((sum, l) => sum + l.length + 1, 0);
  node.scrollIntoView({ block: 'nearest' });
  ta.focus();
  ta.setSelectionRange(start, start + lines[n - 1].length);
  const lineHeight = parseFloat(getComputedStyle(ta).lineHeight) || 16;
  ta.scrollTop = Math.max(0, (n - 1) * lineHeight - ta.clientHeight / 3);
  ta.scrollLeft = 0;
}

function lineButton(stageId, line, label = `line ${line}`) {
  const b = el('button', { type: 'button', class: 'line-link', title: `Show line ${line} in the stage` , text: label });
  b.addEventListener('click', () => goToLine(stageId, line));
  return b;
}

/** Text with a "line N" button for every line number it mentions. */
function withLineLinks(text, stageId) {
  const nodes = [el('span', { text })];
  if (stageId == null) return nodes;
  const lines = [...new Set([...String(text).matchAll(/\blines? (\d+)/gi)].map((m) => Number(m[1])))].slice(0, 6);
  if (lines.length) nodes.push(el('span', { class: 'line-links' }, lines.map((l) => lineButton(stageId, l, `→ line ${l}`))));
  return nodes;
}

function opsTable(operations, site, withResult, stageId) {
  const head = el('tr', {},
    el('th', { class: 'num', text: '#' }),
    el('th', { class: 'line', text: 'Line' }),
    el('th', { class: 'method', text: 'Method' }),
    withResult ? el('th', { class: 'code', text: 'Status' }) : null,
    el('th', { text: 'URL' }));
  const rows = operations.map((op) => {
    const cls = withResult ? ({ failed: 'fail', warning: 'warn' })[op.outcome] || '' : '';
    const urlCell = el('td', { class: 'url', title: op.url || '' }, shortUrl(op.url, site));
    if (withResult && op.outcome !== 'success' && op.errorMessage) {
      urlCell.append(el('span', { class: 'err', text: (op.outcome === 'warning' ? 'Already exists: ' : '') + op.errorMessage }));
    }
    if (withResult && op.attempts > 1) urlCell.append(el('span', { class: 'err', text: `Retried (attempt ${op.attempts})` }));
    return el('tr', { class: cls },
      el('td', { class: 'num', text: op.index + 1 }),
      el('td', { class: 'line' }, op.line ? (stageId != null ? lineButton(stageId, op.line, String(op.line)) : String(op.line)) : '—'),
      el('td', { class: 'method', text: op.method || '?' }),
      withResult ? el('td', { class: 'code', text: op.status == null ? '—' : op.status }) : null,
      urlCell);
  });
  return el('table', {}, el('thead', {}, head), el('tbody', {}, rows));
}

function placeholderTable(placeholders, stageIndex) {
  if (!placeholders.length) return null;
  return el('table', {},
    el('thead', {}, el('tr', {}, el('th', { text: 'Placeholder' }), el('th', { text: 'Resolves to' }))),
    el('tbody', {}, placeholders.map((p) => {
      if (p.id) return el('tr', {}, el('td', { text: p.label }), el('td', { class: 'guid', text: p.id }));
      if (p.pending) {
        return el('tr', {}, el('td', { text: p.label }),
          el('td', { class: 'url', text: `Created in this stage at line ${p.pending.line}; looked up right after that part is sent.` }));
      }
      const earlier = stageIndex > 0;
      return el('tr', { class: earlier ? 'warn' : 'fail' },
        el('td', { text: p.label }),
        el('td', { class: 'url' },
          el('span', { class: 'err', text: earlier ? `${p.error} Fine if an earlier stage creates it; it is looked up again before this stage runs.` : p.error })));
    })));
}

function messages(cls, title, items, stageId) {
  if (!items || !items.length) return null;
  return el('div', { class: `msg ${cls}` }, title, el('ul', {}, items.map((t) => el('li', {}, withLineLinks(t, stageId)))));
}

/** A collapsed check section: "▸ Title (count)". */
function check(title, count, cls, ...content) {
  return el('details', { class: 'check' },
    el('summary', {}, el('span', { class: `chip ${cls}`, text: String(count) }), title),
    el('div', { class: 'check-body' }, ...content));
}

function messageList(cls, items, stageId) {
  return el('div', { class: `msg ${cls}` }, el('ul', {}, items.map((t) => el('li', {}, withLineLinks(t, stageId)))));
}

function renderPreview(siteCheck, previews) {
  const site = siteCheck.site;
  const { errorCount, warningCount } = state.preview;
  ui.output.hidden = false;
  ui.outputHead.textContent = siteCheck.info ? `Preview of "${siteCheck.info.title}" (nothing sent)` : 'Preview (nothing sent)';
  ui.copyLog.hidden = ui.exportJson.hidden = true;

  const summary = errorCount
    ? el('p', { class: 'preview-summary fail', text: `Run blocked: ${errorCount} error(s). Expand the sections marked red.` })
    : el('p', { class: 'preview-summary ok', text: `Ready to run.${warningCount ? ` ${warningCount} warning(s), they don't block the run.` : ''}` });

  const siteLines = siteCheck.info
    ? [`Title: ${siteCheck.info.title}`, `Web: ${siteCheck.info.webUrl}`, `Site collection: ${siteCheck.info.siteCollectionUrl}`]
    : [`Site URL: ${site || '(none)'}`];
  const siteSection = el('details', { class: 'stage-out' },
    el('summary', {},
      siteCheck.problems.length ? el('span', { class: 'chip fail', text: 'error' }) : el('span', { class: 'chip ok', text: 'ok' }),
      el('h3', { text: 'Site check' })),
    el('div', { class: 'inner' },
      siteCheck.problems.length ? messageList('fail', siteCheck.problems) : null,
      messageList('info', siteLines)));

  const stageSections = previews.map((p, i) => {
    const stageId = p.stageId;
    const unresolved = p.placeholders.filter((x) => !x.id && !x.pending).length;
    const chip = p.errors.length ? el('span', { class: 'chip fail', text: 'error' })
      : p.warnings.length || unresolved ? el('span', { class: 'chip warn', text: 'check' })
      : el('span', { class: 'chip ok', text: 'ready' });
    const counts = [
      `${p.operations.length} op(s)`,
      p.errors.length ? `${p.errors.length} error(s)` : null,
      p.warnings.length ? `${p.warnings.length} warning(s)` : null,
      p.placeholders.length ? `${p.placeholders.length - unresolved}/${p.placeholders.length} placeholder(s) ready` : null,
      p.parts > 1 ? `${p.parts} parts` : null,
    ].filter(Boolean).join(' · ');
    return el('details', { class: 'stage-out' },
      el('summary', {}, chip, el('h3', { text: `${i + 1}. ${p.name}` }), el('span', { class: 'summary-text', text: counts })),
      el('div', { class: 'inner' },
        p.errors.length ? check('Errors (block the run)', p.errors.length, 'fail', messageList('fail', p.errors, stageId)) : null,
        p.warnings.length ? check('Warnings', p.warnings.length, 'warn', messageList('warn', p.warnings, stageId)) : null,
        p.notes && p.notes.length ? check('Parts', p.parts, 'muted', messageList('info', p.notes, stageId)) : null,
        p.placeholders.length
          ? check('Placeholders', `${p.placeholders.length - unresolved}/${p.placeholders.length}`, unresolved ? 'warn' : 'ok', placeholderTable(p.placeholders, i))
          : null,
        p.operations.length ? check('Operations', p.operations.length, 'muted', opsTable(p.operations, site, false, stageId)) : null));
  });
  ui.outputBody.replaceChildren(summary, siteSection, ...stageSections);
}

function renderRun(run) {
  if (!run) return;
  ui.output.hidden = false;
  ui.outputHead.textContent = 'Results';
  ui.copyLog.hidden = ui.exportJson.hidden = false;
  ui.outputBody.replaceChildren(...run.stages.map((s, i) => {
    const stageId = (state.lastRunIds || [])[i];
    const failed = s.status === 'error' || (s.summary && s.summary.failed > 0);
    const chip = s.status === 'skipped' ? el('span', { class: 'chip muted', text: 'skipped' })
      : failed ? el('span', { class: 'chip fail', text: s.status === 'error' ? 'error' : 'failures' })
      : s.summary && s.summary.warnings ? el('span', { class: 'chip warn', text: 'done' })
      : el('span', { class: 'chip ok', text: 'done' });
    const waits = (s.waits || []).map((w) => `Waited ${Math.round(w.ms / 1000)} s: ${w.reason} (retry ${w.retry} of ${w.of}).`);
    return el('details', { class: 'stage-out', open: failed || s.status === 'error' },
      el('summary', {}, chip, el('h3', { text: `${i + 1}. ${s.name}` }), el('span', { class: 'summary-text', text: stageSummaryText(s) })),
      el('div', { class: 'inner' },
        s.error ? el('div', { class: 'msg fail' }, withLineLinks(s.error, stageId)) : null,
        messages('warn', 'Warnings:', s.warnings, stageId),
        messages('info', 'Parts:', s.notes, stageId),
        messages('info', 'Throttling:', waits),
        placeholderTable((s.placeholders || []).filter((p) => p.id || s.status === 'error'), 0),
        s.operations && s.operations.length ? opsTable(s.operations, run.site, true, stageId) : null));
  }));
}

/**
 * Outcome of one stage for the "continue?" dialog: summary, stage error, and
 * collapsed tables of failed / already existing / succeeded operations.
 */
function stageOutcome(result, site) {
  const ops = result.operations || [];
  const group = (outcome) => ops.filter((op) => op.outcome === outcome);
  const failed = group('failed');
  const existed = group('warning');
  const succeeded = group('success');
  const copy = el('button', { type: 'button', class: 'secondary small', text: "Copy this stage's log" });
  copy.addEventListener('click', async () => {
    const text = formatRunText({ site, startedAt: result.startedAt, finishedAt: result.finishedAt, stages: [result] });
    await copyText(text);
    copy.textContent = 'Copied';
  });
  return [
    el('p', { class: 'summary-line', text: result.summary ? stageSummaryText(result) + '.' : 'Nothing was sent.' }),
    result.error ? el('div', { class: 'msg fail', text: result.error }) : null,
    failed.length ? check('Failed', failed.length, 'fail', opsTable(failed, site, true)) : null,
    existed.length ? check('Already existed (safe to ignore on a re-run)', existed.length, 'warn', opsTable(existed, site, true)) : null,
    succeeded.length ? check('Succeeded', succeeded.length, 'ok', opsTable(succeeded, site, true)) : null,
    el('div', { class: 'row' }, copy),
  ].filter(Boolean);
}

// ---------------------------------------------------------------- log export

async function copyText(text) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    const ta = el('textarea');
    ta.value = text;
    document.body.append(ta);
    ta.select();
    document.execCommand('copy');
    ta.remove();
  }
}

async function copyLog() {
  if (!state.lastRun) return;
  await copyText(formatRunText(state.lastRun));
  setStatus('Log copied to the clipboard.');
}

function exportJson() {
  if (!state.lastRun) return;
  const blob = new Blob([formatRunJson(state.lastRun)], { type: 'application/json' });
  const url = URL.createObjectURL(blob);
  const stamp = new Date().toISOString().replace(/[:.]/g, '-').slice(0, 19);
  const a = el('a', { href: url, download: `sp-batch-run-${stamp}.json` });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

// ---------------------------------------------------------------- wiring

ui.detect.addEventListener('click', () => detectSite());
ui.site.addEventListener('input', () => {
  setHint('Site URL set manually.');
  save();
  updateRunState();
});
ui.site.addEventListener('change', () => {
  ui.site.value = normalizeSiteUrl(ui.site.value);
  save();
});
ui.addStage.addEventListener('click', () => {
  addStage();
  renderStages();
  save();
  ui.stageList.lastElementChild?.querySelector('.stage-body').focus();
});
ui.loadFolder.addEventListener('click', () => ui.folderInput.click());
ui.loadFile.addEventListener('click', () => ui.fileInput.click());
ui.fileInput.addEventListener('change', async () => {
  const f = ui.fileInput.files[0];
  ui.fileInput.value = '';
  if (f) await loadBatchFile(f);
});
ui.folderInput.addEventListener('change', async () => {
  await loadFolder(ui.folderInput.files);
  ui.folderInput.value = '';
});
ui.preview.addEventListener('click', runPreview);
ui.run.addEventListener('click', runAll);
ui.copyLog.addEventListener('click', copyLog);
ui.exportJson.addEventListener('click', exportJson);

load();
