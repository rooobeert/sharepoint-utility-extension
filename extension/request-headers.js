// "Headers" tab: switch, list of URL paths and list of request headers. The switch is the
// dynamic rule itself (kept by the browser); only the two lists are stored.

import { el } from './shared.js';
import { RULE_ID, DEFAULT_PATHS, DEFAULT_HEADERS, headerRule, addPath, setHeader, normalizeHeader } from './lib/request-headers.js';

const $ = (id) => document.getElementById(id);
const ui = {
  on: $('hdr-on'), paths: $('hdr-paths'), newPath: $('hdr-new-path'), addPath: $('hdr-add-path'), resetPaths: $('hdr-reset-paths'),
  headers: $('hdr-headers'), name: $('hdr-name'), value: $('hdr-value'), values: $('hdr-values'), addHeader: $('hdr-add-header'),
  resetHeaders: $('hdr-reset-headers'), msg: $('hdr-msg'),
};
const KEY_PATHS = 'jsonPaths'; // the name from when this tab only switched Accept to JSON
const KEY_HEADERS = 'requestHeaders';

// Suggested values for the header being typed.
const VALUES = {
  'accept-language': ['de-DE', 'en-US', 'en-GB', 'fr-FR', 'es-ES', 'it-IT', 'nl-NL', 'pl-PL', 'pt-PT', 'sv-SE', 'da-DK', 'cs-CZ'],
  accept: ['application/json;odata=nometadata', 'application/json;odata=minimalmetadata', 'application/json;odata=verbose', 'application/xml'],
};

let paths = [...DEFAULT_PATHS];
let headers = DEFAULT_HEADERS.map((h) => ({ ...h }));

const say = (text, cls = '') => {
  ui.msg.textContent = text;
  ui.msg.className = `hint ${cls}`.trim();
  ui.msg.hidden = !text;
};

const store = () => chrome.storage.local.set({ [KEY_PATHS]: paths, [KEY_HEADERS]: headers });

/** Make the browser's rule match the switch and the lists. */
async function applyRule() {
  const rule = headerRule(paths, headers);
  try {
    await chrome.declarativeNetRequest.updateDynamicRules({
      removeRuleIds: [RULE_ID],
      addRules: ui.on.checked && rule ? [rule] : [],
    });
    say(ui.on.checked && !rule ? 'A list is empty, so no header is sent.' : '', 'error');
  } catch (e) {
    say(`Could not change the setting: ${e.message}`, 'error');
  }
}

async function changed() {
  render();
  await store();
  await applyRule();
}

function listItem(content, label, onRemove) {
  const del = el('button', { type: 'button', class: 'icon', title: 'Remove', 'aria-label': `Remove ${label}`, text: '✕' });
  del.addEventListener('click', onRemove);
  return el('li', { class: 'hdr-item' }, content, del);
}

function render() {
  ui.paths.replaceChildren(...paths.map((p, i) => listItem(el('code', { text: p }), p, () => { paths = paths.filter((_, k) => k !== i); changed(); })));
  ui.headers.replaceChildren(...headers.map((h, i) => listItem(
    el('code', {}, el('strong', { text: h.name }), `: ${h.value}`), h.name, () => { headers = headers.filter((_, k) => k !== i); changed(); })));
}

function fillSuggestions() {
  ui.values.replaceChildren(...(VALUES[ui.name.value.trim().toLowerCase()] || []).map((v) => el('option', { value: v })));
}

ui.on.addEventListener('change', applyRule);

async function addPathFromInput() {
  const next = addPath(paths, ui.newPath.value);
  if (next === paths) {
    say(ui.newPath.value.trim() ? 'That path is already in the list, or contains spaces, ? or #.' : '', 'error');
    return;
  }
  paths = next;
  ui.newPath.value = '';
  await changed();
}
ui.addPath.addEventListener('click', addPathFromInput);
ui.newPath.addEventListener('keydown', (e) => e.key === 'Enter' && addPathFromInput());
ui.resetPaths.addEventListener('click', () => { paths = [...DEFAULT_PATHS]; changed(); });

async function addHeaderFromInput() {
  if (!normalizeHeader(ui.name.value, ui.value.value)) {
    say('Enter a header name (letters, digits and - only) and a value.', 'error');
    return;
  }
  headers = setHeader(headers, ui.name.value, ui.value.value);
  ui.name.value = '';
  ui.value.value = '';
  fillSuggestions();
  await changed();
}
ui.addHeader.addEventListener('click', addHeaderFromInput);
ui.value.addEventListener('keydown', (e) => e.key === 'Enter' && addHeaderFromInput());
ui.name.addEventListener('input', fillSuggestions);
ui.resetHeaders.addEventListener('click', () => { headers = DEFAULT_HEADERS.map((h) => ({ ...h })); changed(); });

(async () => {
  try {
    const stored = await chrome.storage.local.get([KEY_PATHS, KEY_HEADERS]);
    if (Array.isArray(stored[KEY_PATHS])) paths = stored[KEY_PATHS];
    if (Array.isArray(stored[KEY_HEADERS])) headers = stored[KEY_HEADERS];
    ui.on.checked = (await chrome.declarativeNetRequest.getDynamicRules()).some((r) => r.id === RULE_ID);
    // A rule from an older version is replaced by one for the current lists.
    if (ui.on.checked) await applyRule();
  } catch {
    ui.on.disabled = true;
  }
  render();
})();
