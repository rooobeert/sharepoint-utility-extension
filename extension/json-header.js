// "JSON" tab: switch and list of URL paths for which SharePoint returns JSON. The switch
// is the dynamic rule itself (kept by the browser); only the path list is stored.

import { el } from './shared.js';
import { JSON_RULE_ID, DEFAULT_PATHS, jsonRule, addPath } from './lib/json-header.js';

const $ = (id) => document.getElementById(id);
const ui = { on: $('json-on'), rows: $('json-paths'), input: $('json-new'), add: $('json-add'), reset: $('json-reset'), msg: $('json-msg') };
const KEY = 'jsonPaths';

let paths = [...DEFAULT_PATHS];

const say = (text, cls = '') => {
  ui.msg.textContent = text;
  ui.msg.className = `hint ${cls}`.trim();
  ui.msg.hidden = !text;
};

async function savePaths() {
  await chrome.storage.local.set({ [KEY]: paths });
}

/** Make the browser's rule match the switch and the list. */
async function applyRule() {
  const rule = jsonRule(paths);
  try {
    await chrome.declarativeNetRequest.updateDynamicRules({
      removeRuleIds: [JSON_RULE_ID],
      addRules: ui.on.checked && rule ? [rule] : [],
    });
    say(ui.on.checked && !rule ? 'The list is empty, so nothing is returned as JSON.' : '', 'error');
  } catch (e) {
    say(`Could not change the setting: ${e.message}`, 'error');
  }
}

function render() {
  ui.rows.replaceChildren(...paths.map((p, i) => {
    const del = el('button', { type: 'button', class: 'icon', title: 'Remove', 'aria-label': `Remove ${p}`, text: '✕' });
    del.addEventListener('click', async () => {
      paths = paths.filter((_, k) => k !== i);
      render();
      await savePaths();
      await applyRule();
    });
    return el('li', { class: 'json-path' }, el('code', { text: p }), del);
  }));
}

async function addFromInput() {
  const next = addPath(paths, ui.input.value);
  if (next === paths) {
    say(ui.input.value.trim() ? 'That path is already in the list, or contains spaces, ? or #.' : '', 'error');
    return;
  }
  paths = next;
  ui.input.value = '';
  render();
  await savePaths();
  await applyRule();
}

ui.on.addEventListener('change', applyRule);
ui.add.addEventListener('click', addFromInput);
ui.input.addEventListener('keydown', (e) => e.key === 'Enter' && addFromInput());
ui.reset.addEventListener('click', async () => {
  paths = [...DEFAULT_PATHS];
  render();
  await savePaths();
  await applyRule();
});

(async () => {
  try {
    const stored = (await chrome.storage.local.get(KEY))[KEY];
    if (Array.isArray(stored)) paths = stored;
    ui.on.checked = (await chrome.declarativeNetRequest.getDynamicRules()).some((r) => r.id === JSON_RULE_ID);
    // A rule from an older version (before the list) is replaced by one for the current list.
    if (ui.on.checked) await applyRule();
  } catch {
    ui.on.disabled = true;
  }
  render();
})();
