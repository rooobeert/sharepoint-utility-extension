// Header switch "Show _api pages as JSON". The rule lives in the browser (dynamic rules
// survive restarts), so nothing is stored by the extension.

import { JSON_RULE_ID, jsonRule } from './lib/json-header.js';

const box = document.getElementById('json-toggle');

async function isOn() {
  const rules = await chrome.declarativeNetRequest.getDynamicRules();
  return rules.some((r) => r.id === JSON_RULE_ID);
}

async function refresh() {
  try {
    box.checked = await isOn();
  } catch {
    box.disabled = true;
  }
}

box.addEventListener('change', async () => {
  box.disabled = true;
  try {
    await chrome.declarativeNetRequest.updateDynamicRules({
      removeRuleIds: [JSON_RULE_ID],
      addRules: box.checked ? [jsonRule()] : [],
    });
  } catch (e) {
    box.title = `Could not change the setting: ${e.message}`;
  }
  box.disabled = false;
  await refresh();
});

refresh();
