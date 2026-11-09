// Line above the tabs: the list open in the active tab, with its ID ready to copy.

import { el, activeTab } from './shared.js';
import { isSharePointUrl } from './lib/batch.js';

const box = document.getElementById('list-ctx');

/** MAIN world, self-contained: the list SharePoint publishes for the page, else the ?List= parameter of list settings pages. */
function readList() {
  const info = window._spPageContextInfo || {};
  const fromUrl = new URLSearchParams(location.search).get('List');
  const id = String(info.listId || fromUrl || '').replace(/[{}]/g, '').toLowerCase();
  return { id: /^[0-9a-f-]{36}$/.test(id) ? id : null, title: info.listTitle || null };
}

let timer = null;
let last = '';

async function refresh() {
  let list = null;
  try {
    const tab = await activeTab();
    if (tab && tab.url && isSharePointUrl(tab.url)) {
      const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: readList, world: 'MAIN' });
      list = res && res.result && res.result.id ? res.result : null;
    }
  } catch {
    /* page not scriptable or still loading */
  }
  const key = list ? `${list.id}|${list.title}` : '';
  if (key === last) return;
  last = key;
  box.hidden = !list;
  if (!list) return;
  const copy = el('button', { type: 'button', class: 'list-ctx-id', title: 'Copy the list ID', text: list.id });
  copy.addEventListener('click', async () => {
    try {
      await navigator.clipboard.writeText(list.id);
      copy.textContent = 'Copied';
    } catch {
      copy.textContent = 'Copy failed';
    }
    setTimeout(() => { copy.textContent = list.id; }, 1200);
  });
  box.replaceChildren('Open list: ', el('strong', { text: list.title || '(untitled)' }), ' · ', copy);
}

const soon = () => {
  clearTimeout(timer);
  timer = setTimeout(refresh, 300);
};

chrome.tabs.onActivated.addListener(soon);
chrome.tabs.onUpdated.addListener((_id, info) => (info.url || info.status === 'complete') && soon());
window.addEventListener('focus', soon);
refresh();
