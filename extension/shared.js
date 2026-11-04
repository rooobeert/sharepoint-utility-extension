// Helpers shared by the side panel tools (Batch Runner, Export Architecture, Navigation).

import { pageRequest, pageReadContext } from './lib/runner.js';
import { isSharePointUrl, isOnOrigin, normalizeSiteUrl, siteUrlFromPageUrl } from './lib/batch.js';

/** Create an element: el('p', { class: 'x', text: 'hi' }, child, ...). */
export function el(tag, props = {}, ...children) {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(props)) {
    if (v == null || v === false) continue;
    if (k === 'class') node.className = v;
    else if (k === 'text') node.textContent = v;
    else node.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat()) {
    if (c == null || c === false) continue;
    node.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return node;
}

export async function activeTab() {
  const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
  return tab || null;
}

/** Web URL of the page in the tab: _spPageContextInfo, else guessed from the URL. */
export async function siteOfTab(tab) {
  let site = null;
  try {
    const [res] = await chrome.scripting.executeScript({ target: { tabId: tab.id }, func: pageReadContext, world: 'MAIN' });
    site = res && res.result && res.result.webAbsoluteUrl;
  } catch {
    /* fall back to the URL */
  }
  return normalizeSiteUrl(site && isSharePointUrl(site) ? site : siteUrlFromPageUrl(tab.url));
}

/** Run pageRequest inside the SharePoint tab, checking it is still on the same tenant and that the request stays on it. */
export function tabRequester(tabId, origin) {
  return async (req) => {
    if (!isOnOrigin(req.url, origin)) throw new Error(`Refused a request to ${req.url}: only ${origin} is allowed.`);
    const tab = await chrome.tabs.get(tabId);
    if (!tab.url || new URL(tab.url).origin !== origin) {
      throw new Error(`The SharePoint tab no longer shows ${origin}. Keep it open on the site while the run is in progress.`);
    }
    const [res] = await chrome.scripting.executeScript({ target: { tabId }, func: pageRequest, args: [req] });
    if (!res || !res.result) throw new Error('The request could not be run in the SharePoint tab.');
    return res.result;
  };
}

const dialogUi = () => ({
  dialog: document.getElementById('dialog'),
  dialogTitle: document.getElementById('dialog-title'),
  dialogBody: document.getElementById('dialog-body'),
  dialogOk: document.getElementById('dialog-ok'),
  dialogCancel: document.getElementById('dialog-cancel'),
});

let pendingAnswer = null; // answers the open dialog, if any

/** Modal dialog. Resolves true for OK, false for Cancel / Escape. */
export function ask({ title, body, ok = 'OK', cancel = 'Cancel', wide = false }) {
  const ui = dialogUi();
  ui.dialog.classList.toggle('wide', wide);
  ui.dialogTitle.textContent = title;
  ui.dialogBody.replaceChildren(...(Array.isArray(body) ? body : [body]).map((b) => (typeof b === 'string' ? el('p', { text: b }) : b)));
  // A dialog that is still waiting is answered with Cancel before the next one.
  if (pendingAnswer) pendingAnswer(false);
  ui.dialogOk.textContent = ok;
  ui.dialogCancel.hidden = cancel == null;
  ui.dialogCancel.textContent = cancel || '';
  ui.dialog.returnValue = '';
  ui.dialog.showModal();
  ui.dialogOk.focus();
  // Resolve from the buttons and Esc only. The 'close' event is not used: Chrome
  // can deliver it late (e.g. while the panel is hidden), and a late close of
  // the previous dialog would otherwise answer the next one.
  return new Promise((resolve) => {
    const done = (answer) => {
      pendingAnswer = null;
      ui.dialogOk.removeEventListener('click', onOk);
      ui.dialogCancel.removeEventListener('click', onCancel);
      ui.dialog.removeEventListener('cancel', onCancel);
      if (ui.dialog.open) ui.dialog.close(answer ? 'ok' : 'cancel');
      resolve(answer);
    };
    const onOk = (e) => {
      e.preventDefault();
      done(true);
    };
    const onCancel = (e) => {
      e.preventDefault();
      done(false);
    };
    ui.dialogOk.addEventListener('click', onOk);
    ui.dialogCancel.addEventListener('click', onCancel);
    ui.dialog.addEventListener('cancel', onCancel); // Esc
    pendingAnswer = done;
  });
}
