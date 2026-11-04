// Reads and writes a site's Quick Launch or top navigation for the
// "Navigation" builder. Like the batch runner, every request runs in the
// SharePoint tab through `request` (pageRequest).

import { extractError, isThrottled, retryAfterMs } from './response.js';
import { NAV_TARGETS, MAX_DEPTH, sharePointUrl } from './navigation.js';

export const MAX_RETRIES = 3;

export class NavError extends Error {}

const defaultSleep = (ms) => new Promise((r) => setTimeout(r, ms));
const results = (json) => (json && (json.value || (json.d && json.d.results))) || [];
const nodePath = (id) => `/_api/web/navigation/GetNodeById(${Number(id)})`;

/**
 * @param {object} deps
 * @param {(req: object) => Promise<object>} deps.request  runs pageRequest in the tab
 * @param {(ms: number) => Promise<void>} [deps.sleep]
 * @param {(event: object) => void} [deps.onEvent]
 */
export function createNavigator({ request, sleep = defaultSleep, onEvent = () => {} }) {
  /** One request; waits and retries on 429/503. Throws NavError on other failures. */
  async function call(site, path, { method = 'GET', body, headers = {} } = {}) {
    const write = method !== 'GET';
    for (let attempt = 0; ; attempt++) {
      const res = await request({
        url: site + path,
        method: write ? 'POST' : 'GET',
        headers: {
          Accept: 'application/json;odata=nometadata',
          ...(body ? { 'Content-Type': 'application/json;odata=verbose' } : {}),
          ...(write && method !== 'POST' ? { 'X-HTTP-Method': method, 'IF-MATCH': '*' } : {}),
          ...headers,
        },
        body: body ? JSON.stringify(body) : undefined,
        digestFrom: write ? site : undefined,
      });
      if (res.phase === 'network') throw new NavError(res.statusText);
      if (isThrottled(res.status) && attempt < MAX_RETRIES) {
        const ms = retryAfterMs(res.headers && res.headers['retry-after']);
        onEvent({ type: 'wait', ms, retry: attempt + 1, of: MAX_RETRIES });
        await sleep(ms);
        continue;
      }
      if (res.status < 200 || res.status >= 300) {
        const { message } = extractError(res.text);
        const what = res.phase === 'contextinfo' ? 'Could not get a form digest' : `HTTP ${res.status}`;
        const err = new NavError(`${what}${message ? `: ${message}` : ''}`);
        err.status = res.status;
        throw err;
      }
      if (!res.text) return null;
      try {
        return JSON.parse(res.text);
      } catch {
        throw new NavError('The response was not JSON.');
      }
    }
  }

  /** Template and navigation settings of the web, for detectNavigation(). */
  async function settings(site) {
    try {
      return await call(site, '/_api/web?$select=WebTemplate,QuickLaunchEnabled,HorizontalQuickLaunch,MegaMenuEnabled');
    } catch (e) {
      if (e.status !== 400) throw e; // older sites may not know every property
      return call(site, '/_api/web?$select=WebTemplate,QuickLaunchEnabled');
    }
  }

  /** The navigation as REST nodes with Children, up to MAX_DEPTH levels. */
  async function load(site, target) {
    const nodes = results(await call(site, NAV_TARGETS[target].path));
    const expand = async (list, depth) => {
      for (const n of list) {
        n.Children = depth < MAX_DEPTH ? results(await call(site, `${nodePath(n.Id)}/Children`)) : [];
        await expand(n.Children, depth + 1);
      }
    };
    await expand(nodes, 1);
    return nodes;
  }

  /**
   * Run a plan from planSave(). Deletes first, then updates, then creates in
   * order. A failure is recorded and the run goes on; children of a node that
   * could not be created are skipped.
   * Returns [{ action, title, ok, error? }].
   */
  async function save(site, target, plan) {
    const out = [];
    const total = plan.deletes.length + plan.updates.length + plan.creates.length;
    let done = 0;
    const step = async (action, title, fn) => {
      onEvent({ type: 'step', index: done++, total, action, title });
      try {
        const value = await fn();
        out.push({ action, title, ok: true });
        return value;
      } catch (e) {
        if (!(e instanceof NavError)) throw e;
        out.push({ action, title, ok: false, error: e.message });
        return undefined;
      }
    };
    const node = (title, url, external) => ({ __metadata: { type: 'SP.NavigationNode' }, Title: title, Url: sharePointUrl(url), ...(external ? { IsExternal: true } : {}) });

    for (const d of plan.deletes) {
      await step('delete', d.title, async () => {
        try {
          await call(site, nodePath(d.id), { method: 'DELETE' });
        } catch (e) {
          if (e.status !== 404) throw e; // already gone
        }
      });
    }
    for (const u of plan.updates) {
      await step('update', u.title, () => call(site, nodePath(u.id), { method: 'MERGE', body: { __metadata: { type: 'SP.NavigationNode' }, Title: u.title, Url: sharePointUrl(u.url) } }));
    }
    const ids = new Map(); // key -> id of a node created in this run
    for (const c of plan.creates) {
      let parentId = null;
      if (c.parent && c.parent.key != null) {
        parentId = ids.get(c.parent.key);
        if (parentId == null) {
          onEvent({ type: 'step', index: done++, total, action: 'create', title: c.title });
          out.push({ action: 'create', title: c.title, ok: false, error: 'Skipped: its parent could not be created.' });
          continue;
        }
      } else if (c.parent) parentId = c.parent.id;
      const path = parentId == null ? NAV_TARGETS[target].path : `${nodePath(parentId)}/Children`;
      const created = await step('create', c.title, () => call(site, path, { method: 'POST', body: node(c.title, c.url, c.external) }));
      const id = created && (created.Id ?? (created.d && created.d.Id));
      if (id != null) ids.set(c.key, id);
    }
    onEvent({ type: 'done', total });
    return out;
  }

  return { settings, load, save };
}
