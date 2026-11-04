// SharePoint calls of the "Permissions" tool. Like the exporter, every request
// runs in the SharePoint tab through `request` (pageRequest). Loading, search
// and Preview only read; web/ensureuser and the $batch run only on Apply.

import { extractError, isThrottled, retryAfterMs } from './response.js';
import { createExporter } from './exporter.js';
import { createRunner } from './runner.js';
import {
  roleLevels,
  siteGroups,
  defaultGroupIds,
  assignments,
  pickerPrincipals,
  pickerQuery,
  listOperations,
  hasChanges,
  packBatches,
  buildPermissionBatch,
} from './permissions.js';

export const MAX_RETRIES = 3;

const WEB_SELECT = 'Title,Url,AssociatedOwnerGroup/Id,AssociatedMemberGroup/Id,AssociatedVisitorGroup/Id';
const ROLES_PATH = '/_api/web/roledefinitions?$select=Id,Name,Description,Hidden,RoleTypeKind,Order';
const WEB_EXPAND = 'AssociatedOwnerGroup,AssociatedMemberGroup,AssociatedVisitorGroup';

export class PermissionError extends Error {}

const defaultSleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * @param {object} deps
 * @param {(req: object) => Promise<object>} deps.request  runs pageRequest in the tab
 * @param {(ms: number) => Promise<void>} [deps.sleep]
 * @param {(event: object) => void} [deps.onEvent]
 */
export function createPermissionClient({ request, sleep = defaultSleep, onEvent = () => {} }) {
  /** Send one request and return its JSON; waits and retries on 429/503. */
  async function call(req) {
    for (let attempt = 0; ; attempt++) {
      const res = await request(req);
      if (res.phase === 'network') throw new PermissionError(res.statusText);
      if (isThrottled(res.status) && attempt < MAX_RETRIES) {
        const ms = retryAfterMs(res.headers && res.headers['retry-after']);
        onEvent({ type: 'wait', ms, retry: attempt + 1, of: MAX_RETRIES });
        await sleep(ms);
        continue;
      }
      if (res.status < 200 || res.status >= 300) {
        const { message } = extractError(res.text);
        const where = res.phase === 'contextinfo' ? 'Getting the form digest: ' : '';
        throw new PermissionError(`${where}HTTP ${res.status}${message ? `: ${message}` : ''}`);
      }
      try {
        return JSON.parse(res.text);
      } catch {
        throw new PermissionError('The response was not JSON.');
      }
    }
  }

  const get = (site, path) => call({ url: site + path, method: 'GET', headers: { Accept: 'application/json;odata=nometadata' } });

  /** Web, lists, permission levels and site groups. */
  async function loadSite(site) {
    const [web, { lists }, roles, groups] = [
      await get(site, `/_api/web?$select=${WEB_SELECT}&$expand=${WEB_EXPAND}`),
      await createExporter({ request, sleep, onEvent }).loadLists(site),
      await get(site, ROLES_PATH),
      await get(site, '/_api/web/sitegroups?$select=Id,Title&$top=5000'),
    ];
    return {
      web,
      lists,
      roles: roleLevels(roles),
      groups: siteGroups(groups, web),
      defaults: defaultGroupIds(web),
    };
  }

  /** Permission levels only (after a level was added or changed on the site). */
  async function loadRoles(site) {
    return roleLevels(await get(site, ROLES_PATH));
  }

  /** Search users and groups of the tenant with SharePoint's people picker. */
  async function search(site, text) {
    const json = await call({
      url: `${site}/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser`,
      method: 'POST',
      headers: { Accept: 'application/json;odata=verbose', 'Content-Type': 'application/json;odata=verbose' },
      body: pickerQuery(text),
      digestFrom: site,
    });
    return pickerPrincipals(json);
  }

  /** Current permissions of a list: { unique, assignments, readSecurity, writeSecurity }. */
  async function readList(site, listId) {
    const base = `/_api/web/lists(guid'${listId}')`;
    const list = await get(site, `${base}?$select=HasUniqueRoleAssignments,ReadSecurity,WriteSecurity`);
    const ras = await get(site, `${base}/roleassignments?$expand=Member,RoleDefinitionBindings`);
    return {
      unique: !!list.HasUniqueRoleAssignments,
      assignments: assignments(ras),
      readSecurity: list.ReadSecurity ?? null,
      writeSecurity: list.WriteSecurity ?? null,
    };
  }

  /** Principal id for a login name; adds the principal to the site's users if needed. */
  async function ensure(site, login) {
    const json = await call({
      url: `${site}/_api/web/ensureuser`,
      method: 'POST',
      headers: { Accept: 'application/json;odata=nometadata', 'Content-Type': 'application/json;odata=verbose' },
      body: JSON.stringify({ logonName: login }),
      digestFrom: site,
    });
    if (json.Id == null) throw new PermissionError('SharePoint returned no id.');
    return json.Id;
  }

  /**
   * Apply list plans. 1. Resolve principals without an id (ensureuser). 2. Send
   * the operations as $batch requests of up to 100, whole lists together.
   * Returns { lists: [{ id, title, ops: [{ label, method, url, status, outcome, errorMessage }] }], unresolved }.
   */
  async function apply(site, plans) {
    const ids = new Map();
    const unresolved = [];
    const needed = new Map();
    for (const p of plans) {
      for (const a of p.actions) if (a.type === 'add' && a.principal.id == null) needed.set(a.principal.key, a.principal);
    }
    let n = 0;
    for (const pr of needed.values()) {
      onEvent({ type: 'ensure', index: n++, total: needed.size, title: pr.title });
      try {
        ids.set(pr.key, await ensure(site, pr.login));
      } catch (e) {
        if (!(e instanceof PermissionError)) throw e;
        unresolved.push({ principal: pr, error: e.message });
      }
    }

    // Lists with an unresolved principal still get every other change.
    const idOf = (pr) => pr.id ?? ids.get(pr.key);
    const skipped = new Map(); // list id -> [op]
    const opLists = plans.map((p) => {
      const missing = p.actions.filter((a) => a.type === 'add' && idOf(a.principal) == null);
      if (missing.length) {
        skipped.set(p.list.id, missing.map((a) => ({
          label: `Add ${a.role.name}: ${a.title}`,
          listTitle: p.list.title,
          method: 'POST',
          url: '',
          status: null,
          outcome: 'failed',
          errorMessage: `Not sent: "${a.title}" could not be resolved (${unresolved.find((u) => u.principal.key === a.principal.key).error}).`,
        })));
      }
      const usable = { ...p, actions: p.actions.filter((a) => !missing.includes(a)) };
      const changes = usable.actions.some((a) => a.type === 'add' || a.type === 'remove');
      if (!hasChanges(usable)) return [];
      return listOperations(site, { ...usable, breaks: usable.breaks && changes }, idOf).map((op) => ({ ...op, listId: p.list.id }));
    });

    const batches = packBatches(opLists);
    const stages = batches.map((ops, i) => ({ name: `Permissions ${i + 1} of ${batches.length}`, body: buildPermissionBatch(ops, `perm${i + 1}`) }));
    const runner = createRunner({
      request,
      sleep,
      onEvent: (e) => onEvent({ ...e, total: stages.length }),
    });
    const run = stages.length ? await runner.runStages(site, stages, { shouldContinue: async () => true }) : { stages: [] };

    const byList = new Map(plans.map((p) => [p.list.id, { id: p.list.id, title: p.list.title, ops: [] }]));
    batches.forEach((ops, b) => {
      const st = run.stages[b];
      ops.forEach((op, i) => {
        const r = (st.operations || [])[i];
        byList.get(op.listId).ops.push({
          label: op.label,
          method: op.method,
          url: op.url,
          status: r ? r.status : null,
          outcome: r ? r.outcome : 'failed',
          errorMessage: r ? r.errorMessage : st.error || 'Not sent.',
          attempts: r ? r.attempts : undefined,
        });
      });
    });
    for (const [listId, ops] of skipped) byList.get(listId).ops.push(...ops);
    return { lists: [...byList.values()], unresolved, batches: stages.length };
  }

  return { loadSite, loadRoles, search, readList, ensure, apply };
}
