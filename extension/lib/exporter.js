// Reads a site's lists, columns, views, content types and permissions for the
// architecture export. Like the batch runner, every request runs in the
// SharePoint tab through `request` (pageRequest); everything here is GET only.

import { extractError, isThrottled, retryAfterMs } from './response.js';
import {
  pickerLists,
  siteRow,
  permissionRows,
  listRow,
  columnRows,
  viewRows,
  contentTypeRows,
  errorRow,
  isCustomField,
} from './export.js';
import { CULTURES } from './languages.js';

export const MAX_RETRIES = 3;

const WEB_SELECT = 'Title,Url,Id,ServerRelativeUrl,WebTemplate,Configuration,Language,SupportedUILanguageIds,Created,Description,HasUniqueRoleAssignments';
const PICKER_SELECT = 'Id,Title,BaseTemplate,BaseType,Hidden,IsCatalog,IsSystemList,ItemCount,HasUniqueRoleAssignments,RootFolder/ServerRelativeUrl';
const LIST_SELECT = [
  'Id', 'Title', 'Description', 'BaseTemplate', 'BaseType', 'Hidden', 'ItemCount', 'EnableVersioning', 'MajorVersionLimit',
  'EnableMinorVersions', 'MajorWithMinorVersionsLimit', 'EnableModeration', 'ForceCheckout', 'ContentTypesEnabled',
  'EnableAttachments', 'EnableFolderCreation', 'NoCrawl', 'HasUniqueRoleAssignments', 'Created', 'LastItemModifiedDate',
  'RootFolder/ServerRelativeUrl', 'DefaultView/ServerRelativeUrl',
].join(',');
const ROLES = 'roleassignments?$expand=Member,RoleDefinitionBindings';

export class ExportError extends Error {}

const defaultSleep = (ms) => new Promise((r) => setTimeout(r, ms));

/**
 * @param {object} deps
 * @param {(req: object) => Promise<object>} deps.request  runs pageRequest in the tab
 * @param {(ms: number) => Promise<void>} [deps.sleep]
 * @param {(event: object) => void} [deps.onEvent]
 */
export function createExporter({ request, sleep = defaultSleep, onEvent = () => {} }) {
  /** GET JSON; waits and retries on 429/503. Throws ExportError on other failures. */
  async function getJson(site, path, extraHeaders = {}) {
    for (let attempt = 0; ; attempt++) {
      const res = await request({ url: site + path, method: 'GET', headers: { Accept: 'application/json;odata=nometadata', ...extraHeaders } });
      if (res.phase === 'network') throw new ExportError(res.statusText);
      if (isThrottled(res.status) && attempt < MAX_RETRIES) {
        const ms = retryAfterMs(res.headers && res.headers['retry-after']);
        onEvent({ type: 'wait', ms, retry: attempt + 1, of: MAX_RETRIES });
        await sleep(ms);
        continue;
      }
      if (res.status < 200 || res.status >= 300) {
        const { message } = extractError(res.text);
        const err = new ExportError(`HTTP ${res.status}${message ? `: ${message}` : ''}`);
        err.status = res.status;
        throw err;
      }
      try {
        return JSON.parse(res.text);
      } catch {
        throw new ExportError('The response was not JSON.');
      }
    }
  }

  /** GET with $select; if SharePoint rejects the select (400), try `fallback`. */
  async function getWithFallback(site, path, fallback) {
    try {
      return await getJson(site, path);
    } catch (e) {
      if (e.status === 400 && fallback) return getJson(site, fallback);
      throw e;
    }
  }

  /** Web, site collection URL and all lists for the picker. */
  async function loadLists(site) {
    const web = await getJson(site, `/_api/web?$select=${WEB_SELECT}`);
    const sc = await getJson(site, '/_api/site?$select=Url');
    const listsJson = await getWithFallback(
      site,
      `/_api/web/lists?$select=${PICKER_SELECT}&$expand=RootFolder&$top=5000`,
      `/_api/web/lists?$select=${PICKER_SELECT.replace('IsSystemList,', '')}&$expand=RootFolder&$top=5000`
    );
    const lists = pickerLists(listsJson);
    return { web, siteCollectionUrl: sc.Url, lists };
  }

  /**
   * Read the selected lists. Returns { sitePermissions, lists: [{ entry, list,
   * permissions, fields, views, contentTypes, errors }], problems }. A part
   * that can't be read is recorded in `errors` (type -> message) and in
   * `problems`, instead of stopping.
   * `parts` limits what is read: { permissions, fields, views, contentTypes }.
   */
  async function collect(site, { lists }, selectedIds, parts = {}) {
    const want = { permissions: true, fields: true, views: true, contentTypes: true, ...parts };
    const problems = [];
    const read = async (path, label, errors, type) => {
      try {
        return await getJson(site, path);
      } catch (e) {
        if (!(e instanceof ExportError)) throw e;
        errors[type] = e.message;
        problems.push(`${label} could not be read (${e.message}).`);
        return null;
      }
    };

    const siteErrors = {};
    const sitePermissions = want.permissions ? await read(`/_api/web/${ROLES}`, 'Site permissions', siteErrors, 'SitePermission') : null;
    const out = [];
    const selected = lists.filter((l) => selectedIds.includes(l.id));
    for (const [i, entry] of selected.entries()) {
      onEvent({ type: 'list', index: i, total: selected.length, title: entry.title });
      const base = `/_api/web/lists(guid'${entry.id}')`;
      const item = { entry, list: null, errors: {} };
      try {
        item.list = await getWithFallback(site, `${base}?$select=${LIST_SELECT}&$expand=RootFolder,DefaultView`, `${base}?$expand=RootFolder`);
      } catch (e) {
        if (!(e instanceof ExportError)) throw e;
        item.errors.List = e.message;
        problems.push(`${entry.title}: the list could not be read (${e.message}).`);
        out.push(item);
        continue;
      }
      const label = (what) => `${entry.title}: ${what}`;
      if (want.permissions) item.permissions = await read(`${base}/${ROLES}`, label('Permissions'), item.errors, 'ListPermission');
      if (want.fields) item.fields = await read(`${base}/fields`, label('Columns'), item.errors, 'ListColumn');
      if (want.views) item.views = await read(`${base}/views?$expand=ViewFields`, label('Views'), item.errors, 'ListView');
      if (want.contentTypes) item.contentTypes = await read(`${base}/contenttypes?$expand=FieldLinks`, label('Content types'), item.errors, 'ListContentType');
      out.push(item);
    }
    onEvent({ type: 'done', total: selected.length });
    return { sitePermissions, siteErrors, lists: out, problems };
  }

  /**
   * Read the translated titles of the lists in `data.lists` (from collect) and of their custom columns, in the languages
   * `langs` (language IDs). SharePoint returns texts in the language of the Accept-Language header, so this is two
   * requests per list and language: the list (title, description) and its columns. Sets `item.translations` =
   * { list: { Title: { lcid: text }, Description }, columns: { InternalName: { Title: { lcid: text } } } }.
   * A language that can't be read is left out.
   */
  async function readTranslations(site, data, langs, { onProgress = () => {} } = {}) {
    const items = data.lists.filter((i) => i.list);
    const total = items.length * langs.length;
    let done = 0;
    for (const item of items) {
      const base = `/_api/web/lists(guid'${item.entry.id}')`;
      const tr = { list: { Title: {}, Description: {} }, columns: {} };
      item.translations = tr;
      const custom = ((item.fields && (item.fields.value || (item.fields.d && item.fields.d.results))) || []).filter(isCustomField);
      for (const f of custom) tr.columns[f.InternalName] = { Title: {} };
      for (const id of langs) {
        const headers = { 'Accept-Language': CULTURES[id] };
        try {
          const list = await getJson(site, `${base}?$select=Title,Description`, headers);
          const fields = await getJson(site, `${base}/fields?$select=InternalName,Title`, headers);
          if (typeof list.Title === 'string') tr.list.Title[id] = list.Title;
          if (typeof list.Description === 'string') tr.list.Description[id] = list.Description;
          for (const f of fields.value || (fields.d && fields.d.results) || []) {
            if (tr.columns[f.InternalName] && typeof f.Title === 'string') tr.columns[f.InternalName].Title[id] = f.Title;
          }
        } catch (e) {
          if (!(e instanceof ExportError)) throw e;
          // this language could not be read for this list: it is not compared
        }
        onProgress(++done, total);
      }
    }
  }

  /** Export the site and the selected lists as CSV rows. Returns { rows, problems }. */
  async function exportSite(site, loaded, selectedIds, { allColumns = false } = {}) {
    const { web, siteCollectionUrl, lists } = loaded;
    const data = await collect(site, loaded, selectedIds);
    const listsById = Object.fromEntries(lists.map((l) => [l.id, { title: l.title, url: l.url }]));
    const rows = [siteRow(web, siteCollectionUrl)];
    if (data.sitePermissions) rows.push(...permissionRows(data.sitePermissions, { type: 'SitePermission', inherited: false }));
    else rows.push(errorRow('SitePermission', null, 'Site permissions', data.siteErrors.SitePermission));

    for (const { entry, list, errors, permissions, fields, views, contentTypes } of data.lists) {
      if (!list) {
        rows.push(errorRow('List', { Title: entry.title, RootFolder: { ServerRelativeUrl: entry.url } }, 'List', errors.List));
        continue;
      }
      rows.push(listRow(list));
      const inherited = list.HasUniqueRoleAssignments === false ? true : list.HasUniqueRoleAssignments === true ? false : undefined;
      const add = (json, type, what, build) => rows.push(...(json ? build(json) : [errorRow(type, list, what, errors[type])]));
      add(permissions, 'ListPermission', 'Permissions', (j) => permissionRows(j, { type: 'ListPermission', list, inherited }));
      add(fields, 'ListColumn', 'Columns', (j) => columnRows(list, j, { allColumns, listsById }));
      add(views, 'ListView', 'Views', (j) => viewRows(list, j));
      add(contentTypes, 'ListContentType', 'Content types', (j) => contentTypeRows(list, j));
    }
    return { rows, problems: data.problems };
  }

  return { loadLists, collect, readTranslations, exportSite };
}
