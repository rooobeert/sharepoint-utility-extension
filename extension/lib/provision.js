// "Export as batches": turns a read site into ready-made provisioning batches
// for the Batch Runner, written as one text file with stage markers.
// Pure functions: no DOM, no chrome.* APIs.
//
// Stages, in order:
//   01 Lists and libraries   POST /_api/web/lists, created under their URL name
//   02 List settings         MERGE versioning, attachments, folders, ...
//   03 Columns               CreateFieldAsXml (cleaned SchemaXml, field IDs kept),
//                            lookups as {{ListId:<site-relative URL>}}, Title column changes
//   04 Views                 default view changed in place, other views created
//   05 Display names         lists and views renamed to their display titles
// Every URL uses {{WebUrl}} / {{WebServerRelativeUrl}}, so the file runs on any site.

import { isCustomField } from './export.js';

export const STAGE_MARKER = '### STAGE:';
export const MAX_OPS_PER_STAGE = 50;

const BATCH = 'batch_provisioning';
const CHANGESET = 'changeset_provisioning';

// Libraries every site has; they are not created again, only extended.
export const DEFAULT_LIBRARIES = new Set(['shared documents', 'sitepages', 'siteassets', 'style library', 'formservertemplates']);

const VIEW_KINDS = { HTML: 1, GRID: 2048 };

const results = (json) => (json && (json.value || (json.d && json.d.results))) || [];
const items = (v) => (Array.isArray(v) ? v : (v && (v.Items || v.results)) || []);
const stripBraces = (g) => String(g || '').replace(/^\{|\}$/g, '').toLowerCase();

// ---------------------------------------------------------------- URLs

/** Path relative to the web: "/sites/x/Lists/crRequests" -> "Lists/crRequests". */
export function webRelative(serverRelativeUrl, webServerRelativeUrl) {
  const web = (webServerRelativeUrl || '').replace(/\/+$/, '');
  const url = String(serverRelativeUrl || '');
  return (web && url.toLowerCase().startsWith(web.toLowerCase() + '/') ? url.slice(web.length + 1) : url.replace(/^\/+/, ''));
}

/** A value inside ('...') in a request line: quotes doubled, no spaces or #. */
const inQuotes = (s) => encodeURI(String(s)).replace(/#/g, '%23').replace(/'/g, "''");

const listEndpoint = (rel) => `{{WebUrl}}/_api/web/getList('{{WebServerRelativeUrl}}/${inQuotes(rel)}')`;

/** The name a list gets its URL from: the last segment of its web-relative URL. */
export const urlName = (rel) => rel.split('/').pop();

// ---------------------------------------------------------------- operations

const VERBOSE = 'application/json;odata=verbose';

function op(method, url, body, { merge = false } = {}) {
  return { method, url, body: body == null ? '' : JSON.stringify(body), merge };
}

/** Field SchemaXml without what ties it to the source site; lookups as placeholders. */
export function cleanFieldXml(xml, listsById, webServerRelativeUrl) {
  const notes = [];
  let out = String(xml)
    .replace(/\s(SourceID|WebId|Version|ColName|RowOrdinal|ColName2|RowOrdinal2)="[^"]*"/g, '');
  out = out.replace(/\sList="([^"]*)"/, (attr, value) => {
    if (/^self$/i.test(value)) return attr;
    const target = listsById[stripBraces(value)];
    if (!target) {
      notes.push(`lookup target ${value} is not a list of this site; left as is`);
      return attr;
    }
    return ` List="{{{ListId:${webRelative(target.url, webServerRelativeUrl)}}}}"`;
  });
  return { xml: out, notes };
}

const TITLE_DEFAULTS = { title: 'Title' };

function listOps(item, ctx) {
  const { list } = item;
  const rel = webRelative(list.RootFolder.ServerRelativeUrl, ctx.webSru);
  const name = urlName(rel);
  const isDefault = DEFAULT_LIBRARIES.has(rel.toLowerCase());
  const isLibrary = list.BaseType === 1 || [101, 109, 119, 851].includes(list.BaseTemplate);
  const endpoint = listEndpoint(rel);
  const ops = { create: [], settings: [], columns: [], views: [], names: [] };

  // 01 create under the URL name, so the URL matches the source
  if (isDefault) ctx.notes.push(`"${list.Title}" (${rel}) exists on every site: not created, only its columns and views are exported.`);
  else {
    ops.create.push(op('POST', '{{WebUrl}}/_api/web/lists', {
      __metadata: { type: 'SP.List' },
      BaseTemplate: list.BaseTemplate,
      Title: name,
      Description: list.Description || '',
    }));
  }

  // 02 settings that the list type supports
  const s = { __metadata: { type: 'SP.List' } };
  const set = (k, v) => v !== undefined && v !== null && (s[k] = v);
  set('EnableVersioning', list.EnableVersioning);
  if (list.EnableVersioning) set('MajorVersionLimit', list.MajorVersionLimit);
  set('EnableModeration', list.EnableModeration);
  set('ContentTypesEnabled', list.ContentTypesEnabled);
  set('EnableFolderCreation', list.EnableFolderCreation);
  set('NoCrawl', list.NoCrawl);
  if (isLibrary) {
    set('EnableMinorVersions', list.EnableMinorVersions);
    if (list.EnableMinorVersions) set('MajorWithMinorVersionsLimit', list.MajorWithMinorVersionsLimit);
    set('ForceCheckout', list.ForceCheckout);
  } else {
    set('EnableAttachments', list.EnableAttachments);
  }
  if (Object.keys(s).length > 1) ops.settings.push(op('MERGE', endpoint, s, { merge: true }));

  // 03 columns
  if (item.fields) {
    for (const f of results(item.fields)) {
      if (f.InternalName === 'Title') {
        const changed = f.Title !== TITLE_DEFAULTS.title || (f.Required === true) !== !isLibrary;
        if (changed) {
          ops.columns.push(op('MERGE', `${endpoint}/fields/getByInternalNameOrTitle('Title')`,
            { __metadata: { type: 'SP.Field' }, Title: f.Title, Required: !!f.Required }, { merge: true }));
        }
        continue;
      }
      if (!isCustomField(f)) continue;
      if (/^TaxonomyFieldType/i.test(f.TypeAsString || '')) {
        ctx.notes.push(`${list.Title}: managed metadata column "${f.Title}" (${f.InternalName}) skipped; create it in SharePoint.`);
        continue;
      }
      if (!f.SchemaXml) continue;
      const { xml, notes } = cleanFieldXml(f.SchemaXml, ctx.listsById, ctx.webSru);
      for (const n of notes) ctx.notes.push(`${list.Title}: column "${f.Title}": ${n}.`);
      ops.columns.push(op('POST', `${endpoint}/fields/CreateFieldAsXml`, {
        parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: 24 },
      }));
    }
  } else if (item.errors && item.errors.ListColumn) {
    ctx.notes.push(`${list.Title}: columns could not be read (${item.errors.ListColumn}); none exported.`);
  }

  // 04 views (and 05 view names)
  if (item.views) {
    for (const v of results(item.views)) {
      if (v.Hidden || v.PersonalView) continue;
      const fields = items(v.ViewFields);
      if (v.DefaultView) {
        const dv = `${endpoint}/defaultView`;
        ops.views.push(op('MERGE', dv, { __metadata: { type: 'SP.View' }, ViewQuery: v.ViewQuery || '', RowLimit: v.RowLimit, Paged: !!v.Paged }, { merge: true }));
        ops.views.push(op('POST', `${dv}/viewfields/removeallviewfields`));
        for (const name of fields) ops.views.push(op('POST', `${dv}/viewfields/addviewfield('${inQuotes(name)}')`));
        continue;
      }
      const kind = VIEW_KINDS[v.ViewType || 'HTML'];
      if (!kind) {
        ctx.notes.push(`${list.Title}: view "${v.Title}" (${v.ViewType}) skipped; only standard and grid views are exported.`);
        continue;
      }
      const viewName = decodeURIComponent(String(v.ServerRelativeUrl || v.Title).split('/').pop()).replace(/\.aspx$/i, '');
      ops.views.push(op('POST', `${endpoint}/views/add`, {
        parameters: {
          __metadata: { type: 'SP.ViewCreationInformation' },
          Title: viewName,
          ViewTypeKind: kind,
          ViewFields: { __metadata: { type: 'Collection(Edm.String)' }, results: fields },
          Query: v.ViewQuery || '',
          RowLimit: v.RowLimit,
          Paged: !!v.Paged,
          PersonalView: false,
          SetAsDefaultView: false,
        },
      }));
      if (v.Scope) {
        ops.views.push(op('MERGE', `${endpoint}/views/getByTitle('${inQuotes(viewName)}')`, { __metadata: { type: 'SP.View' }, Scope: v.Scope }, { merge: true }));
      }
      if (v.Title && v.Title !== viewName) {
        ops.names.push(op('MERGE', `${endpoint}/views/getByTitle('${inQuotes(viewName)}')`, { __metadata: { type: 'SP.View' }, Title: v.Title }, { merge: true }));
      }
    }
  } else if (item.errors && item.errors.ListView) {
    ctx.notes.push(`${list.Title}: views could not be read (${item.errors.ListView}); none exported.`);
  }

  // 05 the list's display name
  if (!isDefault && list.Title !== name) {
    ops.names.unshift(op('MERGE', endpoint, { __metadata: { type: 'SP.List' }, Title: list.Title }, { merge: true }));
  }
  return ops;
}

// ---------------------------------------------------------------- text

function batchText(ops) {
  const lines = [`--${BATCH}`, `Content-Type: multipart/mixed; boundary=${CHANGESET}`, ''];
  for (const o of ops) {
    lines.push(`--${CHANGESET}`, 'Content-Type: application/http', 'Content-Transfer-Encoding: binary', '');
    // No Content-Type without a body: SharePoint would try to read method
    // parameters (e.g. addviewfield('X')) from the empty JSON and fail.
    lines.push(`${o.method} ${o.url} HTTP/1.1`);
    if (o.body) lines.push(`Content-Type: ${VERBOSE}`);
    lines.push('Accept: application/json;odata=nometadata');
    if (o.merge) lines.push('IF-MATCH: *');
    lines.push('', o.body, '');
  }
  lines.push(`--${CHANGESET}--`, `--${BATCH}--`);
  return lines.join('\n');
}

/** Split a stage's operations into chunks of at most `max`, named "03 Columns (1/2)". */
function chunked(name, ops, max) {
  if (!ops.length) return [];
  const n = Math.ceil(ops.length / max);
  return Array.from({ length: n }, (_, i) => ({
    name: n > 1 ? `${name} (${i + 1}/${n})` : name,
    ops: ops.slice(i * max, (i + 1) * max),
  }));
}

/**
 * Build the provisioning file.
 * @param {object} input
 * @param {object} input.web            web from loadLists (Title, Url, ServerRelativeUrl)
 * @param {Array}  input.allLists       picker lists of the site (to resolve lookup GUIDs)
 * @param {Array}  input.lists          collected lists: { entry, list, fields, views, errors }
 * @param {string} [input.version]      extension version for the header
 * @returns {{ text: string, stages: Array<{name, ops}>, notes: string[] }}
 */
export function buildProvisioning({ web, allLists, lists, version = '', date = new Date(), max = MAX_OPS_PER_STAGE }) {
  const ctx = {
    webSru: web.ServerRelativeUrl,
    listsById: Object.fromEntries(allLists.map((l) => [l.id, { title: l.title, url: l.url }])),
    notes: [],
  };
  const all = { create: [], settings: [], columns: [], views: [], names: [] };
  for (const item of lists) {
    if (!item.list) {
      ctx.notes.push(`${item.entry.title}: the list could not be read${item.errors && item.errors.List ? ` (${item.errors.List})` : ''}; not exported.`);
      continue;
    }
    const ops = listOps(item, ctx);
    for (const k of Object.keys(all)) all[k].push(...ops[k]);
  }
  const stages = [
    ...chunked('01 Lists and libraries', all.create, max),
    ...chunked('02 List settings', all.settings, max),
    ...chunked('03 Columns', all.columns, max),
    ...chunked('04 Views', all.views, max),
    ...chunked('05 Display names', all.names, max),
  ];

  const header = [
    '# SharePoint Utility: provisioning batches',
    `# Source: ${web.Title} (${web.Url}), exported ${date.toISOString().slice(0, 10)}${version ? ` with v${version}` : ''}`,
    `# ${lists.filter((l) => l.list).length} list(s), ${stages.length} stage(s), ${stages.reduce((n, s) => n + s.ops.length, 0)} operation(s).`,
    '# Load this file with "Load batch file…" in the Batch Runner: each "### STAGE:" section becomes a stage.',
    '# URLs use {{WebUrl}} / {{WebServerRelativeUrl}} and lookups {{ListId:...}}, so the file runs on any site.',
    ...(ctx.notes.length ? ['#', '# Notes:', ...ctx.notes.map((n) => `#  - ${n}`)] : []),
  ];
  const body = stages.map((s) => `${STAGE_MARKER} ${s.name}\n${batchText(s.ops)}`);
  return { text: [...header, '', ...body.flatMap((b) => [b, ''])].join('\n'), stages, notes: ctx.notes };
}

export function batchFileName(siteTitle, date = new Date()) {
  const safe = String(siteTitle || 'site').replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, '-').slice(0, 60) || 'site';
  return `${safe}-batches-${date.toISOString().slice(0, 10)}.txt`;
}

// ---------------------------------------------------------------- reading the file back

/**
 * Split a file with "### STAGE: <name>" markers into stages. Lines starting
 * with "#" before the first marker are comments. A file without markers is
 * one stage named `fallbackName`.
 */
export function splitStageFile(text, fallbackName = 'Stage') {
  const lines = String(text).replace(/^﻿/, '').split(/\r\n|\r|\n/);
  if (!lines.some((l) => l.startsWith(STAGE_MARKER))) return [{ name: fallbackName, body: String(text) }];
  const stages = [];
  let current = null;
  for (const line of lines) {
    if (line.startsWith(STAGE_MARKER)) {
      current = { name: line.slice(STAGE_MARKER.length).trim() || `Stage ${stages.length + 1}`, lines: [] };
      stages.push(current);
    } else if (current) current.lines.push(line);
  }
  return stages.map((s) => ({ name: s.name, body: s.lines.join('\n').replace(/\s+$/, '') + '\n' }));
}
