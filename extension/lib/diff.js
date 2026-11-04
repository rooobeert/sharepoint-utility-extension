// "Compare": the differences between the lists and columns of two sites.
// Pure functions: no DOM, no chrome.* APIs.
//
// Lists are matched by their URL relative to the web ("Lists/crRequests"),
// columns by internal name within a matched list. Each item gets a status:
//   missing  on the source only ("missing in target")
//   only     on the target only
//   changed  on both, with property differences
//   same     on both, no differences
//   error    a side could not be read, so it was not compared

import { csvCell, isCustomField, templateName } from './export.js';
import { webRelative, DEFAULT_LIBRARIES } from './provision.js';

export const STATUSES = ['missing', 'changed', 'only', 'same'];
export const STATUS_LABELS = { missing: 'Missing in target', changed: 'Changed', only: 'Only in target', same: 'Same', error: 'Not compared' };
const STATUS_ORDER = { error: 0, missing: 1, changed: 2, only: 3, same: 4 };

const results = (json) => (json && (json.value || (json.d && json.d.results))) || [];
const items = (v) => (Array.isArray(v) ? v : (v && (v.Items || v.results)) || []);
const stripBraces = (g) => String(g || '').replace(/^\{|\}$/g, '').toLowerCase();
const byTitle = (a, b) => String(a.title).localeCompare(String(b.title), undefined, { sensitivity: 'base' });
const yesNo = (v) => (v === true ? 'Yes' : v === false ? 'No' : '');
const isLibrary = (l) => l.BaseType === 1 || [101, 109, 119, 851].includes(l.BaseTemplate);

// ---------------------------------------------------------------- properties

/**
 * List properties that are compared, as label -> display value. Version
 * limits only count while versioning is on; attachments only for lists,
 * minor versions and check-out only for libraries (as in provision.js).
 */
export function listProps(list) {
  const lib = isLibrary(list);
  const p = {
    Template: templateName(list.BaseTemplate),
    Title: list.Title ?? '',
    Description: list.Description ?? '',
    Versioning: yesNo(list.EnableVersioning),
    'Versioning limit': list.EnableVersioning ? String(list.MajorVersionLimit ?? '') : '',
    'Content approval': yesNo(list.EnableModeration),
    'Content types': yesNo(list.ContentTypesEnabled),
    Folders: yesNo(list.EnableFolderCreation),
    'Hidden from search': yesNo(list.NoCrawl),
  };
  if (lib) {
    p['Minor versions'] = yesNo(list.EnableMinorVersions);
    p['Minor versions limit'] = list.EnableMinorVersions ? String(list.MajorWithMinorVersionsLimit ?? '') : '';
    p['Require check-out'] = yesNo(list.ForceCheckout);
  } else {
    p.Attachments = yesNo(list.EnableAttachments);
  }
  return p;
}

/** List properties SharePoint can't change on an existing list. */
const LIST_BLOCKED = new Set(['Template']);

/**
 * Column properties that are compared, as label -> display value (or an
 * array for choices). `lookupUrl` turns a lookup list GUID into the target
 * list's web-relative URL, so lookups on two sites compare equal.
 */
export function columnProps(f, lookupUrl = () => null) {
  const p = {
    Type: f.TypeAsString ?? '',
    'Display name': f.Title ?? '',
    Required: yesNo(f.Required),
    Description: f.Description ?? '',
    Default: f.DefaultValue ?? '',
    Choices: f.Choices ? items(f.Choices).map(String) : [],
    'Max length': f.MaxLength != null && /^Text$/i.test(f.TypeAsString || '') ? String(f.MaxLength) : '',
    Lines: f.NumberOfLines != null && /^Note$/i.test(f.TypeAsString || '') ? String(f.NumberOfLines) : '',
    'Rich text': /^Note$/i.test(f.TypeAsString || '') ? yesNo(!!f.RichText) : '',
    Indexed: yesNo(!!f.Indexed),
    'Unique values': yesNo(!!f.EnforceUniqueValues),
    'Multiple values': f.AllowMultipleValues != null ? yesNo(!!f.AllowMultipleValues) : '',
    Lookup: f.LookupList ? lookupUrl(f.LookupList) || stripBraces(f.LookupList) : '',
    'Lookup column': f.LookupList ? f.LookupField ?? '' : '',
    Formula: f.Formula ?? '',
  };
  return p;
}

/** Column properties SharePoint can't change safely in place. */
const COLUMN_BLOCKED = new Set(['Type', 'Lookup', 'Multiple values']);
// Properties that only exist for some types; when the type differs, their differences are noise.
const TYPE_SPECIFIC = new Set(['Max length', 'Lines', 'Rich text', 'Choices', 'Formula', 'Lookup', 'Lookup column', 'Multiple values']);

const same = (a, b) => (Array.isArray(a) ? a.length === b.length && a.every((v, i) => v === b[i]) : a === b);

/**
 * Property differences: [{ prop, source, target, added?, removed?, blocked? }].
 * For choices, `added` is what the source has and the target lacks, `removed`
 * what only the target has: what the target would gain and lose.
 */
export function propDiffs(source, target, blocked = new Set()) {
  const out = [];
  for (const prop of Object.keys({ ...source, ...target })) {
    const a = source[prop] ?? (Array.isArray(target[prop]) ? [] : '');
    const b = target[prop] ?? (Array.isArray(source[prop]) ? [] : '');
    if (same(a, b)) continue;
    const d = { prop, source: a, target: b };
    if (Array.isArray(a)) {
      // choices: what was added and removed matters more than the order
      d.added = a.filter((v) => !b.includes(v));
      d.removed = b.filter((v) => !a.includes(v));
    }
    if (blocked.has(prop)) d.blocked = true;
    out.push(d);
  }
  return out;
}

// ---------------------------------------------------------------- sites

/**
 * One side of the comparison, from exporter.loadLists + exporter.collect:
 * { web, allLists, lists: [{ entry, list, fields, errors }] }.
 * Returns lists keyed by lower-case web-relative URL.
 */
function indexSide(side) {
  const webSru = side.web.ServerRelativeUrl;
  const urlById = Object.fromEntries((side.allLists || []).map((l) => [l.id, webRelative(l.url, webSru)]));
  const lookupUrl = (guid) => urlById[stripBraces(guid)] || null;
  const byKey = new Map();
  for (const item of side.lists) {
    const url = (item.list && item.list.RootFolder && item.list.RootFolder.ServerRelativeUrl) || item.entry.url;
    const rel = webRelative(url, webSru);
    byKey.set(rel.toLowerCase(), { ...item, rel, title: (item.list && item.list.Title) || item.entry.title });
  }
  const allKeys = new Set(Object.values(urlById).map((u) => u.toLowerCase()));
  return { byKey, lookupUrl, allKeys };
}

function compareColumns(src, tgt, s, t) {
  const fieldsOf = (item, side) => {
    const m = new Map();
    for (const f of results(item.fields)) m.set(f.InternalName, { f, props: columnProps(f, side.lookupUrl) });
    return m;
  };
  const a = fieldsOf(src, s);
  const b = fieldsOf(tgt, t);
  const out = [];
  for (const name of new Set([...a.keys(), ...b.keys()])) {
    const x = a.get(name);
    const y = b.get(name);
    const f = (x || y).f;
    const col = {
      name,
      title: f.Title,
      type: f.TypeAsString || '',
      custom: isCustomField(f),
      id: stripBraces(f.Id),
    };
    if (x && !y) {
      col.status = 'missing';
      const lookup = x.props.Lookup;
      if (lookup && f.LookupList) {
        col.note = t.allKeys.has(lookup.toLowerCase()) ? `Lookup to ${lookup}` : `Lookup to ${lookup} (missing in target)`;
      }
    } else if (!x && y) {
      col.status = 'only';
    } else {
      col.diffs = propDiffs(x.props, y.props, COLUMN_BLOCKED);
      if (col.diffs.some((d) => d.prop === 'Type')) col.diffs = col.diffs.filter((d) => !TYPE_SPECIFIC.has(d.prop));
      col.status = col.diffs.length ? 'changed' : 'same';
      if (stripBraces(x.f.Id) !== stripBraces(y.f.Id)) col.note = 'Same internal name, different field ID';
    }
    out.push(col);
  }
  return out.sort((p, q) => STATUS_ORDER[p.status] - STATUS_ORDER[q.status] || byTitle(p, q));
}

const columnCount = (item) => results(item.fields).filter(isCustomField).length;

/**
 * Compare two sides. Returns [{ key, rel, title, status, diffs, columns,
 * columnCount, note, error }], sorted by title. A list's own `status` is
 * about the list itself; use listStatus() for the status including columns.
 */
export function compareSites(source, target) {
  const s = indexSide(source);
  const t = indexSide(target);
  const out = [];
  for (const key of new Set([...s.byKey.keys(), ...t.byKey.keys()])) {
    const x = s.byKey.get(key);
    const y = t.byKey.get(key);
    const base = { key, rel: (x || y).rel, title: (x || y).title };

    // Libraries every site has are only compared when both sides have them.
    if (DEFAULT_LIBRARIES.has(key) && !(x && y)) continue;

    const errorOf = (item, side) => {
      if (!item) return null;
      if (item.errors.List) return `the list could not be read in ${side} (${item.errors.List})`;
      if (item.errors.ListColumn) return `columns could not be read in ${side} (${item.errors.ListColumn})`;
      return null;
    };
    const error = errorOf(x, 'source') || errorOf(y, 'target');
    if (error) {
      out.push({ ...base, status: 'error', error, columns: [], diffs: [] });
      continue;
    }

    if (x && !y) out.push({ ...base, status: 'missing', diffs: [], columns: [], columnCount: columnCount(x) });
    else if (!x && y) out.push({ ...base, status: 'only', diffs: [], columns: [], columnCount: columnCount(y) });
    else {
      const diffs = propDiffs(listProps(x.list), listProps(y.list), LIST_BLOCKED);
      out.push({ ...base, title: y.title, status: diffs.length ? 'changed' : 'same', diffs, columns: compareColumns(x, y, s, t) });
    }
  }

  // A missing list and a target-only list with the same title are probably the same list under another URL.
  const onlyByTitle = new Map(out.filter((l) => l.status === 'only').map((l) => [l.title.toLowerCase(), l]));
  const missingByTitle = new Map(out.filter((l) => l.status === 'missing').map((l) => [l.title.toLowerCase(), l]));
  for (const l of out) {
    if (l.status === 'missing' && onlyByTitle.has(l.title.toLowerCase())) l.note = `The target has a list with the same title at ${onlyByTitle.get(l.title.toLowerCase()).rel}`;
    if (l.status === 'only' && missingByTitle.has(l.title.toLowerCase())) l.note = `The source has a list with the same title at ${missingByTitle.get(l.title.toLowerCase()).rel}`;
  }
  return out.sort(byTitle);
}

// ---------------------------------------------------------------- view

/** Columns shown with the "custom columns only" switch. */
export const visibleColumns = (list, { customOnly = true } = {}) => list.columns.filter((c) => !customOnly || c.custom);

/** The list's status including its (visible) columns. */
export function listStatus(list, opts) {
  if (list.status !== 'same') return list.status;
  return visibleColumns(list, opts).some((c) => c.status !== 'same') ? 'changed' : 'same';
}

/**
 * Items per status. A missing or target-only list counts once (its columns
 * aren't counted separately); a matched list counts by its own settings, and
 * each of its visible columns counts by itself.
 */
export function countStatuses(lists, opts) {
  const n = { missing: 0, changed: 0, only: 0, same: 0, error: 0 };
  for (const l of lists) {
    n[l.status]++;
    for (const c of visibleColumns(l, opts)) n[c.status]++;
  }
  return n;
}

/** Column counts of a matched list for its summary line: "2 missing · 3 changed". */
export function columnSummary(list, opts) {
  const n = { missing: 0, changed: 0, only: 0 };
  for (const c of visibleColumns(list, opts)) if (c.status in n) n[c.status]++;
  const parts = [];
  if (n.missing) parts.push(`${n.missing} missing`);
  if (n.changed) parts.push(`${n.changed} changed`);
  if (n.only) parts.push(`${n.only} only in target`);
  return parts.join(' · ');
}

/**
 * Lists and columns shown for the checked statuses and the filter text. A
 * matched list is shown when its settings or one of its shown columns match.
 * Returns [{ list, columns, showSettings }].
 */
export function filterView(lists, { statuses, query = '', customOnly = true } = {}) {
  const want = new Set(statuses);
  want.add('error');
  const q = query.trim().toLowerCase();
  const hit = (...vals) => !q || vals.some((v) => String(v || '').toLowerCase().includes(q));
  const out = [];
  for (const list of lists) {
    const listHit = hit(list.title, list.rel);
    if (list.status === 'missing' || list.status === 'only' || list.status === 'error') {
      if (want.has(list.status) && listHit) out.push({ list, columns: [], showSettings: false });
      continue;
    }
    const columns = visibleColumns(list, { customOnly }).filter((c) => want.has(c.status) && (listHit || hit(c.title, c.name)));
    const showSettings = want.has(list.status) && listHit && list.status === 'changed';
    if (columns.length || showSettings || (want.has('same') && listHit && listStatus(list, { customOnly }) === 'same')) {
      out.push({ list, columns, showSettings });
    }
  }
  return out;
}

// ---------------------------------------------------------------- CSV

export const DIFF_CSV_COLUMNS = ['Status', 'Item', 'List', 'ListUrl', 'Name', 'Title', 'Property', 'Source', 'Target', 'Note'];

const show = (v) => (Array.isArray(v) ? v.join('; ') : v);

/** Rows for the diff CSV, from filterView() output: one row per item or per changed property. */
export function diffRows(view) {
  const rows = [];
  for (const { list, columns, showSettings } of view) {
    const base = { List: list.title, ListUrl: list.rel };
    const status = STATUS_LABELS[list.status];
    if (list.status === 'missing' || list.status === 'only' || list.status === 'error') {
      rows.push({ ...base, Status: status, Item: 'List', Name: list.rel, Title: list.title, Note: list.error || list.note });
      continue;
    }
    if (showSettings) {
      for (const d of list.diffs) {
        rows.push({ ...base, Status: status, Item: 'List', Name: list.rel, Title: list.title, Property: d.prop, Source: show(d.source), Target: show(d.target), Note: d.blocked ? "can't be changed in place" : '' });
      }
    } else if (list.status === 'same' && !columns.length) {
      rows.push({ ...base, Status: STATUS_LABELS.same, Item: 'List', Name: list.rel, Title: list.title });
    }
    for (const c of columns) {
      const col = { ...base, Item: 'Column', Name: c.name, Title: c.title, Status: STATUS_LABELS[c.status] };
      if (c.status === 'changed') {
        for (const d of c.diffs) {
          rows.push({ ...col, Property: d.prop, Source: show(d.source), Target: show(d.target), Note: [d.blocked ? "can't be changed in place" : '', c.note].filter(Boolean).join('; ') });
        }
      } else {
        rows.push({ ...col, Note: c.note });
      }
    }
  }
  return rows;
}

/** CSV text in the same format as the architecture export: BOM, ";" and CRLF. */
export function diffCsv(rows, sep = ';') {
  const lines = [DIFF_CSV_COLUMNS.map((c) => csvCell(c, sep)).join(sep)];
  for (const r of rows) lines.push(DIFF_CSV_COLUMNS.map((c) => csvCell(r[c], sep)).join(sep));
  return '﻿' + lines.join('\r\n') + '\r\n';
}

export function diffFileName(siteTitle, date = new Date()) {
  const safe = String(siteTitle || 'site').replace(/[\\/:*?"<>|]+/g, '').replace(/\s+/g, '-').slice(0, 60) || 'site';
  return `${safe}-compare-${date.toISOString().slice(0, 10)}.csv`;
}
