// "Compare" -> Batch Runner: turns differences into stages that make the target match
// the source. Pure functions: no DOM, no chrome.* APIs.
//
// A stage is { name, body } like the ones the Batch Runner loads from a file. URLs keep the
// {{WebUrl}} / {{WebServerRelativeUrl}} placeholders (and {{ListId:...}} for lookups), so the
// stages run on whichever site the Batch Runner is set to; the Compare tab sets it to the target.

import { buildProvisioning, batchText, op, cleanFieldXml, listEndpoint, MAX_OPS_PER_STAGE } from './provision.js';
import { CULTURES, parseTranslationLabel } from './languages.js';

const items = (v) => (Array.isArray(v) ? v : (v && (v.Items || v.results)) || []);
const SP_LIST = { __metadata: { type: 'SP.List' } };

/** List properties (labels from diff.js listProps) -> [REST name, value from the source list]. */
const LIST_API = {
  Title: (l) => ['Title', l.Title],
  Description: (l) => ['Description', l.Description || ''],
  Versioning: (l) => ['EnableVersioning', !!l.EnableVersioning],
  'Versioning limit': (l) => ['MajorVersionLimit', l.MajorVersionLimit],
  'Content approval': (l) => ['EnableModeration', !!l.EnableModeration],
  'Content types': (l) => ['ContentTypesEnabled', !!l.ContentTypesEnabled],
  Folders: (l) => ['EnableFolderCreation', !!l.EnableFolderCreation],
  'Hidden from search': (l) => ['NoCrawl', !!l.NoCrawl],
  'Minor versions': (l) => ['EnableMinorVersions', !!l.EnableMinorVersions],
  'Minor versions limit': (l) => ['MajorWithMinorVersionsLimit', l.MajorWithMinorVersionsLimit],
  'Require check-out': (l) => ['ForceCheckout', !!l.ForceCheckout],
  Attachments: (l) => ['EnableAttachments', !!l.EnableAttachments],
};

/** Column properties -> [REST name, value from the source field]. Type, lookup and multiple values can't be changed in place. */
const COLUMN_API = {
  'Display name': (f) => ['Title', f.Title],
  Required: (f) => ['Required', !!f.Required],
  Description: (f) => ['Description', f.Description || ''],
  Default: (f) => ['DefaultValue', f.DefaultValue ?? null],
  Choices: (f) => ['Choices', { __metadata: { type: 'Collection(Edm.String)' }, results: items(f.Choices).map(String) }],
  'Max length': (f) => ['MaxLength', f.MaxLength],
  Lines: (f) => ['NumberOfLines', f.NumberOfLines],
  'Rich text': (f) => ['RichText', !!f.RichText],
  Indexed: (f) => ['Indexed', !!f.Indexed],
  'Unique values': (f) => ['EnforceUniqueValues', !!f.EnforceUniqueValues],
  Formula: (f) => ['Formula', f.Formula],
};

const FIELD_TYPES = { Text: 'SP.FieldText', Note: 'SP.FieldMultiLineText', Choice: 'SP.FieldChoice', MultiChoice: 'SP.FieldMultiChoice', Calculated: 'SP.FieldCalculated', Number: 'SP.FieldNumber' };

const stage = (name, ops) => ({ name, body: batchText(ops) + '\n' });

/** Operations of `ops` in stages of at most MAX_OPS_PER_STAGE, named "name (1/2)". */
function stagesOf(name, ops) {
  if (!ops.length) return [];
  const n = Math.ceil(ops.length / MAX_OPS_PER_STAGE);
  return Array.from({ length: n }, (_, i) => stage(n > 1 ? `${name} (${i + 1}/${n})` : name, ops.slice(i * MAX_OPS_PER_STAGE, (i + 1) * MAX_OPS_PER_STAGE)));
}

// ---------------------------------------------------------------- operations

const escQuote = (s) => String(s).replace(/'/g, "''");
const fieldEndpoint = (list, name) => `${listEndpoint(list.rel)}/fields/getByInternalNameOrTitle('${escQuote(name)}')`;

/**
 * Translated titles / descriptions: SetValueForUICulture for each, then one MERGE of the item to save them.
 * `values` is [{ resource: 'TitleResource' | 'DescriptionResource', culture, value }].
 */
function translationOps(endpoint, type, values) {
  if (!values.length) return [];
  return [
    ...values.map((v) => op('POST', `${endpoint}/${v.resource}/SetValueForUICulture`, { cultureName: v.culture, value: v.value ?? '' })),
    op('MERGE', endpoint, { __metadata: { type } }, { merge: true }),
  ];
}

/** The changed translated properties of a diff list: [{ resource, culture, value }] with the source's text. */
function translationValues(diffs) {
  const out = [];
  for (const d of diffs || []) {
    const t = parseTranslationLabel(d.prop);
    if (t && !d.blocked) out.push({ resource: `${t.what}Resource`, culture: t.culture, value: d.source });
  }
  return out;
}

/** Every translated text a source item has: from { Title: { lcid: text }, Description: { ... } }. */
function allTranslationValues(tr) {
  const out = [];
  for (const what of ['Title', 'Description']) {
    for (const [id, text] of Object.entries((tr && tr[what]) || {})) if (text && CULTURES[id]) out.push({ resource: `${what}Resource`, culture: CULTURES[id], value: text });
  }
  return out;
}

/** MERGE of the changed list settings, with the source's values, and the changed translations; [] if none of them can be changed. */
function settingsOps(list) {
  if (!list.item || !list.item.list) return [];
  const body = { ...SP_LIST };
  for (const d of list.diffs) {
    if (d.blocked || !LIST_API[d.prop]) continue;
    const [name, value] = LIST_API[d.prop](list.item.list);
    if (value !== undefined) body[name] = value;
  }
  const merge = Object.keys(body).length < 2 ? [] : [op('MERGE', listEndpoint(list.rel), body, { merge: true })];
  return [...merge, ...translationOps(listEndpoint(list.rel), 'SP.List', translationValues(list.diffs))];
}

export const settingsTransferable = (list) => settingsOps(list).length > 0;

const canChange = (d) => !d.blocked && (!!COLUMN_API[d.prop] || !!parseTranslationLabel(d.prop));

export const columnTransferable = (col) => !!col.field && (col.status === 'missing'
  ? !/^TaxonomyFieldType/i.test(col.type || '') && !!col.field.SchemaXml
  : col.status === 'changed' && col.diffs.some(canChange));

/** CreateFieldAsXml (and its translated titles) for a missing column, MERGE of the changed properties and translations for a changed one; [] if there is nothing to send. */
function columnOps(list, col, ctx) {
  if (!columnTransferable(col)) return [];
  const endpoint = listEndpoint(list.rel);
  const type = FIELD_TYPES[col.field.TypeAsString] || 'SP.Field';
  if (col.status === 'missing') {
    const listsById = Object.fromEntries((ctx.allLists || []).map((l) => [l.id, { title: l.title, url: l.url }]));
    const { xml } = cleanFieldXml(col.field.SchemaXml, listsById, ctx.web.ServerRelativeUrl);
    return [
      op('POST', `${endpoint}/fields/CreateFieldAsXml`, {
        parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: 24 },
      }),
      ...translationOps(fieldEndpoint(list, col.name), type, allTranslationValues(col.translations ? { Title: col.translations.Title } : null)),
    ];
  }
  const body = { __metadata: { type } };
  for (const d of col.diffs) {
    if (d.blocked || !COLUMN_API[d.prop]) continue;
    const [key, value] = COLUMN_API[d.prop](col.field);
    if (value !== undefined) body[key] = value;
  }
  const merge = Object.keys(body).length < 2 ? [] : [op('MERGE', fieldEndpoint(list, col.name), body, { merge: true })];
  return [...merge, ...translationOps(fieldEndpoint(list, col.name), type, translationValues(col.diffs))];
}

// ---------------------------------------------------------------- only in the target

/** What to do with lists and columns that only the target has: 'rename' (prefix the title) or 'delete' (lists go to the recycle bin). */
export const ONLY_MODES = ['leave', 'rename', 'delete'];
export const DEFAULT_PREFIX = 'zz_';

const prefixed = (title, prefix) => !prefix || String(title).startsWith(prefix);

/** The operation for a list that only the target has, or null. */
function onlyListOp(list, { mode, prefix }) {
  if (list.status !== 'only') return null;
  if (mode === 'delete') return op('POST', `${listEndpoint(list.rel)}/recycle`, null);
  if (mode === 'rename' && prefix && !prefixed(list.title, prefix)) return op('MERGE', listEndpoint(list.rel), { ...SP_LIST, Title: prefix + list.title }, { merge: true });
  return null;
}

/** The operation for a custom column that only the target has, or null. */
function onlyColumnOp(list, col, { mode, prefix }) {
  if (col.status !== 'only' || !col.custom) return null;
  if (mode === 'delete') return op('DELETE', fieldEndpoint(list, col.name), null, { merge: true });
  if (mode === 'rename' && prefix && !prefixed(col.title, prefix)) return op('MERGE', fieldEndpoint(list, col.name), { __metadata: { type: 'SP.Field' }, Title: prefix + col.title }, { merge: true });
  return null;
}

export const onlyListTransferable = (list, only) => !!onlyListOp(list, only);
export const onlyColumnTransferable = (list, col, only) => !!onlyColumnOp(list, col, only);

// ---------------------------------------------------------------- stages

export function settingsStage(list) {
  const ops = settingsOps(list);
  return ops.length ? stage(`${list.title}: list settings`, ops) : null;
}

export function columnStage(list, col, ctx) {
  const ops = columnOps(list, col, ctx);
  return ops.length ? stage(`${list.title}: column "${col.title}"`, ops) : null;
}

/** A list missing in the target: create it with its settings and columns (stages as in the architecture export). */
export function listStages(list, ctx) {
  return provisionStages([list], ctx, `${list.title}: `);
}

function provisionStages(lists, ctx, prefix = '') {
  const withData = lists.filter((l) => l.item && l.item.list);
  if (!withData.length) return [];
  const { stages } = buildProvisioning({ web: ctx.web, allLists: ctx.allLists || [], lists: withData.map((l) => l.item) });
  return stages.map((s) => ({ name: `${prefix}${s.name}`, body: batchText(s.ops) + '\n' }));
}

// ---------------------------------------------------------------- choosing changes

/** Ids of the changes that can be sent: a missing list, a list's changed settings, a missing or changed column. */
export const listChangeId = (list) => `L:${list.key}`;
export const settingsChangeId = (list) => `S:${list.key}`;
export const columnChangeId = (list, col) => `C:${list.key}:${col.name}`;
/** A list or column that only the target has (renamed or deleted, see ONLY_MODES). */
export const onlyListId = (list) => `O:${list.key}`;
export const onlyColumnId = (list, col) => `OC:${list.key}:${col.name}`;

/** The changes of a view (filterView output) that can be sent: [{ id }]. `only` = { mode, prefix } for what only the target has. */
export function changeItems(view, only = { mode: 'leave' }) {
  const out = [];
  for (const { list, columns, showSettings } of view) {
    if (list.status === 'only') {
      if (onlyListTransferable(list, only)) out.push({ id: onlyListId(list) });
      continue;
    }
    if (list.status === 'missing') {
      if (list.item && list.item.list) out.push({ id: listChangeId(list) });
      continue;
    }
    if (showSettings && settingsTransferable(list)) out.push({ id: settingsChangeId(list) });
    for (const c of columns) {
      if (columnTransferable(c)) out.push({ id: columnChangeId(list, c) });
      else if (onlyColumnTransferable(list, c, only)) out.push({ id: onlyColumnId(list, c) });
    }
  }
  return out;
}

/**
 * The changes in the shown view (filterView output) as stages: missing lists first (as in the
 * architecture export), then changed list settings, then missing and changed columns.
 * `pick(id)` leaves out the changes it returns false for (default: all).
 */
export function allStages(view, ctx, pick = () => true) {
  const missing = view.filter((v) => v.list.status === 'missing' && pick(listChangeId(v.list))).map((v) => v.list);
  const only = ctx.only || { mode: 'leave' };
  const settings = [];
  const columns = [];
  const names = [];
  const leftovers = [];
  // the translated titles of the missing lists and their columns (the lists are created first)
  for (const list of missing) {
    const tr = list.item && list.item.translations;
    if (!tr) continue;
    names.push(...translationOps(listEndpoint(list.rel), 'SP.List', allTranslationValues(tr.list)));
    for (const [name, col] of Object.entries(tr.columns || {})) {
      names.push(...translationOps(fieldEndpoint(list, name), 'SP.Field', allTranslationValues({ Title: col.Title })));
    }
  }
  for (const { list, columns: cols, showSettings } of view) {
    if (list.status === 'only') {
      const o = pick(onlyListId(list)) ? onlyListOp(list, only) : null;
      if (o) leftovers.push(o);
      continue;
    }
    if (list.status === 'missing') continue; // its columns come with the list
    if (showSettings && pick(settingsChangeId(list))) settings.push(...settingsOps(list));
    for (const c of cols) {
      if (pick(columnChangeId(list, c))) columns.push(...columnOps(list, c, ctx));
      if (pick(onlyColumnId(list, c))) {
        const o = onlyColumnOp(list, c, only);
        if (o) leftovers.push(o);
      }
    }
  }
  const leftoverName = only.mode === 'delete' ? 'Delete what only the target has' : 'Rename what only the target has';
  return [
    ...provisionStages(missing, ctx), ...stagesOf('List settings', settings), ...stagesOf('Columns', columns),
    ...stagesOf('Translated titles', names), ...stagesOf(leftoverName, leftovers),
  ];
}
