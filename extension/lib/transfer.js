// "Compare" -> Batch Runner: turns differences into stages that make the target match
// the source. Pure functions: no DOM, no chrome.* APIs.
//
// A stage is { name, body } like the ones the Batch Runner loads from a file. URLs keep the
// {{WebUrl}} / {{WebServerRelativeUrl}} placeholders (and {{ListId:...}} for lookups), so the
// stages run on whichever site the Batch Runner is set to; the Compare tab sets it to the target.

import { buildProvisioning, batchText, op, cleanFieldXml, listEndpoint, MAX_OPS_PER_STAGE } from './provision.js';

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

/** MERGE of the changed list settings, with the source's values; null if none of them can be changed. */
function settingsOp(list) {
  if (!list.item || !list.item.list) return null;
  const body = { ...SP_LIST };
  for (const d of list.diffs) {
    if (d.blocked || !LIST_API[d.prop]) continue;
    const [name, value] = LIST_API[d.prop](list.item.list);
    if (value !== undefined) body[name] = value;
  }
  return Object.keys(body).length < 2 ? null : op('MERGE', listEndpoint(list.rel), body, { merge: true });
}

export const settingsTransferable = (list) => !!settingsOp(list);

export const columnTransferable = (col) => !!col.field && (col.status === 'missing'
  ? !/^TaxonomyFieldType/i.test(col.type || '') && !!col.field.SchemaXml
  : col.status === 'changed' && col.diffs.some((d) => !d.blocked && COLUMN_API[d.prop]));

/** CreateFieldAsXml for a missing column, MERGE of the changed properties for a changed one; null if there is nothing to send. */
function columnOp(list, col, ctx) {
  if (!columnTransferable(col)) return null;
  const endpoint = listEndpoint(list.rel);
  if (col.status === 'missing') {
    const listsById = Object.fromEntries((ctx.allLists || []).map((l) => [l.id, { title: l.title, url: l.url }]));
    const { xml } = cleanFieldXml(col.field.SchemaXml, listsById, ctx.web.ServerRelativeUrl);
    return op('POST', `${endpoint}/fields/CreateFieldAsXml`, {
      parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: 24 },
    });
  }
  const body = { __metadata: { type: FIELD_TYPES[col.field.TypeAsString] || 'SP.Field' } };
  for (const d of col.diffs) {
    if (d.blocked || !COLUMN_API[d.prop]) continue;
    const [key, value] = COLUMN_API[d.prop](col.field);
    if (value !== undefined) body[key] = value;
  }
  if (Object.keys(body).length < 2) return null;
  return op('MERGE', `${endpoint}/fields/getByInternalNameOrTitle('${col.name.replace(/'/g, "''")}')`, body, { merge: true });
}

// ---------------------------------------------------------------- stages

export function settingsStage(list) {
  const o = settingsOp(list);
  return o ? stage(`${list.title}: list settings`, [o]) : null;
}

export function columnStage(list, col, ctx) {
  const o = columnOp(list, col, ctx);
  return o ? stage(`${list.title}: column "${col.title}"`, [o]) : null;
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

/** The changes of a view (filterView output) that can be sent: [{ id }]. */
export function changeItems(view) {
  const out = [];
  for (const { list, columns, showSettings } of view) {
    if (list.status === 'missing') {
      if (list.item && list.item.list) out.push({ id: listChangeId(list) });
      continue;
    }
    if (showSettings && settingsTransferable(list)) out.push({ id: settingsChangeId(list) });
    for (const c of columns) if (columnTransferable(c)) out.push({ id: columnChangeId(list, c) });
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
  const settings = [];
  const columns = [];
  for (const { list, columns: cols, showSettings } of view) {
    if (list.status === 'missing') continue; // its columns come with the list
    if (showSettings && pick(settingsChangeId(list))) {
      const o = settingsOp(list);
      if (o) settings.push(o);
    }
    for (const c of cols) {
      if (!pick(columnChangeId(list, c))) continue;
      const o = columnOp(list, c, ctx);
      if (o) columns.push(o);
    }
  }
  return [...provisionStages(missing, ctx), ...stagesOf('List settings', settings), ...stagesOf('Columns', columns)];
}
