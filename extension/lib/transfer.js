// "Compare" -> Batch Runner: turns one difference into a stage that makes the
// target match the source. Pure functions: no DOM, no chrome.* APIs.
//
// A stage is { name, body } like the ones the Batch Runner loads from a file.
// The target's URLs are written into the operations (no {{WebUrl}}), so the
// Batch Runner takes the target site from the batch, not from the active tab.
// Only {{ListId:...}} placeholders for lookups remain; the Batch Runner resolves them.

import { buildProvisioning, batchText, op, cleanFieldXml, listEndpoint } from './provision.js';

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

/** Put the target's URLs in place of the site placeholders. */
export function forTarget(text, targetSite) {
  const site = String(targetSite).replace(/\/+$/, '');
  const path = new URL(site).pathname.replace(/\/+$/, '');
  return text.replace(/\{\{\s*WebUrl\s*\}\}/gi, site).replace(/\{\{\s*WebServerRelativeUrl\s*\}\}/gi, path);
}

const stage = (name, ops, targetSite) => ({ name, body: forTarget(batchText(ops), targetSite) + '\n' });

/** Does the transfer of this list's settings have anything to send? */
export const settingsTransferable = (list) => !!list.item && list.diffs.some((d) => LIST_API[d.prop]);

/** MERGE of the changed list settings, with the source's values. Null if none of them can be changed. */
export function settingsStage(list, targetSite) {
  if (!list.item || !list.item.list) return null;
  const body = { ...SP_LIST };
  for (const d of list.diffs) {
    if (d.blocked || !LIST_API[d.prop]) continue;
    const [name, value] = LIST_API[d.prop](list.item.list);
    if (value !== undefined) body[name] = value;
  }
  if (Object.keys(body).length < 2) return null;
  return stage(`${list.title}: list settings`, [op('MERGE', listEndpoint(list.rel), body, { merge: true })], targetSite);
}

export const columnTransferable = (col) => !!col.field && (col.status === 'missing'
  ? !/^TaxonomyFieldType/i.test(col.type || '') && !!col.field.SchemaXml
  : col.status === 'changed' && col.diffs.some((d) => !d.blocked && COLUMN_API[d.prop]));

/** CreateFieldAsXml for a missing column, MERGE of the changed properties for a changed one. Null if there is nothing to send. */
export function columnStage(list, col, ctx) {
  if (!columnTransferable(col)) return null;
  const endpoint = listEndpoint(list.rel);
  const name = `${list.title}: column "${col.title}"`;
  if (col.status === 'missing') {
    const listsById = Object.fromEntries((ctx.allLists || []).map((l) => [l.id, { title: l.title, url: l.url }]));
    const { xml } = cleanFieldXml(col.field.SchemaXml, listsById, ctx.web.ServerRelativeUrl);
    return stage(name, [op('POST', `${endpoint}/fields/CreateFieldAsXml`, {
      parameters: { __metadata: { type: 'SP.XmlSchemaFieldCreationInformation' }, SchemaXml: xml, Options: 24 },
    })], ctx.targetSite);
  }
  const type = FIELD_TYPES[col.field.TypeAsString] || 'SP.Field';
  const body = { __metadata: { type } };
  for (const d of col.diffs) {
    if (d.blocked || !COLUMN_API[d.prop]) continue;
    const [key, value] = COLUMN_API[d.prop](col.field);
    if (value !== undefined) body[key] = value;
  }
  if (Object.keys(body).length < 2) return null;
  return stage(name, [op('MERGE', `${endpoint}/fields/getByInternalNameOrTitle('${col.name.replace(/'/g, "''")}')`, body, { merge: true })], ctx.targetSite);
}

/** A list missing in the target: create it with its settings and columns (stages as in the architecture export). */
export function listStages(list, ctx) {
  if (!list.item || !list.item.list) return [];
  const { stages } = buildProvisioning({ web: ctx.web, allLists: ctx.allLists || [], lists: [list.item] });
  return stages.map((s) => ({ name: `${list.title}: ${s.name}`, body: forTarget(batchText(s.ops), ctx.targetSite) + '\n' }));
}
