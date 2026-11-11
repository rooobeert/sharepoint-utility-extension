import { test } from 'node:test';
import assert from 'node:assert/strict';
import { settingsStage, columnStage, columnTransferable, settingsTransferable, listStages, allStages, changeItems, listChangeId, settingsChangeId, columnChangeId } from '../extension/lib/transfer.js';

const ctx = { web: { ServerRelativeUrl: '/sites/dev', Title: 'Dev', Url: 'https://t.sharepoint.com/sites/dev' }, allLists: [] };

test('settingsStage merges the source values of the changed settings, with placeholders', () => {
  const list = { title: 'Req', rel: 'Lists/Req', item: { list: { Title: 'Req', EnableVersioning: true, MajorVersionLimit: 50 } },
    diffs: [{ prop: 'Versioning', source: 'Yes', target: 'No' }, { prop: 'Template', blocked: true }] };
  assert.ok(settingsTransferable(list));
  const st = settingsStage(list);
  assert.match(st.body, /^MERGE \{\{WebUrl\}\}\/_api\/web\/getList\('\{\{WebServerRelativeUrl\}\}\/Lists\/Req'\)/m);
  assert.match(st.body, /"EnableVersioning":true/);
  assert.doesNotMatch(st.body, /BaseTemplate|"Template"|sharepoint\.com/);
});

test('columnStage: changed column merges typed properties, blocked ones are left out', () => {
  const list = { title: 'Req', rel: 'Lists/Req' };
  const col = { name: 'Notes', title: 'Notes', type: 'Note', status: 'changed', field: { TypeAsString: 'Note', Title: 'Notes', NumberOfLines: 6, Required: true },
    diffs: [{ prop: 'Lines', source: '6', target: '3' }, { prop: 'Required', source: 'Yes', target: 'No' }, { prop: 'Type', blocked: true }] };
  const st = columnStage(list, col, ctx);
  assert.match(st.body, /fields\/getByInternalNameOrTitle\('Notes'\)/);
  assert.match(st.body, /SP\.FieldMultiLineText/);
  assert.match(st.body, /"NumberOfLines":6/);
  assert.match(st.body, /"Required":true/);
  assert.match(st.body, /\{\{WebUrl\}\}/);
});

test('columnStage: missing column is created from its cleaned schema', () => {
  const col = { name: 'Foo', title: 'Foo', type: 'Text', status: 'missing', field: { SchemaXml: '<Field Name="Foo" SourceID="x" Type="Text"/>' } };
  const st = columnStage({ title: 'Req', rel: 'Lists/Req' }, col, ctx);
  assert.match(st.body, /POST .*fields\/CreateFieldAsXml/);
  assert.doesNotMatch(st.body, /SourceID/);
});

test('nothing to send for managed metadata or blocked-only changes', () => {
  assert.ok(!columnTransferable({ status: 'missing', type: 'TaxonomyFieldType', field: { SchemaXml: '<Field/>' } }));
  assert.ok(!columnTransferable({ status: 'changed', field: {}, diffs: [{ prop: 'Type', blocked: true }] }));
});

test('allStages: missing lists, then settings, then columns, all with placeholders', () => {
  const missingList = { title: 'New', rel: 'Lists/New', status: 'missing', diffs: [],
    item: { entry: { title: 'New' }, list: { Title: 'New', BaseTemplate: 100, BaseType: 0, RootFolder: { ServerRelativeUrl: '/sites/dev/Lists/New' } }, fields: { value: [] } } };
  const changedList = { title: 'Req', rel: 'Lists/Req', status: 'changed', item: { list: { Title: 'Req', EnableVersioning: true } }, diffs: [{ prop: 'Versioning', source: 'Yes', target: 'No' }] };
  const col = { name: 'Foo', title: 'Foo', type: 'Text', status: 'missing', field: { SchemaXml: '<Field Name="Foo" Type="Text"/>' } };
  const view = [
    { list: changedList, columns: [col], showSettings: true },
    { list: missingList, columns: [], showSettings: false },
  ];
  const stages = allStages(view, ctx);
  assert.deepEqual(stages.map((s) => s.name), ['01 Lists and libraries', 'List settings', 'Columns']);
  assert.ok(stages.every((s) => !/sharepoint\.com/.test(s.body)));
  assert.match(stages[0].body, /POST \{\{WebUrl\}\}\/_api\/web\/lists/);
  assert.deepEqual(allStages([{ list: { title: 'Gone', status: 'only', diffs: [] }, columns: [], showSettings: false }], ctx), []);
});

test('large sets are split into stages of at most 50 operations', () => {
  const cols = Array.from({ length: 60 }, (_, i) => ({ name: `c${i}`, title: `C${i}`, type: 'Text', status: 'missing', field: { SchemaXml: `<Field Name="c${i}"/>` } }));
  const stages = allStages([{ list: { title: 'L', rel: 'Lists/L', status: 'changed', diffs: [] }, columns: cols, showSettings: false }], ctx);
  assert.deepEqual(stages.map((s) => s.name), ['Columns (1/2)', 'Columns (2/2)']);
});

test('changeItems lists what can be sent; allStages honours the selection', () => {
  const missingList = { key: 'lists/new', title: 'New', rel: 'Lists/New', status: 'missing', diffs: [],
    item: { entry: { title: 'New' }, list: { Title: 'New', BaseTemplate: 100, BaseType: 0, RootFolder: { ServerRelativeUrl: '/sites/dev/Lists/New' } }, fields: { value: [] } } };
  const changedList = { key: 'lists/req', title: 'Req', rel: 'Lists/Req', status: 'changed', item: { list: { Title: 'Req', EnableVersioning: true } }, diffs: [{ prop: 'Versioning', source: 'Yes', target: 'No' }] };
  const c1 = { name: 'A', title: 'A', type: 'Text', status: 'missing', field: { SchemaXml: '<Field Name="A"/>' } };
  const c2 = { name: 'B', title: 'B', type: 'Text', status: 'missing', field: { SchemaXml: '<Field Name="B"/>' } };
  const c3 = { name: 'T', title: 'T', type: 'TaxonomyFieldType', status: 'missing', field: { SchemaXml: '<Field/>' } };
  const view = [{ list: changedList, columns: [c1, c2, c3], showSettings: true }, { list: missingList, columns: [], showSettings: false }];
  const ids = changeItems(view).map((i) => i.id);
  assert.deepEqual(ids, [settingsChangeId(changedList), columnChangeId(changedList, c1), columnChangeId(changedList, c2), listChangeId(missingList)]);
  const only = allStages(view, ctx, (id) => id === columnChangeId(changedList, c2));
  assert.deepEqual(only.map((s) => s.name), ['Columns']);
  assert.match(only[0].body, /Name=\\"B\\"/);
  assert.doesNotMatch(only[0].body, /Name=\\"A\\"/);
  assert.deepEqual(allStages(view, ctx, () => false), []);
});
