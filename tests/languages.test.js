import { test } from 'node:test';
import assert from 'node:assert/strict';
import { compareLanguages, parseLanguages, translationProps, bothSides, parseTranslationLabel } from '../extension/lib/languages.js';
import { compareSites } from '../extension/lib/diff.js';
import { settingsStage, columnStage, allStages, changeItems, onlyListId, onlyColumnId } from '../extension/lib/transfer.js';

const ctx = { web: { ServerRelativeUrl: '/sites/dev' }, allLists: [] };

test('parseLanguages and compareLanguages: only the chosen languages, not the default language of both sites', () => {
  assert.deepEqual(parseLanguages('de-DE, EN-us fr-XX'), { ids: [1031, 1033], unknown: ['fr-XX'] });
  assert.deepEqual(compareLanguages([1031, 1033], { Language: 1033 }, { Language: 1033 }), [1031]);
  assert.deepEqual(compareLanguages([1031, 1033], { Language: 1033 }, { Language: 1031 }), [1031, 1033]);
});

test('translationProps: labels, default language skipped, columns have no description', () => {
  const tr = { Title: { 1031: 'Anfragen', 1033: 'Requests' }, Description: { 1031: 'Text' } };
  assert.deepEqual(translationProps('list', tr, new Set([1033])), { 'Title (de-DE)': 'Anfragen', 'Description (de-DE)': 'Text' });
  assert.deepEqual(translationProps('column', tr), { 'Display name (de-DE)': 'Anfragen', 'Display name (en-US)': 'Requests' });
  assert.deepEqual(parseTranslationLabel('Display name (de-DE)'), { what: 'Title', culture: 'de-DE' });
  assert.equal(parseTranslationLabel('Required'), null);
  const a = { 'Title (de-DE)': 'x', Required: 'No' }; const b = { Required: 'No' };
  bothSides(a, b);
  assert.deepEqual(a, { Required: 'No' });
});

const side = (lang, title, fields, translations) => ({
  web: { ServerRelativeUrl: '/sites/dev', Language: lang }, allLists: [],
  lists: [{ entry: { title: 'Req', url: '/sites/dev/Lists/Req' }, errors: {}, translations,
    list: { Title: 'Req', BaseTemplate: 100, RootFolder: { ServerRelativeUrl: '/sites/dev/Lists/Req' } },
    fields: { value: fields } }],
});
const field = (name, title) => ({ InternalName: name, Title: title, Id: name, TypeAsString: 'Text', SchemaXml: '<Field/>', CanBeDeleted: true, FromBaseType: false, Group: 'Custom' });

test('compareSites finds differing translations, and the stage sets them in the target', () => {
  const f = field('Foo', 'Foo');
  const src = side(1033, 'Req', [f], { list: { Title: { 1031: 'Anfragen' }, Description: {} }, columns: { Foo: { Title: { 1031: 'Fu' } } } });
  const tgt = side(1033, 'Req', [f], { list: { Title: { 1031: 'Alt' }, Description: {} }, columns: { Foo: { Title: { 1031: 'Fu' } } } });
  const [list] = compareSites(src, tgt);
  assert.equal(list.status, 'changed');
  assert.deepEqual(list.diffs.map((d) => [d.prop, d.source, d.target]), [['Title (de-DE)', 'Anfragen', 'Alt']]);
  list.item = src.lists[0];
  const st = settingsStage(list);
  assert.match(st.body, /POST .*getList\(.*\)\/TitleResource\/SetValueForUICulture/);
  assert.match(st.body, /"cultureName":"de-DE","value":"Anfragen"/);
  assert.match(st.body, /^MERGE /m);
});

test('a column translation is sent with the column', () => {
  const col = { name: 'Foo', title: 'Foo', type: 'Text', status: 'changed', field: { TypeAsString: 'Text' }, diffs: [{ prop: 'Display name (de-DE)', source: 'Fu', target: 'Alt' }] };
  const st = columnStage({ title: 'Req', rel: 'Lists/Req' }, col, ctx);
  assert.match(st.body, /fields\/getByInternalNameOrTitle\('Foo'\)\/TitleResource\/SetValueForUICulture/);
});

const only = { title: 'Old', rel: 'Lists/Old', key: 'lists/old', status: 'only', diffs: [], columns: [] };
const gone = { name: 'Gone', title: 'Gone', custom: true, type: 'Text', status: 'only' };
const view = [{ list: only, columns: [], showSettings: false }, { list: { title: 'Req', rel: 'Lists/Req', key: 'lists/req', status: 'same', diffs: [] }, columns: [gone], showSettings: false }];

test('only in the target: nothing is sent unless renaming or deleting is chosen', () => {
  assert.deepEqual(changeItems(view), []);
  assert.deepEqual(allStages(view, ctx), []);
  assert.deepEqual(changeItems(view, { mode: 'rename', prefix: 'zz_' }).map((i) => i.id), [onlyListId(only), onlyColumnId(view[1].list, gone)]);
});

test('only in the target: rename adds the prefix once, delete recycles lists and deletes columns', () => {
  const rename = allStages(view, { ...ctx, only: { mode: 'rename', prefix: 'zz_' } });
  assert.deepEqual(rename.map((s) => s.name), ['Rename what only the target has']);
  assert.match(rename[0].body, /"Title":"zz_Old"/);
  assert.match(rename[0].body, /"Title":"zz_Gone"/);
  const done = [{ list: { ...only, title: 'zz_Old' }, columns: [], showSettings: false }];
  assert.deepEqual(allStages(done, { ...ctx, only: { mode: 'rename', prefix: 'zz_' } }), []);
  const del = allStages(view, { ...ctx, only: { mode: 'delete', prefix: 'zz_' } });
  assert.match(del[0].body, /POST \{\{WebUrl\}\}\/_api\/web\/getList\([^)]*\)\/recycle/);
  assert.match(del[0].body, /DELETE .*fields\/getByInternalNameOrTitle\('Gone'\)/);
  assert.deepEqual(allStages(view, { ...ctx, only: { mode: 'delete' } }, (id) => id === onlyListId(only)).length, 1);
});

import { createExporter } from '../extension/lib/exporter.js';

test('readTranslations asks with an Accept-Language header: two requests per list and language (one when the list fails)', async () => {
  const calls = [];
  const request = async (req) => {
    calls.push(req);
    const lang = req.headers['Accept-Language'];
    if (lang === 'fr-FR') return { status: 500, headers: {}, text: '{"error":{}}' };
    const de = lang === 'de-DE';
    const json = /\/fields\?/.test(req.url)
      ? { value: [{ InternalName: 'c1', Title: de ? 'Spalte' : 'Column' }, { InternalName: 'sys', Title: 'System' }] }
      : { Title: de ? 'Anfragen' : 'Requests', Description: de ? 'Beschreibung' : 'Description' };
    return { status: 200, headers: {}, text: JSON.stringify(json) };
  };
  const fields = { value: [{ InternalName: 'c1', Id: 'a', Title: 'c1', FromBaseType: false }, { InternalName: 'sys', Id: 'b', Title: 'sys', FromBaseType: true }] };
  const data = { lists: [{ entry: { id: 'L1' }, list: { Title: 'L' }, fields }] };
  await createExporter({ request, sleep: async () => {} }).readTranslations('https://t.sharepoint.com/sites/x', data, [1031, 1033, 1036]);
  assert.equal(calls.length, 5); // fr-FR fails on the list, so its columns are not asked
  assert.ok(calls.every((c) => c.method === 'GET' && c.headers['Accept-Language']));
  const tr = data.lists[0].translations;
  assert.equal(tr.list.Title[1031], 'Anfragen');
  assert.equal(tr.list.Description[1033], 'Description');
  assert.equal(tr.list.Title[1036], undefined);
  assert.equal(tr.columns.c1.Title[1031], 'Spalte');
  assert.equal(tr.columns.sys, undefined);
});
