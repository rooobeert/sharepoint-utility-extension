import { test } from 'node:test';
import assert from 'node:assert/strict';
import {
  REFINABLE, splitCrawled, SP_PROPSET, parseQuoted, quoteList, crawledUrl, columnsFromLists, typeWarning,
  crawledStatus, patchStatus, mergeStatusLists, freeCount, parsePaste, crawledSuggestions, managedSuggestions, refinableType, rowState, pendingChanges, nextMappings, verifyProblems,
} from '../extension/lib/schema.js';

test('Refinable suggestions cover all site collection Refinable properties', () => {
  assert.equal(REFINABLE.length, 297);
  assert.equal(REFINABLE[0], 'RefinableString00');
  assert.equal(REFINABLE[199], 'RefinableString199');
  assert.ok(REFINABLE.includes('RefinableDateInvariant01'));
});

test('crawled names: default property set, explicit set, URL', () => {
  assert.deepEqual(splitCrawled(' ows_Kunde '), { name: 'ows_Kunde', propset: SP_PROPSET });
  assert.deepEqual(splitCrawled('ows_X@abc'), { name: 'ows_X', propset: 'abc' });
  assert.equal(crawledUrl('https://t.sharepoint.com/sites/a/_layouts/15', 'ows_Kunde'),
    `https://t.sharepoint.com/sites/a/_layouts/15/crawledproperty.aspx?property=ows_Kunde&propertyset=${SP_PROPSET}&level=sitecol`);
});

test('hidden mapping field round-trips', () => {
  const s = quoteList(['RefinableString17', 'RefinableInt18']);
  assert.equal(s, '"RefinableString17" "RefinableInt18" ');
  assert.deepEqual(parseQuoted(s), ['RefinableString17', 'RefinableInt18']);
  assert.equal(quoteList([]), '');
});

test('list columns become crawled names; taxonomy gets two', () => {
  const cols = columnsFromLists({ value: [
    { Title: 'Angebote', Fields: [
      { InternalName: 'Kunde', Title: 'Kunde', TypeAsString: 'Text' },
      { InternalName: 'Produkt', Title: 'Produkt', TypeAsString: 'TaxonomyFieldType' },
      { InternalName: 'Title', Title: 'Title', TypeAsString: 'Text', FromBaseType: true },
      { InternalName: 'Hid', Title: 'Hid', TypeAsString: 'Text', Hidden: true },
    ] },
    { Title: 'Aufträge', Fields: [{ InternalName: 'Kunde', Title: 'Kunde', TypeAsString: 'Text' }] },
  ] });
  assert.deepEqual([...cols.keys()].sort(), ['ows_kunde', 'ows_produkt', 'ows_taxid_produkt']);
  assert.deepEqual([...cols.get('ows_kunde').lists], ['Angebote', 'Aufträge']);
});

test('type warnings only for numeric/date properties fed by other column types', () => {
  assert.equal(typeWarning('Text', 'RefinableString01', null), null);
  assert.equal(typeWarning('Number', 'RefinableInt01', null), null);
  assert.equal(typeWarning('DateTime', 'RefinableDate01', null), null);
  assert.equal(typeWarning('Text', 'RefinableInt01', null), 'Text column into a numeric property');
  assert.equal(typeWarning('Choice', 'RefinableDate01', 'DateTime'), 'Choice column into a date property');
  assert.equal(typeWarning(null, 'RefinableInt01', null), null);
});

test('status: crawled view, patch after save, merge, free count', () => {
  const items = mergeStatusLists([
    [{ name: 'RefinableString00', type: 'Text', crawled: ['OWS_KUNDE'] }, { name: 'Title', type: 'Text', crawled: [] }],
    [{ name: 'RefinableString01', type: 'Text', crawled: [] }, { name: 'refinablestring00', type: 'Text', crawled: ['OWS_KUNDE'] }],
  ]);
  assert.equal(items.length, 2);
  const status = new Map(items.map((i) => [i.name.toLowerCase(), i]));
  assert.deepEqual(crawledStatus(status.values()).get('ows_kunde').managed, ['refinablestring00']);
  assert.equal(freeCount(status), 1);
  patchStatus(status, 'ows_Land', ['RefinableString01'], []);
  patchStatus(status, 'ows_Kunde', [], ['RefinableString00']);
  assert.deepEqual(status.get('refinablestring01').crawled, ['ows_Land']);
  assert.deepEqual(status.get('refinablestring00').crawled, []);
});

test('paste: one row per managed property, single values pass through', () => {
  assert.equal(parsePaste('ows_Kunde'), null);
  assert.deepEqual(parsePaste('ows_Kunde\tRefinableString00\r\nows_Land;RefinableString01;RefinableString02\nows_X\n\n'), [
    ['ows_Kunde', 'RefinableString00'], ['ows_Land', 'RefinableString01'], ['ows_Land', 'RefinableString02'], ['ows_X', ''],
  ]);
  assert.equal(parsePaste('ows_A => RefinableInt00\n'), null); // one line without a tab: normal paste
  assert.deepEqual(parsePaste('ows_A => RefinableInt00\nows_B -> RefinableInt01'), [['ows_A', 'RefinableInt00'], ['ows_B', 'RefinableInt01']]);
});

const crawled = (current = []) => ({ exists: true, current });
const managed = (mapped = [], type = 'Text') => ({ exists: true, type, mapped });

test('row state', () => {
  assert.equal(rowState({ c: '', m: '' }).state, 'empty');
  assert.equal(rowState({ c: 'ows_A', m: '', cr: { loading: true } }).state, 'checking');
  assert.deepEqual(rowState({ c: 'ows_A', m: 'X', cr: { exists: false }, mr: { exists: false } }),
    { state: 'error', detail: 'crawled property not found, managed property not found', warnings: [] });
  assert.equal(rowState({ c: 'ows_A', m: '', cr: crawled() }).state, 'incomplete');
  assert.equal(rowState({ c: '', m: 'RefinableString00', mr: managed() }).state, 'incomplete');
  assert.equal(rowState({ c: 'ows_A', m: 'RefinableString00', cr: crawled(), mr: managed() }).state, 'ready');
  assert.equal(rowState({ c: 'ows_A', m: 'RefinableString00', cr: crawled(['refinablestring00']), mr: managed(['ows_A']) }).state, 'already');
  // Removing an existing mapping makes the row savable, even without a managed property.
  assert.deepEqual(rowState({ c: 'ows_A', m: '', cr: crawled(['RefinableString03']), removing: ['RefinableString03'] }),
    { state: 'ready', detail: 'removes RefinableString03', warnings: [] });
});

test('row state warnings: shared managed property and type mismatch', () => {
  const shared = rowState({ c: 'ows_Land', m: 'RefinableString12', cr: crawled(), mr: managed(['OWS_REGION']), display: () => 'ows_Region' });
  assert.equal(shared.state, 'warn');
  assert.equal(shared.detail, 'RefinableString12 is also fed by ows_Region');
  const typed = rowState({ c: 'ows_A', m: 'RefinableInt00', cr: crawled(), mr: managed([], 'Integer'), columnType: 'Text' });
  assert.equal(typed.state, 'warn');
  assert.equal(typed.detail, 'Text column into a numeric property');
  // The crawled property itself doesn't count as another feed (with or without property set).
  assert.equal(rowState({ c: 'ows_A@x', m: 'RefinableString00', cr: crawled(), mr: managed(['OWS_A']) }).state, 'ready');
});

test('pending changes are grouped per crawled property', () => {
  const rows = [
    { state: 'ready', c: 'ows_A', m: 'RefinableString00', cr: crawled(['RefinableString03']) },
    { state: 'warn', c: 'OWS_A', m: 'RefinableString01', cr: crawled(['RefinableString03']) },
    { state: 'ready', c: 'ows_A', m: 'RefinableString00', cr: crawled(['RefinableString03']) },
    { state: 'already', c: 'ows_B', m: 'RefinableString05', cr: crawled(['RefinableString05']) },
    { state: 'error', c: 'ows_C', m: 'RefinableString06', cr: { exists: false } },
  ];
  const removals = new Map([['ows_a', new Map([['refinablestring03', 'RefinableString03']])],
    ['ows_b', new Map([['refinablestring99', 'RefinableString99']])]]);
  assert.deepEqual(pendingChanges(rows, removals), [
    { crawled: 'ows_A', add: ['RefinableString00', 'RefinableString01'], remove: ['RefinableString03'] },
  ]);
  assert.deepEqual(pendingChanges(rows.slice(3), new Map()), []);
});

test('next mappings and verification after save', () => {
  assert.deepEqual(nextMappings(['RefinableString03', 'RefinableString40'], ['RefinableString00', 'refinablestring40'], ['refinablestring03']),
    { next: ['RefinableString40', 'RefinableString00'], deleted: ['RefinableString03'] });
  assert.deepEqual(verifyProblems(['RefinableString00'], ['RefinableString00'], []), []);
  assert.deepEqual(verifyProblems(['RefinableString03'], ['RefinableString00'], ['RefinableString03']),
    ['RefinableString00 missing', 'RefinableString03 still mapped']);
});

test('crawled suggestions: columns and schema-only names, prefix matches first, mapping from the status', () => {
  const columns = columnsFromLists({ value: [{ Title: 'Angebote', Fields: [
    { InternalName: 'Land', Title: 'Land', TypeAsString: 'Choice' },
    { InternalName: 'Kunde', Title: 'Kundenname', TypeAsString: 'Text' },
    { InternalName: 'Region', Title: 'Region', TypeAsString: 'Choice' },
  ] }] });
  const empty = crawledSuggestions(columns, new Map(), '');
  assert.deepEqual(empty.items.map((x) => x.name), ['ows_Kunde', 'ows_Land', 'ows_Region']);
  assert.deepEqual(empty.items[0], { name: 'ows_Kunde', detail: 'Kundenname · Text · Angebote', mapped: [] });

  const status = new Map([
    ['refinablestring12', { name: 'RefinableString12', type: 'Text', crawled: ['OWS_REGION'] }],
    ['refinablestring13', { name: 'RefinableString13', type: 'Text', crawled: ['OWS_OTHER'] }],
  ]);
  const all = crawledSuggestions(columns, status, '');
  assert.deepEqual(all.items.find((x) => x.name === 'ows_Region').mapped, ['RefinableString12']);
  assert.deepEqual(all.items.find((x) => x.name === 'OWS_OTHER'), { name: 'OWS_OTHER', detail: 'not a list column of this site', mapped: ['RefinableString13'] });

  // Name and column title both match.
  assert.deepEqual(crawledSuggestions(columns, status, 'ows_r').items.map((x) => x.name), ['ows_Region']);
  assert.deepEqual(crawledSuggestions(columns, status, 'kundenname').items.map((x) => x.name), ['ows_Kunde']);
  const limited = crawledSuggestions(columns, status, 'ows', 2);
  assert.equal(limited.items.length, 2);
  assert.equal(limited.more, 2);
});

test('managed suggestions: free, used, missing, and nothing before the status is loaded', () => {
  assert.equal(refinableType('RefinableDateSingle01'), 'DateTime');
  assert.equal(refinableType('RefinableInt07'), 'Integer');
  assert.equal(refinableType('Title'), null);

  const before = managedSuggestions(new Map(), false, 'RefinableString1');
  assert.equal(before.items[0].name, 'RefinableString10');
  assert.equal(before.items[0].state, null);
  assert.equal(before.items[0].detail, 'Text');
  assert.equal(before.items.length + before.more, 110); // 10-19 and 100-199

  const status = new Map([
    ['refinablestring10', { name: 'RefinableString10', type: 'Text', crawled: [] }],
    ['refinablestring12', { name: 'RefinableString12', type: 'Text', crawled: ['OWS_REGION'] }],
  ]);
  const after = managedSuggestions(status, true, 'refinablestring1', 4);
  assert.deepEqual(after.items.map((x) => [x.name, x.state]), [
    ['RefinableString10', 'free'], ['RefinableString11', 'missing'], ['RefinableString12', 'used'], ['RefinableString13', 'missing'],
  ]);
  assert.deepEqual(after.items[2].crawled, ['OWS_REGION']);
  assert.equal(managedSuggestions(status, true, '').items[0].name, 'RefinableString00');
});
