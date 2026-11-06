// Search schema mapping (site collection level), ported from the sp-schema-mapper bookmarklet.
// Unsupported by Microsoft: there is no API for this, so the classic search schema pages are read
// and the postback that OK on crawledproperty.aspx sends is replayed.
//
// Everything here is pure except the page parsers (parse*Page, formBody), which take a parsed
// Document and are used by the side panel only.

export const SP_PROPSET = '00130329-0000-0130-c000-000000131346';
export const STATUS_MAX_AGE = 60 * 60 * 1000;
export const STATUS_FILTERS = ['RefinableString', 'RefinableInt', 'RefinableDate', 'RefinableDecimal', 'RefinableDouble'];

export const lc = (s) => s.toLowerCase();
export const sameName = (a, b) => lc(a) === lc(b);
export const has = (list, x) => list.some((y) => sameName(x, y));
export const parseQuoted = (s) => [...s.matchAll(/"([^"]+)"/g)].map((m) => m[1]);
/** The mappings hidden field's format: "RefinableString17" "RefinableInt18" */
export const quoteList = (list) => list.map((x) => `"${x}" `).join('');

export const REFINABLE = [['RefinableString', 200], ['RefinableInt', 50], ['RefinableDate', 20], ['RefinableDateSingle', 5],
  ['RefinableDateInvariant', 2], ['RefinableDecimal', 10], ['RefinableDouble', 10]]
  .flatMap(([prefix, n]) => Array.from({ length: n }, (_, i) => prefix + String(i).padStart(2, '0')));
export const REFINABLE_SET = new Set(REFINABLE.map(lc));

/** "ows_Kunde" or "ows_Kunde@<propset>" -> { name, propset } */
export function splitCrawled(value) {
  const [name, propset = SP_PROPSET] = value.trim().split('@');
  return { name, propset };
}

export const crawledUrl = (layouts, value) => {
  const { name, propset } = splitCrawled(value);
  return `${layouts}/crawledproperty.aspx?property=${encodeURIComponent(name)}` +
    `&propertyset=${encodeURIComponent(propset)}&level=sitecol`;
};
export const managedUrl = (layouts, name) => `${layouts}/managedproperty.aspx?property=${encodeURIComponent(name)}&level=sitecol`;
export const managedListUrl = (layouts) => `${layouts}/listmanagedproperties.aspx?level=sitecol`;
export const columnsUrl = (site) => `${site}/_api/web/lists?$filter=Hidden eq false&$select=Title,Fields/InternalName,Fields/Title,` +
  'Fields/TypeAsString,Fields/FromBaseType,Fields/Hidden,Fields/Sealed&$expand=Fields';

// ---------------------------------------------------------------- list columns

/** Crawled property names are predictable from list columns: ows_<InternalName>, plus ows_taxId_<InternalName> for managed metadata. */
export function crawledNamesForField(f) {
  return /^Taxonomy/.test(f.TypeAsString) ? ['ows_taxId_' + f.InternalName, 'ows_' + f.InternalName] : ['ows_' + f.InternalName];
}

/** /_api/web/lists?$expand=Fields JSON -> Map(lc(crawled) -> { name, title, type, lists: Set }) */
export function columnsFromLists(json) {
  const columns = new Map();
  for (const list of json.value || []) {
    for (const f of list.Fields || []) {
      if (f.Hidden || f.FromBaseType || f.Sealed) continue;
      for (const name of crawledNamesForField(f)) {
        const e = columns.get(lc(name)) || { name, title: f.Title, type: f.TypeAsString, lists: new Set() };
        e.lists.add(list.Title);
        columns.set(lc(name), e);
      }
    }
  }
  return columns;
}

// ---------------------------------------------------------------- types

// Only flag combinations that can't work: a numeric/date managed property fed from an incompatible column.
export const fieldKind = (type) =>
  /^(Integer|Counter|Number|Currency)$/.test(type) ? 'num' : type === 'DateTime' ? 'date' : 'text';

export function managedKind(type, name) {
  const t = type || (/^RefinableInt/i.test(name) ? 'Integer' : /^RefinableDate/i.test(name) ? 'DateTime' :
    /^Refinable(Double|Decimal)/i.test(name) ? 'Double' : /^RefinableString/i.test(name) ? 'Text' : null);
  return /^(Integer|Decimal|Double)$/.test(t) ? 'num' : t === 'DateTime' ? 'date' : t ? 'text' : null;
}

/** Warning text, or null. `columnType` is the list column's TypeAsString (unknown -> null). */
export function typeWarning(columnType, managed, managedType) {
  const mk = managedKind(managedType, managed);
  if (!columnType || !mk || mk === 'text') return null;
  return fieldKind(columnType) === mk ? null : `${columnType} column into ${mk === 'date' ? 'a date' : 'a numeric'} property`;
}

// ---------------------------------------------------------------- mapping status

/** Managed status items [{ name, type, crawled: [] }] -> Map(lc(crawled) -> { name, managed: [] }) */
export function crawledStatus(items) {
  const map = new Map();
  for (const m of items) {
    for (const c of m.crawled) {
      const e = map.get(lc(c)) || { name: c, managed: [] };
      e.managed.push(m.name);
      map.set(lc(c), e);
    }
  }
  return map;
}

/** Update the status Map(lc(managed) -> item) after a save, without reloading the lists. */
export function patchStatus(status, crawled, added, removed) {
  for (const m of added) {
    const e = status.get(lc(m)) || { name: m, type: null, crawled: [] };
    if (!has(e.crawled, crawled)) e.crawled.push(crawled);
    status.set(lc(m), e);
  }
  for (const m of removed) {
    const e = status.get(lc(m));
    if (e) e.crawled = e.crawled.filter((c) => !sameName(c, crawled));
  }
}

/** Keep only Refinable* items, one per name. */
export function mergeStatusLists(lists) {
  const items = new Map();
  for (const i of lists.flat()) if (/^Refinable/i.test(i.name)) items.set(lc(i.name), i);
  return [...items.values()];
}

export const freeCount = (status) => REFINABLE.filter((n) => status.get(lc(n))?.crawled.length === 0).length;

// ---------------------------------------------------------------- suggestions

const REFINABLE_TYPES = { string: 'Text', int: 'Integer', date: 'DateTime', decimal: 'Decimal', double: 'Double' };
/** Type of a Refinable* property from its name (Text, Integer, DateTime, Decimal, Double), or null. */
export const refinableType = (name) => REFINABLE_TYPES[lc(name.match(/^Refinable(String|Int|Date|Decimal|Double)/i)?.[1] || '')] || null;

/** Matches `q` (lower case) first by name prefix, then by containing it; then by name. */
function rank(list, q, keys) {
  const score = (x) => (lc(x.name).startsWith(q) ? 0 : 1);
  return list
    .filter((x) => !q || keys(x).some((k) => k && lc(k).includes(q)))
    .sort((a, b) => score(a) - score(b) || a.name.localeCompare(b.name, undefined, { numeric: true, sensitivity: 'base' }));
}

/**
 * Crawled property suggestions for the dropdown: list columns, plus crawled properties the status shows as mapped.
 *   columns: Map from columnsFromLists, status: Map(lc(managed) -> item).
 * Returns { items: [{ name, detail, mapped: [] }], more } (mapped is [] until the status is loaded, then the managed names).
 */
export function crawledSuggestions(columns, status, query, limit = 50) {
  const cs = crawledStatus(status.values());
  const all = new Map([...columns].map(([k, c]) => [k, { name: c.name, detail: `${c.title} · ${c.type} · ${[...c.lists].join(', ')}`, title: c.title }]));
  for (const [k, c] of cs) if (!all.has(k)) all.set(k, { name: c.name, detail: 'not a list column of this site' });
  const found = rank([...all.entries()].map(([k, x]) => ({ ...x, mapped: cs.get(k)?.managed || [] })), lc(query.trim()), (x) => [x.name, x.title]);
  return { items: found.slice(0, limit).map(({ title, ...x }) => x), more: Math.max(0, found.length - limit) };
}

/**
 * Managed property suggestions: every Refinable* property with its type and who feeds it.
 *   state: 'free' | 'used' | 'missing' (not in the loaded status) | null (status not loaded yet)
 * Returns { items: [{ name, detail, state, crawled: [] }], more }.
 */
export function managedSuggestions(status, loaded, query, limit = 50) {
  const all = REFINABLE.map((name) => {
    const s = status.get(lc(name));
    const state = !loaded ? null : !s ? 'missing' : s.crawled.length ? 'used' : 'free';
    return { name, detail: s?.type || refinableType(name), state, crawled: s?.crawled || [] };
  });
  const q = lc(query.trim());
  const found = q ? rank(all, q, (x) => [x.name]) : all;
  return { items: found.slice(0, limit), more: Math.max(0, found.length - limit) };
}

// ---------------------------------------------------------------- paste

/**
 * Excel / text paste: one row per line, "crawled<TAB>managed[<TAB>managed...]" (also ; , => -> as separators).
 * Returns [[crawled, managed], ...], or null when the text is a single value (let the browser paste it).
 */
export function parsePaste(text) {
  if (!/[\n\t]/.test(text.trim())) return null;
  const out = [];
  for (const line of text.split(/\r?\n/)) {
    const [crawled, ...rest] = line.split(/\t|;|,|=>|->/).map((s) => s.trim());
    if (!crawled) continue;
    const managed = rest.filter(Boolean);
    for (const m of managed.length ? managed : ['']) out.push([crawled, m]);
  }
  return out;
}

// ---------------------------------------------------------------- row state

/**
 * State of one row from its lookups.
 *   c, m: the values checked; cr: crawled lookup ({ loading } | { error } | { exists, current }),
 *   mr: managed lookup ({ loading } | { error } | { exists, type, mapped }), removing: managed names
 *   to unmap from c, columnType: the list column type of c (or null), display: shows a crawled name
 *   (the managed-property list upper-cases them).
 * Returns { state, detail, warnings } with state in empty | incomplete | checking | error | ready | warn | already.
 */
export function rowState({ c, m, cr, mr, removing = [], columnType = null, display = (x) => x }) {
  const errors = [];
  const warnings = [];
  if (cr?.error) errors.push('crawled property check failed');
  else if (cr && !cr.loading && !cr.exists) errors.push('crawled property not found');
  if (mr?.error) errors.push('managed property check failed');
  else if (mr && !mr.loading && !mr.exists) errors.push('managed property not found');
  if (mr?.exists) {
    const others = otherFeeds(mr, c);
    if (others.length) warnings.push(`${m} is also fed by ${others.map(display).join(', ')}`);
  }

  const removeNote = removing.length ? `removes ${removing.join(', ')}` : '';
  const out = (state, detail = '') => ({ state, detail, warnings });
  if (!c && !m) return out('empty');
  if (cr?.loading || mr?.loading) return out('checking');
  if (errors.length) return out('error', errors.join(', '));
  if (!m) return out(removing.length ? 'ready' : 'incomplete', removeNote || 'pick a managed property');
  if (!c) return out('incomplete', 'pick a crawled property');
  if (has(cr.current, m) && !has(removing, m)) return out(removing.length ? 'ready' : 'already', removeNote);
  if (has(removing, m)) return out('ready', removeNote);
  const tw = typeWarning(columnType, m, mr.type);
  if (tw) warnings.push(tw);
  return out(warnings.length ? 'warn' : 'ready', [removeNote, ...warnings].filter(Boolean).join(' · '));
}

/** Crawled properties feeding the managed property, other than c. */
export const otherFeeds = (mr, c) => mr.mapped.filter((x) => !c || !sameName(x, splitCrawled(c).name));

/**
 * Changes to write, grouped per crawled property.
 *   rows: [{ state, c, m, cr }], removals: Map(lc(crawled) -> Map(lc(managed) -> managed))
 * Returns [{ crawled, add: [], remove: [] }] with at least one change each.
 */
export function pendingChanges(rows, removals) {
  const groups = new Map();
  const group = (c) => {
    const g = groups.get(lc(c)) || { crawled: c, add: [], remove: [] };
    groups.set(lc(c), g);
    return g;
  };
  const removalsFor = (c) => [...(removals.get(lc(c))?.values() || [])];
  for (const r of rows) {
    if (r.state !== 'ready' && r.state !== 'warn') continue;
    const g = group(r.c);
    if (r.m && !has(r.cr.current, r.m) && !has(removalsFor(r.c), r.m) && !has(g.add, r.m)) g.add.push(r.m);
  }
  for (const [k, set] of removals) {
    const r = rows.find((x) => x.c && lc(x.c) === k && x.cr?.exists);
    if (!r) continue;
    const g = group(r.c);
    for (const m of set.values()) if (has(r.cr.current, m) && !has(g.remove, m)) g.remove.push(m);
  }
  return [...groups.values()].filter((g) => g.add.length || g.remove.length);
}

/** Mapping list to post for one crawled property: { next, deleted } */
export function nextMappings(current, add, remove) {
  const deleted = current.filter((x) => has(remove, x));
  const next = current.filter((x) => !has(remove, x));
  for (const m of add) if (!has(next, m)) next.push(m);
  return { next, deleted };
}

/** After saving: what the re-read page still gets wrong (empty = confirmed). */
export function verifyProblems(after, add, remove) {
  return [...add.filter((m) => !has(after, m)).map((m) => `${m} missing`),
    ...remove.filter((m) => has(after, m)).map((m) => `${m} still mapped`)];
}

// ---------------------------------------------------------------- classic pages (need a parsed Document)

const field = (root, suffix) => root.querySelector(`[name$="$${suffix}"]`);
export const formAction = (form, pageUrl) => new URL(form.getAttribute('action') || pageUrl, pageUrl).href;

/** Serialize a WebForms form the way the browser would, then apply overrides (e.g. which button was clicked). */
export function formBody(form, overrides) {
  const body = new URLSearchParams();
  for (const node of form.elements) {
    if (!node.name || node.disabled) continue;
    const type = (node.type || '').toLowerCase();
    if (['submit', 'button', 'image', 'reset', 'file'].includes(type)) continue;
    if ((type === 'checkbox' || type === 'radio') && !node.checked) continue;
    if (node.tagName === 'SELECT') {
      for (const o of node.selectedOptions) body.append(node.name, o.value);
      continue;
    }
    body.append(node.name, node.value);
  }
  for (const [k, v] of Object.entries(overrides)) body.set(k, v);
  return body;
}

/** crawledproperty.aspx -> { exists, current, page } (page is what postMappings needs). */
export function parseCrawledPage(doc, url) {
  const hidden = field(doc, 'mappingsHiddenTextBox');
  if (!hidden) return { exists: false };
  const form = hidden.form || doc.forms[0];
  let current = parseQuoted(hidden.value);
  if (!current.length) {
    const list = field(form, 'mappedPropertiesListBox');
    current = list ? [...list.options].map((o) => o.value).filter(Boolean) : [];
  }
  return { exists: true, current, page: { url, form, hidden } };
}

/** The OK postback for crawledproperty.aspx: { url, body } */
export function mappingsPostback(page, next, deleted) {
  const okButton = field(page.form, 'okButton');
  if (!okButton) throw new Error('OK button not found on crawledproperty.aspx');
  const hiddenName = page.hidden.name;
  const body = formBody(page.form, {
    [hiddenName]: quoteList(next),
    [hiddenName.replace(/mappingsHiddenTextBox$/, 'deletedMappingsHiddenTextBox')]: quoteList(deleted),
    [okButton.name]: okButton.value || 'OK',
  });
  return { url: formAction(page.form, page.url), body };
}

/** Validation message on a rejected postback page, or ''. */
export const pageError = (doc) =>
  doc.querySelector('.ms-error, .ms-formvalidation, [id$="errorLabel"], [id$="ErrorMessage"]')?.textContent.trim() || '';

const TYPE_BY_RADIO = { textRadioButton: 'Text', integerRadioButton: 'Integer', decimalRadioButton: 'Decimal',
  dateTimeRadioButton: 'DateTime', yesNoRadioButton: 'YesNo', doubleRadioButton: 'Double', binaryRadioButton: 'Binary' };

/** managedproperty.aspx -> { exists, name, type, mapped }; name is the property's own spelling (case), which saving needs. */
export function parseManagedPage(doc) {
  if (!field(doc, 'nameTextBox')) return { exists: false };
  const radio = doc.querySelector('input[type=radio][name$="$typeButtonGroup"][checked]');
  const list = field(doc, 'mappedPropertiesListBox');
  // Options: value "<propset>;ows_Name", text "ows_Name"
  const mapped = list ? [...list.options].map((o) => o.textContent.trim()).filter(Boolean) : [];
  return { exists: true, name: field(doc, 'nameTextBox').value.trim() || null, type: TYPE_BY_RADIO[radio?.value] || null, mapped };
}

/** listmanagedproperties.aspx: the filter postback ({ form, overrides }), or throws. */
export function managedListFilter(doc, filter) {
  const box = field(doc, 'managedPropertyTextBox');
  const apply = field(doc, 'btnDefaults');
  if (!box || !apply) throw new Error('filter box not found on listmanagedproperties.aspx');
  return { [box.name]: filter, [apply.name + '.x']: '0', [apply.name + '.y']: '0' };
}

/** listmanagedproperties.aspx: overrides for the next page, or null on the last page. */
export function managedListNext(doc) {
  const next = field(doc, 'nextButton');
  if (!next || next.disabled) return null;
  return { [next.name + '.x']: '4', [next.name + '.y']: '5' };
}

/** listmanagedproperties.aspx table -> [{ name, type, crawled: [] }] (crawled names come upper-cased). */
export function parseManagedList(doc, baseUrl) {
  const table = doc.querySelector('table[id$="_ListManagedProperties"]');
  if (!table) return [];
  const headers = [...table.rows[0].cells].map((c) => c.textContent.trim());
  const typeIdx = headers.indexOf('Type');
  const mappedIdx = headers.indexOf('Mapped Crawled Properties');
  const prop = (a) => new URL(a.getAttribute('href'), baseUrl).searchParams.get('property');
  const out = [];
  for (const tr of [...table.rows].slice(1)) {
    const link = tr.cells[0]?.querySelector('a.ms-core-menu-root');
    if (!link) continue;
    const name = prop(link) || link.firstChild.textContent.trim();
    const crawled = [...(tr.cells[mappedIdx]?.querySelectorAll('a') || [])].map((a) => prop(a) || a.textContent.trim());
    out.push({ name, type: tr.cells[typeIdx]?.textContent.trim() || null, crawled });
  }
  return out;
}
