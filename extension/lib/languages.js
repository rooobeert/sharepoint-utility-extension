// Multi-language titles for "Compare": which languages a site has enabled, and the
// translated titles / descriptions of lists and columns (TitleResource, DescriptionResource).
// Pure functions: no DOM, no chrome.* APIs.

/** Language ID (LCID) -> culture name, for the languages SharePoint offers. */
export const CULTURES = {
  1025: 'ar-SA', 1026: 'bg-BG', 1027: 'ca-ES', 1028: 'zh-TW', 1029: 'cs-CZ', 1030: 'da-DK', 1031: 'de-DE', 1032: 'el-GR',
  1033: 'en-US', 1035: 'fi-FI', 1036: 'fr-FR', 1037: 'he-IL', 1038: 'hu-HU', 1040: 'it-IT', 1041: 'ja-JP', 1042: 'ko-KR',
  1043: 'nl-NL', 1044: 'nb-NO', 1045: 'pl-PL', 1046: 'pt-BR', 1048: 'ro-RO', 1049: 'ru-RU', 1050: 'hr-HR', 1051: 'sk-SK',
  1053: 'sv-SE', 1054: 'th-TH', 1055: 'tr-TR', 1058: 'uk-UA', 1060: 'sl-SI', 1061: 'et-EE', 1062: 'lv-LV', 1063: 'lt-LT',
  2052: 'zh-CN', 2057: 'en-GB', 2070: 'pt-PT', 3082: 'es-ES', 2074: 'sr-Latn-CS', 1087: 'kk-KZ',
};

/** What is compared when nothing else is typed. */
export const DEFAULT_LANGUAGES = 'de-DE, en-US';

/** "de-DE, en-US" -> { ids: [1031, 1033], unknown: [] }: language IDs of the culture names typed (any case, separated by comma, semicolon or space). */
export function parseLanguages(text) {
  const byName = new Map(Object.entries(CULTURES).map(([id, name]) => [name.toLowerCase(), Number(id)]));
  const ids = [];
  const unknown = [];
  for (const word of String(text || '').split(/[\s,;]+/).filter(Boolean)) {
    const id = byName.get(word.toLowerCase());
    if (id == null) unknown.push(word);
    else if (!ids.includes(id)) ids.push(id);
  }
  return { ids, unknown };
}

/** The languages to read: the chosen ones, except a language that is the default language of both sites (that is the title itself). */
export function compareLanguages(ids, ...webs) {
  return ids.filter((id) => !webs.every((w) => w && w.Language === id));
}

/** Culture names in the labels "Display name (de-DE)": [{ culture, id }]. */
const LABEL = /^(Title|Description|Display name) \(([A-Za-z]{2,3}(?:-[A-Za-z0-9]+)*)\)$/;

/** Label of a translated value: "Title (de-DE)" for lists, "Display name (de-DE)" for columns. */
export const translationLabel = (kind, what, id) => `${kind === 'column' ? 'Display name' : what} (${CULTURES[id]})`;

/** { what: 'Title' | 'Description', culture } of a translation label, or null for any other property. */
export function parseTranslationLabel(label) {
  const m = LABEL.exec(label || '');
  return m ? { what: m[1] === 'Description' ? 'Description' : 'Title', culture: m[2] } : null;
}

/**
 * Translated values as compared properties: label -> text. `tr` is { Title: { [lcid]: text }, Description: { ... } }
 * (a language that could not be read is missing). `skip` is a Set of language IDs that are left out.
 */
export function translationProps(kind, tr, skip = new Set()) {
  const out = {};
  for (const what of ['Title', 'Description']) {
    if (kind === 'column' && what === 'Description') continue;
    for (const [id, text] of Object.entries((tr && tr[what]) || {})) {
      if (text == null || skip.has(Number(id))) continue;
      out[translationLabel(kind, what, id)] = text;
    }
  }
  return out;
}

/** Keep only the translated properties both sides could read, so a language that failed on one side is not a difference. */
export function bothSides(a, b) {
  const isTr = (k) => LABEL.test(k);
  const dropA = Object.keys(a).filter((k) => isTr(k) && !(k in b));
  const dropB = Object.keys(b).filter((k) => isTr(k) && !(k in a));
  for (const k of dropA) delete a[k];
  for (const k of dropB) delete b[k];
}
