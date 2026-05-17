import type { i18n as I18nInstance } from 'i18next';

/**
 * Maps exact workbook / Excel strings (typically Swedish in the mall file) to i18n keys
 * under the `entities` or `app` namespace. Expand as you add locale entries.
 *
 * Prefer stable entity ids (e.g. ISO 3166-1 alpha-2) under entities.countries.*.
 */
const WORKBOOK_STRING_TO_I18N_KEY: Record<string, { ns: 'app' | 'entities'; key: string }> = {
  Bosnien: { ns: 'entities', key: 'countries.BA' },
  Mexiko: { ns: 'entities', key: 'countries.MX' },
  Ja: { ns: 'app', key: 'choices.yes' },
  Nej: { ns: 'app', key: 'choices.no' },
};

export function translateWorkbookString(raw: string, i18n: I18nInstance): string {
  const rule = WORKBOOK_STRING_TO_I18N_KEY[raw];
  if (!rule) {
    return raw;
  }
  return i18n.t(rule.key, { ns: rule.ns });
}
