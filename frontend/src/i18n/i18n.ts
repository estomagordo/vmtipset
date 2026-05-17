import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';

import appEn from './locales/en/app.json';
import appSv from './locales/sv/app.json';
import workbookLabelsEn from './locales/en/workbookLabels.json';
import workbookLabelsSv from './locales/sv/workbookLabels.json';

/**
 * i18n setup for the SPA.
 *
 * - Default locale: Swedish (`sv`), fallback: English (`en`).
 * - `app`: chrome (errors, hints, spreadsheet blurb).
 * - `workbookLabels`: all strings extracted from the Excel-backed dump (+ sheet title); regenerate via
 *   `python tools/generate_workbook_labels.py`.
 *
 * To switch language: `import i18n from './i18n/i18n'; i18n.changeLanguage('en')`.
 */
void i18n.use(initReactI18next).init({
  lng: 'sv',
  fallbackLng: 'en',
  defaultNS: 'app',
  ns: ['app', 'workbookLabels'],
  resources: {
    sv: { app: appSv, workbookLabels: workbookLabelsSv },
    en: { app: appEn, workbookLabels: workbookLabelsEn },
  },
  interpolation: {
    escapeValue: true,
  },
});

function syncDocumentLanguage(language: string): void {
  document.documentElement.lang = language;
}

syncDocumentLanguage(i18n.language);
i18n.on('languageChanged', syncDocumentLanguage);

export default i18n;
