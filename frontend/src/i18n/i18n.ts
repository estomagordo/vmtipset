import i18n from 'i18next';
import { initReactI18next } from 'react-i18next';

import appEn from './locales/en/app.json';
import appSv from './locales/sv/app.json';
import entitiesEn from './locales/en/entities.json';
import entitiesSv from './locales/sv/entities.json';

/**
 * i18n setup for the SPA.
 *
 * - Default locale: Swedish (`sv`), fallback: English (`en`).
 * - `app`: chrome strings (errors, hints, spreadsheet blurb).
 * - `entities`: domain catalog (countries, etc.). Use stable ids (e.g. ISO codes) as JSON keys.
 *
 * To switch language in code (e.g. future settings UI): `import i18n from './i18n/i18n'; i18n.changeLanguage('en')`.
 * `document.documentElement.lang` is updated on init and when the language changes.
 */
void i18n.use(initReactI18next).init({
  lng: 'sv',
  fallbackLng: 'en',
  defaultNS: 'app',
  ns: ['app', 'entities'],
  resources: {
    sv: { app: appSv, entities: entitiesSv },
    en: { app: appEn, entities: entitiesEn },
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
