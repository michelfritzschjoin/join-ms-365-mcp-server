/**
 * UQAS Pro - Internationalization Module (DE/EN)
 *
 * Provides bilingual support for the Universal Question Answering System:
 * - Language detection (German/English)
 * - Cross-language thesaurus for search optimization
 * - Bilingual intent and entity recognition
 * - Language-aware response templates
 */

export { LanguageDetector, type LanguageDetectionResult } from './language-detector.js';
export {
  BilingualThesaurus,
  BILINGUAL_THESAURUS,
  type ThesaurusEntry,
} from './bilingual-thesaurus.js';
export {
  BilingualIntentRecognizer,
  INTENT_PATTERNS,
  type BilingualIntent,
  type IntentType,
} from './intent-patterns.js';
export {
  BilingualEntityRecognizer,
  PERSON_PATTERNS,
  TEMPORAL_PATTERNS,
  type ExtractedEntity,
  type TemporalExpression,
} from './entity-patterns.js';
export {
  BilingualResponseBuilder,
  RESPONSE_TEMPLATES,
  type ResponseData,
  type ResponseTemplates,
} from './response-templates.js';

export type SupportedLanguage = 'de' | 'en';
