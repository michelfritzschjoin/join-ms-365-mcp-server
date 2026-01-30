/**
 * User Profile Service - Manages user profile data and profession-based personalization
 *
 * Loads user profile from Microsoft Graph API (/me) and maps job titles/departments
 * to profession profiles for personalized output formatting.
 *
 * Features:
 * - Automatic profile loading from Microsoft Graph
 * - Caching with configurable TTL (default: 1 hour)
 * - Profession profile mapping based on jobTitle/department
 * - Support for manual profession override via chat preferences
 */

import logger from './logger.js';

/**
 * Detail level for output formatting
 * - executive: High-level summary for executives
 * - summary: Concise overview
 * - detailed: Comprehensive information
 * - technical: Full technical details
 */
export type DetailLevel = 'executive' | 'summary' | 'detailed' | 'technical';

/**
 * Language style for output formatting
 * - formal: Formal business language
 * - professional: Standard professional tone
 * - technical: Technical terminology
 * - customer-focused: Customer-oriented, persuasive
 */
export type LanguageStyle = 'formal' | 'professional' | 'technical' | 'customer-focused';

/**
 * Format preference for output structure
 * - bullet: Bullet points and highlights
 * - structured: Structured lists and sections
 * - narrative: Narrative text with context
 * - code-examples: Include code examples and structured data
 */
export type FormatPreference = 'bullet' | 'structured' | 'narrative' | 'code-examples';

/**
 * Profession profile for personalized output
 */
export interface ProfessionProfile {
  /** Unique identifier for the profession */
  id: string;
  /** Display name of the profession */
  name: string;
  /** Detail level for outputs */
  detailLevel: DetailLevel;
  /** Language style for text */
  languageStyle: LanguageStyle;
  /** Format preference for structure */
  formatPreference: FormatPreference;
  /** Keywords that indicate this profession */
  keywords: string[];
}

/**
 * User profile data from Microsoft Graph
 */
export interface UserProfile {
  /** User's unique ID (from Microsoft Graph) */
  userId: string;
  /** Display name */
  displayName?: string;
  /** Job title from Microsoft Graph */
  jobTitle?: string;
  /** Department from Microsoft Graph */
  department?: string;
  /** Email address */
  email?: string;
  /** Detected or overridden profession profile */
  professionProfile: ProfessionProfile;
  /** Whether the profession was manually overridden */
  isManualOverride: boolean;
  /** When the profile was last updated */
  lastUpdated: Date;
}

/**
 * Configuration for UserProfileService
 */
interface UserProfileServiceConfig {
  /** Cache TTL in milliseconds (default: 1 hour) */
  cacheTtlMs: number;
  /** Whether to auto-fetch profiles from Graph API */
  autoFetch: boolean;
}

/**
 * Default configuration
 */
const DEFAULT_CONFIG: UserProfileServiceConfig = {
  cacheTtlMs: 60 * 60 * 1000, // 1 hour
  autoFetch: true,
};

/**
 * Default profession profile (fallback)
 */
export const DEFAULT_PROFESSION_PROFILE: ProfessionProfile = {
  id: 'default',
  name: 'Professional',
  detailLevel: 'summary',
  languageStyle: 'professional',
  formatPreference: 'structured',
  keywords: [],
};

/**
 * Predefined profession profiles with their characteristics
 */
export const PROFESSION_PROFILES: Record<string, ProfessionProfile> = {
  executive: {
    id: 'executive',
    name: 'Executive/Manager',
    detailLevel: 'executive',
    languageStyle: 'formal',
    formatPreference: 'bullet',
    keywords: [
      'ceo',
      'cfo',
      'cto',
      'coo',
      'cio',
      'chief',
      'president',
      'vice president',
      'vp',
      'director',
      'head of',
      'geschäftsführer',
      'vorstand',
      'leiter',
      'abteilungsleiter',
      'bereichsleiter',
    ],
  },
  manager: {
    id: 'manager',
    name: 'Manager',
    detailLevel: 'executive',
    languageStyle: 'formal',
    formatPreference: 'bullet',
    keywords: [
      'manager',
      'team lead',
      'teamleiter',
      'projektleiter',
      'project manager',
      'product manager',
      'program manager',
      'supervisor',
      'coordinator',
      'lead',
    ],
  },
  developer: {
    id: 'developer',
    name: 'Developer/Engineer',
    detailLevel: 'technical',
    languageStyle: 'technical',
    formatPreference: 'code-examples',
    keywords: [
      'developer',
      'engineer',
      'programmer',
      'software',
      'entwickler',
      'softwareentwickler',
      'architekt',
      'architect',
      'devops',
      'sre',
      'backend',
      'frontend',
      'full stack',
      'fullstack',
      'data engineer',
      'ml engineer',
      'ai engineer',
    ],
  },
  it: {
    id: 'it',
    name: 'IT Professional',
    detailLevel: 'detailed',
    languageStyle: 'technical',
    formatPreference: 'structured',
    keywords: [
      'it',
      'system',
      'administrator',
      'admin',
      'support',
      'technician',
      'analyst',
      'specialist',
      'security',
      'network',
      'infrastructure',
      'cloud',
      'helpdesk',
    ],
  },
  sales: {
    id: 'sales',
    name: 'Sales',
    detailLevel: 'summary',
    languageStyle: 'customer-focused',
    formatPreference: 'structured',
    keywords: [
      'sales',
      'account',
      'vertrieb',
      'verkauf',
      'business development',
      'bd',
      'customer success',
      'key account',
      'sales representative',
      'account executive',
      'ae',
    ],
  },
  marketing: {
    id: 'marketing',
    name: 'Marketing',
    detailLevel: 'summary',
    languageStyle: 'customer-focused',
    formatPreference: 'narrative',
    keywords: [
      'marketing',
      'brand',
      'content',
      'social media',
      'communications',
      'pr',
      'public relations',
      'campaign',
      'growth',
      'digital marketing',
    ],
  },
  hr: {
    id: 'hr',
    name: 'Human Resources',
    detailLevel: 'summary',
    languageStyle: 'professional',
    formatPreference: 'structured',
    keywords: [
      'hr',
      'human resources',
      'recruiting',
      'recruiter',
      'talent',
      'people',
      'personal',
      'personalabteilung',
      'personalwesen',
    ],
  },
  finance: {
    id: 'finance',
    name: 'Finance',
    detailLevel: 'detailed',
    languageStyle: 'formal',
    formatPreference: 'structured',
    keywords: [
      'finance',
      'finanzen',
      'accounting',
      'buchhaltung',
      'controller',
      'controlling',
      'treasurer',
      'auditor',
      'financial',
      'rechnungswesen',
    ],
  },
  legal: {
    id: 'legal',
    name: 'Legal',
    detailLevel: 'detailed',
    languageStyle: 'formal',
    formatPreference: 'structured',
    keywords: [
      'legal',
      'lawyer',
      'attorney',
      'counsel',
      'rechtsanwalt',
      'jurist',
      'compliance',
      'paralegal',
      'contract',
    ],
  },
  research: {
    id: 'research',
    name: 'Research/Science',
    detailLevel: 'technical',
    languageStyle: 'technical',
    formatPreference: 'structured',
    keywords: [
      'research',
      'scientist',
      'researcher',
      'forscher',
      'wissenschaftler',
      'analyst',
      'data scientist',
      'phd',
      'dr.',
      'professor',
    ],
  },
  consulting: {
    id: 'consulting',
    name: 'Consultant',
    detailLevel: 'detailed',
    languageStyle: 'professional',
    formatPreference: 'structured',
    keywords: [
      'consultant',
      'berater',
      'advisor',
      'consulting',
      'beratung',
      'strategy',
      'strategieberater',
    ],
  },
  design: {
    id: 'design',
    name: 'Design/Creative',
    detailLevel: 'summary',
    languageStyle: 'professional',
    formatPreference: 'narrative',
    keywords: [
      'design',
      'designer',
      'ux',
      'ui',
      'creative',
      'art director',
      'graphic',
      'visual',
      'gestalter',
    ],
  },
  operations: {
    id: 'operations',
    name: 'Operations',
    detailLevel: 'detailed',
    languageStyle: 'professional',
    formatPreference: 'structured',
    keywords: [
      'operations',
      'ops',
      'logistics',
      'supply chain',
      'procurement',
      'einkauf',
      'logistik',
      'betrieb',
    ],
  },
  default: DEFAULT_PROFESSION_PROFILE,
};

/**
 * Cached user profile entry
 */
interface CachedProfile {
  profile: UserProfile;
  expiresAt: Date;
}

/**
 * User Profile Service - Singleton for managing user profiles
 */
export class UserProfileService {
  private static instance: UserProfileService | null = null;
  private cache: Map<string, CachedProfile> = new Map();
  private config: UserProfileServiceConfig;

  private constructor(config?: Partial<UserProfileServiceConfig>) {
    this.config = {
      ...DEFAULT_CONFIG,
      ...config,
    };
    logger.info('UserProfileService initialized', {
      cacheTtlMs: this.config.cacheTtlMs,
      autoFetch: this.config.autoFetch,
    });
  }

  /**
   * Get the singleton instance
   */
  static getInstance(config?: Partial<UserProfileServiceConfig>): UserProfileService {
    if (!UserProfileService.instance) {
      UserProfileService.instance = new UserProfileService(config);
    }
    return UserProfileService.instance;
  }

  /**
   * Reset the singleton instance (for testing)
   */
  static resetInstance(): void {
    UserProfileService.instance = null;
  }

  /**
   * Detect profession profile from job title and department
   *
   * @param jobTitle - User's job title from Microsoft Graph
   * @param department - User's department from Microsoft Graph
   * @returns Matched profession profile or default
   */
  detectProfessionProfile(jobTitle?: string, department?: string): ProfessionProfile {
    const searchText = `${jobTitle || ''} ${department || ''}`.toLowerCase();

    if (!searchText.trim()) {
      logger.debug('No jobTitle or department provided, using default profile');
      return DEFAULT_PROFESSION_PROFILE;
    }

    // Check each profession profile for keyword matches
    // Priority: more specific matches first
    const profileScores: Array<{ profile: ProfessionProfile; score: number }> = [];

    for (const [id, profile] of Object.entries(PROFESSION_PROFILES)) {
      if (id === 'default') continue;

      let score = 0;
      for (const keyword of profile.keywords) {
        if (searchText.includes(keyword.toLowerCase())) {
          // Longer keywords get higher scores (more specific)
          score += keyword.length;
        }
      }

      if (score > 0) {
        profileScores.push({ profile, score });
      }
    }

    // Sort by score (highest first) and return best match
    if (profileScores.length > 0) {
      profileScores.sort((a, b) => b.score - a.score);
      const bestMatch = profileScores[0].profile;
      logger.debug(`Detected profession profile: ${bestMatch.id}`, {
        jobTitle,
        department,
        score: profileScores[0].score,
      });
      return bestMatch;
    }

    logger.debug('No profession profile matched, using default', { jobTitle, department });
    return DEFAULT_PROFESSION_PROFILE;
  }

  /**
   * Get profession profile by ID
   *
   * @param professionId - Profession profile ID
   * @returns Profession profile or default if not found
   */
  getProfessionProfileById(professionId: string): ProfessionProfile {
    const profile = PROFESSION_PROFILES[professionId.toLowerCase()];
    if (profile) {
      return profile;
    }
    logger.warn(`Unknown profession profile ID: ${professionId}, using default`);
    return DEFAULT_PROFESSION_PROFILE;
  }

  /**
   * Get cached profile for a user
   *
   * @param userId - User ID
   * @returns Cached profile or undefined if not cached or expired
   */
  getCachedProfile(userId: string): UserProfile | undefined {
    const cached = this.cache.get(userId);
    if (!cached) {
      return undefined;
    }

    if (new Date() > cached.expiresAt) {
      logger.debug('Cached profile expired', { userId: userId.substring(0, 8) });
      this.cache.delete(userId);
      return undefined;
    }

    logger.debug('Using cached profile', { userId: userId.substring(0, 8) });
    return cached.profile;
  }

  /**
   * Cache a user profile
   *
   * @param profile - User profile to cache
   */
  cacheProfile(profile: UserProfile): void {
    const expiresAt = new Date(Date.now() + this.config.cacheTtlMs);
    this.cache.set(profile.userId, { profile, expiresAt });
    logger.debug('Cached user profile', {
      userId: profile.userId.substring(0, 8),
      professionId: profile.professionProfile.id,
      expiresAt: expiresAt.toISOString(),
    });
  }

  /**
   * Create a user profile from Microsoft Graph data
   *
   * @param graphData - Data from Microsoft Graph /me endpoint
   * @param professionOverride - Optional manual profession override
   * @returns UserProfile
   */
  createProfileFromGraphData(
    graphData: {
      id: string;
      displayName?: string;
      jobTitle?: string;
      department?: string;
      mail?: string;
      userPrincipalName?: string;
    },
    professionOverride?: string
  ): UserProfile {
    let professionProfile: ProfessionProfile;
    let isManualOverride = false;

    if (professionOverride) {
      professionProfile = this.getProfessionProfileById(professionOverride);
      isManualOverride = true;
      logger.info('Using manual profession override', {
        userId: graphData.id.substring(0, 8),
        professionId: professionProfile.id,
      });
    } else {
      professionProfile = this.detectProfessionProfile(graphData.jobTitle, graphData.department);
    }

    const profile: UserProfile = {
      userId: graphData.id,
      displayName: graphData.displayName,
      jobTitle: graphData.jobTitle,
      department: graphData.department,
      email: graphData.mail || graphData.userPrincipalName,
      professionProfile,
      isManualOverride,
      lastUpdated: new Date(),
    };

    // Cache the profile
    this.cacheProfile(profile);

    return profile;
  }

  /**
   * Update profession override for a user
   *
   * @param userId - User ID
   * @param professionId - New profession profile ID (or null to clear override)
   * @returns Updated profile or undefined if user not cached
   */
  updateProfessionOverride(userId: string, professionId: string | null): UserProfile | undefined {
    const cached = this.cache.get(userId);
    if (!cached) {
      logger.warn('Cannot update profession override - user not in cache', {
        userId: userId.substring(0, 8),
      });
      return undefined;
    }

    const profile = cached.profile;

    if (professionId === null) {
      // Clear override - re-detect from job title
      profile.professionProfile = this.detectProfessionProfile(
        profile.jobTitle,
        profile.department
      );
      profile.isManualOverride = false;
      logger.info('Cleared profession override', { userId: userId.substring(0, 8) });
    } else {
      // Set override
      profile.professionProfile = this.getProfessionProfileById(professionId);
      profile.isManualOverride = true;
      logger.info('Set profession override', {
        userId: userId.substring(0, 8),
        professionId: profile.professionProfile.id,
      });
    }

    profile.lastUpdated = new Date();
    this.cacheProfile(profile);

    return profile;
  }

  /**
   * Clear cache for a specific user
   *
   * @param userId - User ID
   */
  clearCache(userId: string): void {
    this.cache.delete(userId);
    logger.debug('Cleared cache for user', { userId: userId.substring(0, 8) });
  }

  /**
   * Clear all cached profiles
   */
  clearAllCache(): void {
    this.cache.clear();
    logger.info('Cleared all user profile cache');
  }

  /**
   * Get all available profession profiles
   */
  getAvailableProfessionProfiles(): ProfessionProfile[] {
    return Object.values(PROFESSION_PROFILES);
  }

  /**
   * Get cache statistics
   */
  getCacheStats(): { size: number; profiles: string[] } {
    return {
      size: this.cache.size,
      profiles: Array.from(this.cache.keys()).map((k) => k.substring(0, 8) + '...'),
    };
  }
}

/**
 * Get the UserProfileService singleton instance
 */
export function getUserProfileService(
  config?: Partial<UserProfileServiceConfig>
): UserProfileService {
  return UserProfileService.getInstance(config);
}

export default UserProfileService;
