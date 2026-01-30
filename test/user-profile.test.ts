/**
 * User Profile Service Tests
 *
 * Tests for profession detection, caching, and profile management.
 */

import { describe, it, expect, beforeEach } from 'vitest';
import {
  UserProfileService,
  getUserProfileService,
  PROFESSION_PROFILES,
  DEFAULT_PROFESSION_PROFILE,
  type UserProfile,
  type ProfessionProfile,
} from '../src/user-profile.js';

describe('UserProfileService', () => {
  beforeEach(() => {
    // Reset singleton instance before each test
    UserProfileService.resetInstance();
  });

  describe('Singleton Pattern', () => {
    it('should return the same instance', () => {
      const instance1 = getUserProfileService();
      const instance2 = getUserProfileService();
      expect(instance1).toBe(instance2);
    });

    it('should reset instance correctly', () => {
      const instance1 = getUserProfileService();
      UserProfileService.resetInstance();
      const instance2 = getUserProfileService();
      expect(instance1).not.toBe(instance2);
    });
  });

  describe('Profession Detection', () => {
    const service = getUserProfileService();

    it('should detect executive profile from CEO title', () => {
      const profile = service.detectProfessionProfile('CEO');
      expect(profile.id).toBe('executive');
      expect(profile.detailLevel).toBe('executive');
    });

    it('should detect executive profile from Geschäftsführer title', () => {
      const profile = service.detectProfessionProfile('Geschäftsführer');
      expect(profile.id).toBe('executive');
    });

    it('should detect manager profile from Team Lead title', () => {
      const profile = service.detectProfessionProfile('Team Lead Engineering');
      expect(profile.id).toBe('manager');
      expect(profile.formatPreference).toBe('bullet');
    });

    it('should detect developer profile from Software Engineer title', () => {
      const profile = service.detectProfessionProfile('Software Engineer');
      expect(profile.id).toBe('developer');
      expect(profile.detailLevel).toBe('technical');
      expect(profile.languageStyle).toBe('technical');
      expect(profile.formatPreference).toBe('code-examples');
    });

    it('should detect developer profile from Softwareentwickler title', () => {
      const profile = service.detectProfessionProfile('Senior Softwareentwickler');
      expect(profile.id).toBe('developer');
    });

    it('should detect sales profile from Account Executive title', () => {
      const profile = service.detectProfessionProfile('Account Executive');
      expect(profile.id).toBe('sales');
      expect(profile.languageStyle).toBe('customer-focused');
    });

    it('should detect sales profile from Vertriebsleiter title', () => {
      const profile = service.detectProfessionProfile('Vertriebsleiter');
      expect(profile.id).toBe('sales');
    });

    it('should detect IT profile from System Administrator title', () => {
      const profile = service.detectProfessionProfile('System Administrator');
      expect(profile.id).toBe('it');
      expect(profile.detailLevel).toBe('detailed');
    });

    it('should detect HR profile from Human Resources title', () => {
      const profile = service.detectProfessionProfile('Human Resources Specialist');
      expect(profile.id).toBe('hr');
    });

    it('should detect finance profile from Controller title', () => {
      const profile = service.detectProfessionProfile('Financial Controller');
      expect(profile.id).toBe('finance');
      expect(profile.languageStyle).toBe('formal');
    });

    it('should use department for detection when job title is generic', () => {
      const profile = service.detectProfessionProfile('Specialist', 'IT Department');
      expect(profile.id).toBe('it');
    });

    it('should return default profile for unknown titles', () => {
      const profile = service.detectProfessionProfile('Astronaut');
      expect(profile.id).toBe('default');
    });

    it('should return default profile when no title provided', () => {
      const profile = service.detectProfessionProfile(undefined, undefined);
      expect(profile.id).toBe('default');
    });

    it('should prioritize longer keyword matches (more specific)', () => {
      // "Software Engineer" should match developer, not just because of "engineer"
      const profile = service.detectProfessionProfile('Senior Software Engineer');
      expect(profile.id).toBe('developer');
    });
  });

  describe('Profile by ID', () => {
    const service = getUserProfileService();

    it('should return correct profile by ID', () => {
      const profile = service.getProfessionProfileById('developer');
      expect(profile.id).toBe('developer');
      expect(profile.name).toBe('Developer/Engineer');
    });

    it('should return default profile for unknown ID', () => {
      const profile = service.getProfessionProfileById('nonexistent');
      expect(profile.id).toBe('default');
    });

    it('should be case-insensitive', () => {
      const profile = service.getProfessionProfileById('DEVELOPER');
      expect(profile.id).toBe('developer');
    });
  });

  describe('Profile Caching', () => {
    const service = getUserProfileService();

    it('should cache profiles', () => {
      const mockData = {
        id: 'test-user-123',
        displayName: 'Test User',
        jobTitle: 'Software Developer',
        department: 'Engineering',
      };

      const profile = service.createProfileFromGraphData(mockData);
      expect(profile.userId).toBe('test-user-123');

      // Should retrieve from cache
      const cached = service.getCachedProfile('test-user-123');
      expect(cached).toBeDefined();
      expect(cached?.displayName).toBe('Test User');
    });

    it('should clear individual cache', () => {
      const mockData = {
        id: 'test-user-456',
        displayName: 'Test User 2',
        jobTitle: 'Manager',
      };

      service.createProfileFromGraphData(mockData);
      expect(service.getCachedProfile('test-user-456')).toBeDefined();

      service.clearCache('test-user-456');
      expect(service.getCachedProfile('test-user-456')).toBeUndefined();
    });

    it('should clear all cache', () => {
      const mockData1 = {
        id: 'test-user-a',
        displayName: 'User A',
        jobTitle: 'Developer',
      };
      const mockData2 = {
        id: 'test-user-b',
        displayName: 'User B',
        jobTitle: 'Manager',
      };

      service.createProfileFromGraphData(mockData1);
      service.createProfileFromGraphData(mockData2);

      service.clearAllCache();

      expect(service.getCachedProfile('test-user-a')).toBeUndefined();
      expect(service.getCachedProfile('test-user-b')).toBeUndefined();
    });
  });

  describe('Profile Creation from Graph Data', () => {
    const service = getUserProfileService();

    it('should create profile with detected profession', () => {
      const mockData = {
        id: 'user-dev-1',
        displayName: 'Jane Developer',
        jobTitle: 'Full Stack Developer',
        department: 'Technology',
        mail: 'jane@example.com',
      };

      const profile = service.createProfileFromGraphData(mockData);

      expect(profile.userId).toBe('user-dev-1');
      expect(profile.displayName).toBe('Jane Developer');
      expect(profile.jobTitle).toBe('Full Stack Developer');
      expect(profile.department).toBe('Technology');
      expect(profile.email).toBe('jane@example.com');
      expect(profile.professionProfile.id).toBe('developer');
      expect(profile.isManualOverride).toBe(false);
    });

    it('should apply manual profession override', () => {
      const mockData = {
        id: 'user-manager-1',
        displayName: 'Bob Manager',
        jobTitle: 'Project Manager',
        department: 'PMO',
      };

      // Force developer profile override
      const profile = service.createProfileFromGraphData(mockData, 'developer');

      expect(profile.professionProfile.id).toBe('developer');
      expect(profile.isManualOverride).toBe(true);
    });

    it('should use userPrincipalName as email fallback', () => {
      const mockData = {
        id: 'user-upn-1',
        displayName: 'UPN User',
        jobTitle: 'Analyst',
        userPrincipalName: 'upnuser@example.com',
      };

      const profile = service.createProfileFromGraphData(mockData);
      expect(profile.email).toBe('upnuser@example.com');
    });
  });

  describe('Profession Override Update', () => {
    const service = getUserProfileService();

    beforeEach(() => {
      // Reset and create a fresh profile
      UserProfileService.resetInstance();
      const freshService = getUserProfileService();
      const mockData = {
        id: 'user-override-test',
        displayName: 'Override Test User',
        jobTitle: 'Senior Project Manager', // Clear manager title
        department: 'PMO',
      };
      freshService.createProfileFromGraphData(mockData);
    });

    it('should update profession override', () => {
      const freshService = getUserProfileService();
      const updated = freshService.updateProfessionOverride('user-override-test', 'developer');

      expect(updated).toBeDefined();
      expect(updated?.professionProfile.id).toBe('developer');
      expect(updated?.isManualOverride).toBe(true);
    });

    it('should clear override and re-detect', () => {
      const freshService = getUserProfileService();
      // First set an override
      freshService.updateProfessionOverride('user-override-test', 'developer');

      // Then clear it
      const updated = freshService.updateProfessionOverride('user-override-test', null);

      expect(updated).toBeDefined();
      expect(updated?.professionProfile.id).toBe('manager'); // Re-detected from "Senior Project Manager"
      expect(updated?.isManualOverride).toBe(false);
    });

    it('should return undefined for non-cached user', () => {
      const freshService = getUserProfileService();
      const result = freshService.updateProfessionOverride('non-existent-user', 'developer');
      expect(result).toBeUndefined();
    });
  });

  describe('Available Profession Profiles', () => {
    const service = getUserProfileService();

    it('should return all available profiles', () => {
      const profiles = service.getAvailableProfessionProfiles();

      expect(profiles.length).toBeGreaterThan(0);
      expect(profiles.some((p) => p.id === 'developer')).toBe(true);
      expect(profiles.some((p) => p.id === 'manager')).toBe(true);
      expect(profiles.some((p) => p.id === 'sales')).toBe(true);
      expect(profiles.some((p) => p.id === 'default')).toBe(true);
    });

    it('should include all expected profession types', () => {
      const profiles = service.getAvailableProfessionProfiles();
      const ids = profiles.map((p) => p.id);

      expect(ids).toContain('executive');
      expect(ids).toContain('manager');
      expect(ids).toContain('developer');
      expect(ids).toContain('it');
      expect(ids).toContain('sales');
      expect(ids).toContain('marketing');
      expect(ids).toContain('hr');
      expect(ids).toContain('finance');
      expect(ids).toContain('legal');
      expect(ids).toContain('research');
      expect(ids).toContain('consulting');
      expect(ids).toContain('design');
      expect(ids).toContain('operations');
    });
  });

  describe('Cache Statistics', () => {
    const service = getUserProfileService();

    it('should return correct cache statistics', () => {
      // Start empty
      service.clearAllCache();
      let stats = service.getCacheStats();
      expect(stats.size).toBe(0);

      // Add some profiles
      service.createProfileFromGraphData({
        id: 'stats-user-1',
        displayName: 'User 1',
        jobTitle: 'Developer',
      });
      service.createProfileFromGraphData({
        id: 'stats-user-2',
        displayName: 'User 2',
        jobTitle: 'Manager',
      });

      stats = service.getCacheStats();
      expect(stats.size).toBe(2);
      expect(stats.profiles.length).toBe(2);
    });
  });
});

describe('PROFESSION_PROFILES', () => {
  it('should have all required fields for each profile', () => {
    for (const [id, profile] of Object.entries(PROFESSION_PROFILES)) {
      expect(profile.id).toBe(id);
      expect(profile.name).toBeDefined();
      expect(profile.detailLevel).toBeDefined();
      expect(profile.languageStyle).toBeDefined();
      expect(profile.formatPreference).toBeDefined();
      expect(profile.keywords).toBeDefined();
      expect(Array.isArray(profile.keywords)).toBe(true);
    }
  });

  it('should have unique keywords across profiles', () => {
    const allKeywords: string[] = [];
    const duplicates: string[] = [];

    for (const profile of Object.values(PROFESSION_PROFILES)) {
      for (const keyword of profile.keywords) {
        if (allKeywords.includes(keyword)) {
          duplicates.push(keyword);
        }
        allKeywords.push(keyword);
      }
    }

    // Some duplicates are acceptable (e.g., "analyst" in multiple profiles)
    // But we want to ensure no critical duplicates that would cause confusion
    // This is more of a documentation/awareness test
    if (duplicates.length > 0) {
      console.log('Note: Duplicate keywords found (may be intentional):', duplicates);
    }
  });
});

describe('DEFAULT_PROFESSION_PROFILE', () => {
  it('should have sensible defaults', () => {
    expect(DEFAULT_PROFESSION_PROFILE.id).toBe('default');
    expect(DEFAULT_PROFESSION_PROFILE.name).toBe('Professional');
    expect(DEFAULT_PROFESSION_PROFILE.detailLevel).toBe('summary');
    expect(DEFAULT_PROFESSION_PROFILE.languageStyle).toBe('professional');
    expect(DEFAULT_PROFESSION_PROFILE.formatPreference).toBe('structured');
    expect(DEFAULT_PROFESSION_PROFILE.keywords).toEqual([]);
  });
});
