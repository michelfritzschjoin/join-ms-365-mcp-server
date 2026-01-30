/**
 * Entity Type Validation Utility
 *
 * Validates and filters entity types according to Microsoft Graph API compatibility rules.
 *
 * Microsoft Graph API has strict rules about which entity types can be combined:
 * - Answer types (acronym, bookmark, qna): Can only combine with each other
 * - Message types (message, chatMessage): Can only combine with each other
 * - File types (drive, driveItem, list, listItem, site, externalItem): Can combine with each other
 * - Standalone types (event, person): Cannot combine with anything
 *
 * Reference: https://learn.microsoft.com/en-us/graph/api/resources/search-api-overview#known-limitations
 */

/**
 * Validate and filter entity types according to Microsoft Graph API compatibility rules
 *
 * This function ensures only compatible entity types are returned.
 * If incompatible types are detected, it prioritizes the most common types.
 *
 * @param entityTypes - Array of entity type strings to validate
 * @returns Filtered array containing only compatible entity types
 */
export function validateEntityTypeCombinations(entityTypes: string[]): string[] {
  if (entityTypes.length === 0) {
    return [];
  }

  // Define compatibility groups based on Microsoft Graph API documentation
  const answerTypes = new Set(['acronym', 'bookmark', 'qna']);
  const messageTypes = new Set(['message', 'chatMessage']);
  const fileTypes = new Set(['drive', 'driveItem', 'list', 'listItem', 'site', 'externalItem']);
  const standaloneTypes = new Set(['event', 'person']);

  // Categorize input types
  const hasAnswerTypes = entityTypes.some((t) => answerTypes.has(t));
  const hasMessageTypes = entityTypes.some((t) => messageTypes.has(t));
  const hasFileTypes = entityTypes.some((t) => fileTypes.has(t));
  const hasStandaloneTypes = entityTypes.some((t) => standaloneTypes.has(t));

  // Count how many groups are present
  const groupCount = [hasAnswerTypes, hasMessageTypes, hasFileTypes, hasStandaloneTypes].filter(
    Boolean
  ).length;

  // If only one group or standalone types, return as-is (but filter standalone if multiple)
  if (groupCount <= 1) {
    // If we have standalone types with others, remove standalone
    if (hasStandaloneTypes && entityTypes.length > 1) {
      return entityTypes.filter((t) => !standaloneTypes.has(t));
    }
    return entityTypes;
  }

  // Multiple incompatible groups detected - prioritize based on common use cases
  // Priority: File types > Message types > Answer types > Standalone types
  const filtered: string[] = [];

  if (hasFileTypes) {
    // File types are most common and useful - keep all file types
    filtered.push(...entityTypes.filter((t) => fileTypes.has(t)));
  }

  if (hasMessageTypes && !hasFileTypes) {
    // Message types if no file types
    filtered.push(...entityTypes.filter((t) => messageTypes.has(t)));
  }

  if (hasAnswerTypes && !hasFileTypes && !hasMessageTypes) {
    // Answer types only if no other groups
    filtered.push(...entityTypes.filter((t) => answerTypes.has(t)));
  }

  // Never include standalone types when mixing with others
  // (They're already filtered out above)

  return filtered.length > 0 ? filtered : entityTypes.slice(0, 1); // Fallback to first type
}
