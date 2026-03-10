/**
 * MCP tools that expose server capabilities and example questions.
 * Use when the user asks "Was kannst du?" / "What can you do?" / "Welche Fragen kann ich dir stellen?"
 */

import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { exampleQuestionsCategories } from './data/example-questions.js';

export function registerCapabilityTools(server: McpServer): void {
  server.registerTool(
    'get-example-questions',
    {
      title: 'get-example-questions',
      description:
        'Returns example questions that this Microsoft 365 MCP server can answer 100% using its tools. ' +
        'Call this when the user asks "What can you do?", "Was kannst du?", "Welche Fragen kann ich dir stellen?" or similar capability questions. ' +
        'Use the returned categories and questions to formulate a helpful answer with concrete examples.',
      inputSchema: z.object({
        language: z
          .enum(['de', 'en', 'both'])
          .optional()
          .default('both')
          .describe('Language for question text: "de", "en", or "both" (default).'),
      }),
    },
    async ({ language }) => {
      const categories = exampleQuestionsCategories.map((cat) => ({
        id: cat.id,
        name: language === 'de' ? cat.nameDe : language === 'en' ? cat.nameEn : undefined,
        nameDe: cat.nameDe,
        nameEn: cat.nameEn,
        questions:
          language === 'de'
            ? cat.questions.map((q) => q.de)
            : language === 'en'
              ? cat.questions.map((q) => q.en)
              : cat.questions.map((q) => ({ de: q.de, en: q.en })),
      }));

      return {
        content: [
          {
            type: 'text' as const,
            text: JSON.stringify(
              {
                purpose:
                  'Example questions this MCP server can answer 100% using Microsoft 365 (Outlook, Calendar, Teams, OneDrive, SharePoint, To-Do, Planner, OneNote, Search).',
                categories,
              },
              null,
              2
            ),
          },
        ],
      };
    }
  );
}
