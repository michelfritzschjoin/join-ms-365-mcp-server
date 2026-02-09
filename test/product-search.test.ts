import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import type GraphClient from '../src/graph-client.js';
import { registerSuperTools } from '../src/super-tools.js';
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';

// Mock dependencies
vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
    debug: vi.fn(),
  },
}));

vi.mock('../src/thinking-process.js', () => ({
  addThinkingToResponse: (response: string) => response,
  isThinkingEnabled: () => false,
}));

vi.mock('../src/request-context.js', () => ({
  getRequestTokens: vi.fn(() => ({ accessToken: 'mock-token' })),
  getUserId: vi.fn(() => 'mock-user-id'),
  getUserProfile: vi.fn(() => null),
  getProfessionProfile: vi.fn(() => null),
}));

vi.mock('../src/query-store.js', () => ({
  getQueryStore: vi.fn(() => ({
    hashUserId: vi.fn(() => 'hashed-id'),
    recordQueryPattern: vi.fn(),
  })),
}));

vi.mock('../src/uqas/integration/index.js', () => ({
  getUQAS: vi.fn(() => ({
    analyzeQuery: vi.fn(() => ({
      language: 'en',
      languageConfidence: 0.9,
      variants: [],
      crossLangVariants: [],
    })),
    setDownloadLinkGenerator: vi.fn(),
  })),
}));

describe('Product Search Tool', () => {
  let mockServer: McpServer;
  let mockGraphClient: GraphClient;
  let productSearchHandler: (input: {
    action: string;
    query: string;
    maxResults?: number;
    topPerProduct?: number;
  }) => Promise<{ content: Array<{ type: string; text: string }>; isError?: boolean }>;

  beforeEach(() => {
    vi.clearAllMocks();

    // Create mock server
    mockServer = {
      tool: vi.fn((name, description, schema, handler) => {
        if (name === 'product-search') {
          productSearchHandler = handler as typeof productSearchHandler;
        }
      }),
    } as unknown as McpServer;

    // Create mock graph client
    mockGraphClient = {
      makeRequest: vi.fn(),
    } as unknown as GraphClient;

    // Register tools
    registerSuperTools(mockServer, mockGraphClient, true);
  });

  afterEach(() => {
    vi.resetAllMocks();
  });

  describe('Query Tests - Different Products', () => {
    it('should return results for Outlook/Email queries', async () => {
      // Mock initial search response with email results
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 5,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.message',
                      id: 'msg-1',
                      subject: 'Test Email',
                      bodyPreview: 'This is a test email',
                    },
                    summary: 'Test Email',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      // Mock Outlook-specific search response
      const mockOutlookResponse = {
        value: [
          {
            id: 'msg-1',
            subject: 'Test Email',
            from: { emailAddress: { address: 'test@example.com' } },
            receivedDateTime: '2024-01-01T10:00:00Z',
            bodyPreview: 'This is a test email',
            webLink: 'https://outlook.office.com/mail/id/msg-1',
            hasAttachments: false,
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockOutlookResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'test email',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      // Validate response structure
      expect(response).toHaveProperty('query', 'test email');
      expect(response).toHaveProperty('initialSearchResults');
      expect(response).toHaveProperty('productResults');
      expect(response).toHaveProperty('thinking');

      // Validate initial search results
      expect(response.initialSearchResults).toHaveProperty('totalHits');
      expect(response.initialSearchResults).toHaveProperty('productsDetected');
      expect(response.initialSearchResults.productsDetected).toContain('Outlook');

      // Validate product results
      expect(Array.isArray(response.productResults)).toBe(true);
      const outlookResult = response.productResults.find(
        (r: { product: string }) => r.product === 'Outlook'
      );
      expect(outlookResult).toBeDefined();
      expect(outlookResult.resultCount).toBeGreaterThan(0);
      expect(Array.isArray(outlookResult.topResults)).toBe(true);
    });

    it('should return results for Calendar/Event queries', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 3,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.event',
                      id: 'event-1',
                      subject: 'Team Meeting',
                    },
                    summary: 'Team Meeting',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockCalendarResponse = {
        value: [
          {
            id: 'event-1',
            subject: 'Team Meeting',
            start: { dateTime: '2024-01-15T10:00:00Z', timeZone: 'UTC' },
            end: { dateTime: '2024-01-15T11:00:00Z', timeZone: 'UTC' },
            location: { displayName: 'Conference Room A' },
            organizer: { emailAddress: { address: 'organizer@example.com' } },
            webLink: 'https://outlook.office.com/calendar/id/event-1',
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockCalendarResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'meeting', // Use "meeting" which matches "Team Meeting"
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      expect(response.initialSearchResults.productsDetected).toContain('Calendar');
      const calendarResult = response.productResults.find(
        (r: { product: string }) => r.product === 'Calendar'
      );
      expect(calendarResult).toBeDefined();
      expect(calendarResult.resultCount).toBeGreaterThan(0);
    });

    it('should return results for OneDrive/File queries', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 10,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.driveItem',
                      id: 'file-1',
                      name: 'document.pdf',
                    },
                    summary: 'document.pdf',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockOneDriveResponse = {
        value: [
          {
            id: 'file-1',
            name: 'document.pdf',
            webUrl: 'https://onedrive.live.com/file-1',
            size: 1024,
            lastModifiedDateTime: '2024-01-01T10:00:00Z',
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockOneDriveResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'document pdf',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      expect(response.initialSearchResults.productsDetected).toContain('OneDrive');
      const oneDriveResult = response.productResults.find(
        (r: { product: string }) => r.product === 'OneDrive'
      );
      expect(oneDriveResult).toBeDefined();
    });

    it('should return results for SharePoint queries', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 7,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.site',
                      id: 'site-1',
                      displayName: 'Project Site',
                    },
                    summary: 'Project Site',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockSharePointResponse = {
        value: [
          {
            id: 'site-1',
            displayName: 'Project Site',
            webUrl: 'https://contoso.sharepoint.com/sites/project',
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockSharePointResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'project site',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      expect(response.initialSearchResults.productsDetected).toContain('SharePoint');
      const sharePointResult = response.productResults.find(
        (r: { product: string }) => r.product === 'SharePoint'
      );
      expect(sharePointResult).toBeDefined();
    });

    it('should return results for Teams/Chat queries', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 4,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.chatMessage',
                      id: 'msg-1',
                      body: { content: 'Team discussion' },
                    },
                    summary: 'Team discussion',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockTeamsResponse = {
        value: [
          {
            id: 'chat-1',
            topic: 'Team Discussion',
            webUrl: 'https://teams.microsoft.com/chat/chat-1',
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockTeamsResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'team discussion',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      expect(response.initialSearchResults.productsDetected).toContain('Teams');
      const teamsResult = response.productResults.find(
        (r: { product: string }) => r.product === 'Teams'
      );
      expect(teamsResult).toBeDefined();
    });

    it('should return results for User/Person queries', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 2,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.person',
                      id: 'person-1',
                      displayName: 'John Doe',
                    },
                    summary: 'John Doe',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockUsersResponse = {
        value: [
          {
            id: 'user-1',
            displayName: 'John Doe',
            mail: 'john.doe@example.com',
            jobTitle: 'Developer',
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockUsersResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'John Doe',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      expect(response.initialSearchResults.productsDetected).toContain('Users');
      const usersResult = response.productResults.find(
        (r: { product: string }) => r.product === 'Users'
      );
      expect(usersResult).toBeDefined();
    });
  });

  describe('Query Tests - Multiple Products', () => {
    it('should handle queries that return results from multiple products', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 15,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.message',
                      id: 'msg-1',
                      subject: 'Project Update',
                    },
                    summary: 'Project Update',
                    rank: 1,
                  },
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.event',
                      id: 'event-1',
                      subject: 'Project Meeting',
                    },
                    summary: 'Project Meeting',
                    rank: 2,
                  },
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.driveItem',
                      id: 'file-1',
                      name: 'project-doc.pdf',
                    },
                    summary: 'project-doc.pdf',
                    rank: 3,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockOutlookResponse = {
        value: [
          {
            id: 'msg-1',
            subject: 'Project Update',
            from: { emailAddress: { address: 'team@example.com' } },
            receivedDateTime: '2024-01-01T10:00:00Z',
            bodyPreview: 'Project update email',
            webLink: 'https://outlook.office.com/mail/id/msg-1',
            hasAttachments: false,
          },
        ],
      };

      const mockCalendarResponse = {
        value: [
          {
            id: 'event-1',
            subject: 'Project Meeting',
            start: { dateTime: '2024-01-15T10:00:00Z', timeZone: 'UTC' },
            end: { dateTime: '2024-01-15T11:00:00Z', timeZone: 'UTC' },
            webLink: 'https://outlook.office.com/calendar/id/event-1',
          },
        ],
      };

      const mockOneDriveResponse = {
        value: [
          {
            id: 'file-1',
            name: 'project-doc.pdf',
            webUrl: 'https://onedrive.live.com/file-1',
            size: 2048,
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockOutlookResponse))
        .mockResolvedValueOnce(JSON.stringify(mockCalendarResponse))
        .mockResolvedValueOnce(JSON.stringify(mockOneDriveResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'project',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      // Should detect multiple products
      expect(response.initialSearchResults.productsDetected.length).toBeGreaterThan(1);
      expect(response.initialSearchResults.productsDetected).toContain('Outlook');
      expect(response.initialSearchResults.productsDetected).toContain('Calendar');
      expect(response.initialSearchResults.productsDetected).toContain('OneDrive');

      // Should have results for multiple products
      expect(response.productResults.length).toBeGreaterThan(1);
      expect(
        response.productResults.some((r: { product: string }) => r.product === 'Outlook')
      ).toBe(true);
      expect(
        response.productResults.some((r: { product: string }) => r.product === 'Calendar')
      ).toBe(true);
      expect(
        response.productResults.some((r: { product: string }) => r.product === 'OneDrive')
      ).toBe(true);
    });
  });

  describe('Query Tests - Edge Cases', () => {
    it('should handle queries with no results', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 0,
                hits: [],
              },
            ],
          },
        ],
      };

      // Mock responses for all products (since no products detected, it tries all)
      const mockEmptyResponse = { value: [] };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValue(JSON.stringify(mockEmptyResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'nonexistent query xyz123',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      expect(response.initialSearchResults.totalHits).toBe(0);
      expect(response.initialSearchResults.productsDetected).toHaveLength(0);
      expect(Array.isArray(response.productResults)).toBe(true);
    });

    it('should handle queries when initial search fails but product searches succeed', async () => {
      // Mock initial search failure
      const mockEmptyResponse = { value: [] };

      const mockOutlookResponse = {
        value: [
          {
            id: 'msg-1',
            subject: 'Test Email',
            from: { emailAddress: { address: 'test@example.com' } },
            receivedDateTime: '2024-01-01T10:00:00Z',
            bodyPreview: 'Test email content',
            webLink: 'https://outlook.office.com/mail/id/msg-1',
            hasAttachments: false,
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockRejectedValueOnce(new Error('Initial search failed'))
        .mockResolvedValue(JSON.stringify(mockOutlookResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'test email',
        maxResults: 50,
        topPerProduct: 5,
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      // Should still have product results even if initial search failed
      expect(response.thinking).toBeDefined();
      expect(Array.isArray(response.thinking)).toBe(true);
      // Should contain error message about initial search
      expect(response.thinking.some((t: string) => t.includes('Initial search failed'))).toBe(true);
    });

    it('should respect maxResults parameter', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 100,
                hits: Array.from({ length: 30 }, (_, i) => ({
                  resource: {
                    '@odata.type': '#microsoft.graph.message',
                    id: `msg-${i}`,
                    subject: `Email ${i}`,
                  },
                  summary: `Email ${i}`,
                  rank: i + 1,
                })),
              },
            ],
          },
        ],
      };

      const mockOutlookResponse = {
        value: Array.from({ length: 5 }, (_, i) => ({
          id: `msg-${i}`,
          subject: `Email ${i}`,
          from: { emailAddress: { address: 'test@example.com' } },
          receivedDateTime: '2024-01-01T10:00:00Z',
          bodyPreview: 'Test email',
          webLink: `https://outlook.office.com/mail/id/msg-${i}`,
          hasAttachments: false,
        })),
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockOutlookResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'test',
        maxResults: 20, // Custom maxResults
        topPerProduct: 3, // Custom topPerProduct
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      // Verify that the request was made with correct parameters
      const calls = (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>).mock.calls;
      const searchCall = calls.find((call: unknown[]) => {
        const endpoint = call[0] as string;
        return endpoint === '/search/query';
      });

      expect(searchCall).toBeDefined();
      if (searchCall && searchCall[1]?.body) {
        const body = JSON.parse(searchCall[1].body);
        expect(body.requests[0].size).toBeLessThanOrEqual(20);
      }
    });

    it('should respect topPerProduct parameter', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 5,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.message',
                      id: 'msg-1',
                      subject: 'Test Email',
                    },
                    summary: 'Test Email',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockOutlookResponse = {
        value: Array.from({ length: 10 }, (_, i) => ({
          id: `msg-${i}`,
          subject: `Email ${i}`,
          from: { emailAddress: { address: 'test@example.com' } },
          receivedDateTime: '2024-01-01T10:00:00Z',
          bodyPreview: 'Test email',
          webLink: `https://outlook.office.com/mail/id/msg-${i}`,
          hasAttachments: false,
        })),
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockOutlookResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'test',
        maxResults: 50,
        topPerProduct: 3, // Limit to 3 results per product
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      const outlookResult = response.productResults.find(
        (r: { product: string }) => r.product === 'Outlook'
      );
      if (outlookResult && outlookResult.topResults) {
        // Should be limited to topPerProduct (3)
        expect(outlookResult.topResults.length).toBeLessThanOrEqual(3);
      }
    });
  });

  describe('Response Structure Validation', () => {
    it('should return valid response structure with all required fields', async () => {
      const mockInitialSearchResponse = {
        value: [
          {
            hitsContainers: [
              {
                total: 1,
                hits: [
                  {
                    resource: {
                      '@odata.type': '#microsoft.graph.message',
                      id: 'msg-1',
                      subject: 'Test',
                    },
                    summary: 'Test',
                    rank: 1,
                  },
                ],
              },
            ],
          },
        ],
      };

      const mockOutlookResponse = {
        value: [
          {
            id: 'msg-1',
            subject: 'Test',
            from: { emailAddress: { address: 'test@example.com' } },
            receivedDateTime: '2024-01-01T10:00:00Z',
            bodyPreview: 'Test content',
            webLink: 'https://outlook.office.com/mail/id/msg-1',
            hasAttachments: false,
          },
        ],
      };

      (mockGraphClient.makeRequest as ReturnType<typeof vi.fn>)
        .mockResolvedValueOnce(JSON.stringify(mockInitialSearchResponse))
        .mockResolvedValueOnce(JSON.stringify(mockOutlookResponse));

      const result = await productSearchHandler({
        action: 'search',
        query: 'test',
      });

      expect(result.isError).toBeFalsy();
      const responseText = result.content[0].text;
      const response = JSON.parse(responseText);

      // Validate top-level structure
      expect(response).toHaveProperty('query');
      expect(response).toHaveProperty('initialSearchResults');
      expect(response).toHaveProperty('productResults');
      expect(response).toHaveProperty('thinking');

      // Validate initialSearchResults structure
      expect(response.initialSearchResults).toHaveProperty('totalHits');
      expect(response.initialSearchResults).toHaveProperty('productsDetected');
      expect(typeof response.initialSearchResults.totalHits).toBe('number');
      expect(Array.isArray(response.initialSearchResults.productsDetected)).toBe(true);

      // Validate productResults structure
      expect(Array.isArray(response.productResults)).toBe(true);
      if (response.productResults.length > 0) {
        const productResult = response.productResults[0];
        expect(productResult).toHaveProperty('product');
        expect(productResult).toHaveProperty('resultCount');
        expect(productResult).toHaveProperty('topResults');
        expect(typeof productResult.product).toBe('string');
        expect(typeof productResult.resultCount).toBe('number');
        expect(Array.isArray(productResult.topResults)).toBe(true);

        // Validate topResults structure
        if (productResult.topResults.length > 0) {
          const topResult = productResult.topResults[0];
          expect(topResult).toHaveProperty('title');
          expect(topResult).toHaveProperty('summary');
          expect(topResult).toHaveProperty('relevance');
          expect(topResult).toHaveProperty('metadata');
          expect(typeof topResult.title).toBe('string');
          expect(typeof topResult.summary).toBe('string');
          expect(typeof topResult.relevance).toBe('number');
          expect(typeof topResult.metadata).toBe('object');
        }
      }

      // Validate thinking array
      expect(Array.isArray(response.thinking)).toBe(true);
      response.thinking.forEach((thought: unknown) => {
        expect(typeof thought).toBe('string');
      });
    });
  });
});
