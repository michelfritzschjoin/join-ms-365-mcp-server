import { z } from 'zod';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import AuthManager from './auth.js';

export function registerAuthTools(server: McpServer, authManager: AuthManager): void {
  server.registerTool(
    'login',
    {
      title: 'login',
      description:
        'REQUIRED: Authenticate with Microsoft 365 using device code flow. ' +
        'This tool MUST be called before using any other Microsoft 365 tools. ' +
        'It will provide a URL and code that the user must enter in a browser to authenticate.',
      inputSchema: z.object({
        force: z.boolean().default(false).describe('Force a new login even if already logged in'),
      }),
    },
    async ({ force }) => {
      try {
        if (!force) {
          const loginStatus = await authManager.testLogin();
          if (loginStatus.success) {
            return {
              content: [
                {
                  type: 'text',
                  text: JSON.stringify({
                    status: 'Already logged in',
                    ...loginStatus,
                  }),
                },
              ],
            };
          }
        }

        const text = await new Promise<string>((resolve, reject) => {
          authManager.acquireTokenByDeviceCode(resolve).catch(reject);
        });
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: 'device_code_required',
                message: text.trim(),
              }),
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: `Authentication failed: ${(error as Error).message}` }),
            },
          ],
        };
      }
    }
  );

  server.registerTool(
    'logout',
    {
      title: 'logout',
      description: 'Log out from Microsoft account',
      inputSchema: z.object({}),
    },
    async () => {
      try {
        await authManager.logout();
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ message: 'Logged out successfully' }),
            },
          ],
        };
      } catch {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ error: 'Logout failed' }),
            },
          ],
        };
      }
    }
  );

  server.registerTool(
    'verify-login',
    {
      title: 'verify-login',
      description:
        'Check if the user is currently authenticated with Microsoft 365. ' +
        'Use this to verify login status before attempting to use other tools.',
      inputSchema: z.object({}),
    },
    async () => {
      const testResult = await authManager.testLogin();

      return {
        content: [
          {
            type: 'text',
            text: JSON.stringify(testResult),
          },
        ],
      };
    }
  );

  server.registerTool(
    'list-accounts',
    {
      title: 'list-accounts',
      description: 'List all available Microsoft accounts',
      inputSchema: z.object({}),
    },
    async () => {
      try {
        const accounts = await authManager.listAccounts();
        const selectedAccountId = authManager.getSelectedAccountId();
        const result = accounts.map((account) => ({
          id: account.homeAccountId,
          username: account.username,
          name: account.name,
          selected: account.homeAccountId === selectedAccountId,
        }));

        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({ accounts: result }),
            },
          ],
        };
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Failed to list accounts: ${(error as Error).message}`,
              }),
            },
          ],
        };
      }
    }
  );

  server.registerTool(
    'select-account',
    {
      title: 'select-account',
      description: 'Select a specific Microsoft account to use',
      inputSchema: z.object({
        accountId: z.string().describe('The account ID to select'),
      }),
    },
    async ({ accountId }) => {
      try {
        const success = await authManager.selectAccount(accountId);
        if (success) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({ message: `Selected account: ${accountId}` }),
              },
            ],
          };
        } else {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({ error: `Account not found: ${accountId}` }),
              },
            ],
          };
        }
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Failed to select account: ${(error as Error).message}`,
              }),
            },
          ],
        };
      }
    }
  );

  server.registerTool(
    'remove-account',
    {
      title: 'remove-account',
      description: 'Remove a Microsoft account from the cache',
      inputSchema: z.object({
        accountId: z.string().describe('The account ID to remove'),
      }),
    },
    async ({ accountId }) => {
      try {
        const success = await authManager.removeAccount(accountId);
        if (success) {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({ message: `Removed account: ${accountId}` }),
              },
            ],
          };
        } else {
          return {
            content: [
              {
                type: 'text',
                text: JSON.stringify({ error: `Account not found: ${accountId}` }),
              },
            ],
          };
        }
      } catch (error) {
        return {
          content: [
            {
              type: 'text',
              text: JSON.stringify({
                error: `Failed to remove account: ${(error as Error).message}`,
              }),
            },
          ],
        };
      }
    }
  );
}
