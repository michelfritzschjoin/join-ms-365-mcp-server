#!/usr/bin/env node
/**
 * Script to find and read the roadmap email from January 30, 2026
 * Uses the same logic as the Super-Tools email search and get functionality
 */

import { config } from 'dotenv';
import { resolve } from 'path';
import { existsSync } from 'fs';

// Load environment variables
const envPath = resolve(process.cwd(), '.env');
if (existsSync(envPath)) {
  config({ path: envPath });
} else {
  config();
}

import logger from '../src/logger.js';
import AuthManager from '../src/auth.js';
import GraphClient from '../src/graph-client.js';
import { getSecrets } from '../src/secrets.js';
import { getQueryOptimizer, type OptimizationContext } from '../src/query-optimizer.js';
import { getQueryStore } from '../src/query-store.js';

/**
 * Format search query for Microsoft Graph API (simplified version)
 */
function formatSearchQuery(searchValue: string | undefined | null): string {
  if (!searchValue) return '';
  if (typeof searchValue !== 'string') {
    const stringValue = String(searchValue);
    if (stringValue === 'null' || stringValue === 'undefined') return '';
    searchValue = stringValue;
  }

  const trimmedValue = searchValue.trim();
  // Wrap in quotes for Microsoft Graph $search
  if (trimmedValue.includes(' ')) {
    return `"${trimmedValue}"`;
  }
  return trimmedValue;
}

/**
 * Search for emails containing "Roadmap" from January 30, 2026
 */
async function findRoadmapEmail(): Promise<void> {
  try {
    logger.info('Initializing authentication and Graph client...');

    // Initialize secrets
    const secrets = await getSecrets();

    // Initialize auth manager
    const authManager = new AuthManager(secrets);

    // Initialize Graph client
    const graphClient = new GraphClient(authManager, secrets);

    // Check if user is logged in
    const accounts = await authManager.listAccounts();
    if (accounts.length === 0) {
      logger.error('No accounts found. Please login first using the login tool.');
      process.exit(1);
    }

    logger.info(`Found ${accounts.length} account(s). Using first account.`);

    // Set user context for query optimization
    const userId = accounts[0].userId;
    const userIdHash = userId ? getQueryStore().hashUserId(userId) : undefined;

    // Search for roadmap email from January 30, 2026
    const searchQuery = 'Roadmap';
    const targetDate = '2026-01-30';

    logger.info(`Searching for emails containing "${searchQuery}" from ${targetDate}...`);

    // Optimize query (same as Super-Tools do)
    const queryOptimizer = getQueryOptimizer();
    const optimizationContext: OptimizationContext = {
      tool: 'email',
      entityTypes: ['message'],
      userIdHash,
    };
    const optimized = queryOptimizer.optimizeQuery(searchQuery, optimizationContext);

    logger.info(`Optimized query: "${optimized.optimizedQuery}"`);

    // Build search parameters
    const params: Record<string, string> = {
      $search: formatSearchQuery(optimized.optimizedQuery),
      $top: '50', // Get more results to find the right one
      $select: 'id,subject,bodyPreview,receivedDateTime,from,toRecipients,hasAttachments,webLink',
    };

    // Filter by date: emails from January 30, 2026
    const startDate = new Date('2026-01-30T00:00:00Z');
    const endDate = new Date('2026-01-30T23:59:59Z');
    params.$filter = `receivedDateTime ge ${startDate.toISOString()} and receivedDateTime le ${endDate.toISOString()}`;

    logger.info(`Date filter: ${params.$filter}`);

    // Search emails
    const result = await graphClient.makeRequest('/me/messages', {
      method: 'GET',
      queryParams: params,
    });

    const parsedResult = typeof result === 'string' ? JSON.parse(result) : result;

    if (!parsedResult.value || parsedResult.value.length === 0) {
      logger.warn('No emails found matching the search criteria.');
      logger.info('Trying broader search without date filter...');

      // Try without date filter
      const paramsNoDate: Record<string, string> = {
        $search: formatSearchQuery(optimized.optimizedQuery),
        $top: '50',
        $select: 'id,subject,bodyPreview,receivedDateTime,from,toRecipients,hasAttachments,webLink',
        $orderby: 'receivedDateTime desc',
      };

      const resultNoDate = await graphClient.makeRequest('/me/messages', {
        method: 'GET',
        queryParams: paramsNoDate,
      });

      const parsedResultNoDate =
        typeof resultNoDate === 'string' ? JSON.parse(resultNoDate) : resultNoDate;

      if (!parsedResultNoDate.value || parsedResultNoDate.value.length === 0) {
        logger.error('No emails found with "Roadmap" in the search.');
        process.exit(1);
      }

      // Filter results to find emails from January 30, 2026
      const emails = parsedResultNoDate.value.filter((email: { receivedDateTime: string }) => {
        const emailDate = new Date(email.receivedDateTime);
        return emailDate >= startDate && emailDate <= endDate;
      });

      if (emails.length === 0) {
        logger.error(`No emails found from ${targetDate}.`);
        logger.info('Found emails with "Roadmap":');
        parsedResultNoDate.value
          .slice(0, 10)
          .forEach((email: { subject: string; receivedDateTime: string }) => {
            logger.info(
              `  - ${email.subject} (${new Date(email.receivedDateTime).toLocaleDateString()})`
            );
          });
        process.exit(1);
      }

      parsedResult.value = emails;
    }

    const emails = parsedResult.value;
    logger.info(`Found ${emails.length} email(s) matching the criteria.`);

    // Find the most relevant email (preferably with "Roadmap" in subject)
    let roadmapEmail = emails.find((email: { subject: string }) =>
      email.subject.toLowerCase().includes('roadmap')
    );

    if (!roadmapEmail && emails.length > 0) {
      roadmapEmail = emails[0];
      logger.info('No email with "Roadmap" in subject, using first result.');
    }

    if (!roadmapEmail) {
      logger.error('Could not find a roadmap email.');
      process.exit(1);
    }

    logger.info(`\n📧 Found email: "${roadmapEmail.subject}"`);
    logger.info(
      `   From: ${roadmapEmail.from?.emailAddress?.name || 'Unknown'} <${roadmapEmail.from?.emailAddress?.address || 'Unknown'}>`
    );
    logger.info(`   Date: ${new Date(roadmapEmail.receivedDateTime).toLocaleString()}`);
    logger.info(`   ID: ${roadmapEmail.id}\n`);

    // Get full email content
    logger.info('Reading full email content...');
    const fullEmail = await graphClient.makeRequest(`/me/messages/${roadmapEmail.id}`, {
      method: 'GET',
      queryParams: {
        $select:
          'id,subject,body,bodyPreview,receivedDateTime,from,toRecipients,ccRecipients,hasAttachments,webLink',
      },
    });

    const emailContent = typeof fullEmail === 'string' ? JSON.parse(fullEmail) : fullEmail;

    // Display email content
    console.log('\n' + '='.repeat(80));
    console.log('📋 ROADMAP EMAIL CONTENT');
    console.log('='.repeat(80));
    console.log(`\nSubject: ${emailContent.subject}`);
    console.log(
      `From: ${emailContent.from?.emailAddress?.name || 'Unknown'} <${emailContent.from?.emailAddress?.address || 'Unknown'}>`
    );
    console.log(`Date: ${new Date(emailContent.receivedDateTime).toLocaleString()}`);
    console.log(`\n${'─'.repeat(80)}\n`);

    // Extract and display body content
    if (emailContent.body?.content) {
      // Remove HTML tags if present
      let bodyText = emailContent.body.content;
      if (emailContent.body.contentType === 'html') {
        // Simple HTML tag removal (for basic cases)
        bodyText = bodyText
          .replace(/<[^>]*>/g, '')
          .replace(/&nbsp;/g, ' ')
          .replace(/&amp;/g, '&')
          .replace(/&lt;/g, '<')
          .replace(/&gt;/g, '>')
          .replace(/&quot;/g, '"')
          .replace(/&#39;/g, "'")
          .trim();
      }
      console.log(bodyText);
    } else if (emailContent.bodyPreview) {
      console.log(emailContent.bodyPreview);
    } else {
      console.log('(No body content available)');
    }

    console.log('\n' + '='.repeat(80));

    // Extract roadmap information
    const bodyText = emailContent.body?.content || emailContent.bodyPreview || '';
    const roadmapMatch = bodyText.match(
      /roadmap|meilenstein|timeline|zeitplan|verantwortlich|responsible/gi
    );

    if (roadmapMatch) {
      console.log('\n✅ Roadmap keywords found in email content.');
    }

    logger.info('\n✅ Email content retrieved successfully.');
  } catch (error) {
    logger.error('Error finding roadmap email:', error);
    process.exit(1);
  }
}

// Run the script
findRoadmapEmail().catch((error) => {
  logger.error('Fatal error:', error);
  process.exit(1);
});
