#!/usr/bin/env node
/**
 * Patches generated client.ts:
 * 1. response type "binary" -> "json" (fixes TS2322 when openapi-zod-client emits "binary").
 * 2. get-chat endpoint: fix unterminated template literal and wrong response/body merge from codegen.
 */

import fs from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const rootDir = path.resolve(__dirname, '..');
const clientPath = path.join(rootDir, 'src', 'generated', 'client.ts');

if (!fs.existsSync(clientPath)) {
  console.log('patch-generated-client: src/generated/client.ts not found, skipping.');
  process.exit(0);
}

let content = fs.readFileSync(clientPath, 'utf-8');
const original = content;

// Fix: Type '"binary"' is not assignable to type '"json"' (openapi-zod-client endpoint responseType)
content = content.replace(/: "binary"/g, ': "json"');
content = content.replace(/: 'binary'/g, ": 'json'");

// Fixed get-chat endpoint block (codegen can produce unterminated template literal or merged send_mail body).
const fixedGetChatBlockContent = `{
    method: 'get',
    path: '/chats/:chatId',
    alias: 'get-chat',
    description: \`Get chat (without its messages). This supports federation. To access a chat, at least one chat member must belong to the tenant the request initiated from.\`,
    requestFormat: 'json',
    parameters: [
      {
        name: '$select',
        type: 'Query',
        schema: z.array(z.string()).describe('Select properties to be returned').optional(),
      },
      {
        name: '$expand',
        type: 'Query',
        schema: z.array(z.string()).describe('Expand related entities').optional(),
      },
    ],
    response: z.lazy(() => microsoft_graph_chat),
  },
`;

// Always normalize get-chat block when present: find its boundaries and replace with fixed block.
// This fixes broken codegen (unterminated template literal, merged send_mail) regardless of format.
if (content.includes("path: '/chats/:chatId'")) {
  const pathIdx = content.indexOf("path: '/chats/:chatId'");
  const alreadyFixed = content.includes('response: z.lazy(() => microsoft_graph_chat)', pathIdx);
  if (!alreadyFixed) {
    // Find start of get-chat endpoint block (generator may use 2 or 4 spaces).
    let blockStart = content.lastIndexOf('\n  {', pathIdx);
    if (blockStart < 0) blockStart = content.lastIndexOf('\n    {', pathIdx);
    if (blockStart < 0) blockStart = content.lastIndexOf('{\n    method:', pathIdx);
    if (blockStart < 0) {
      const beforePath = content.slice(0, pathIdx);
      const re = /\n\s*\{\s*\n\s*method: 'get'/g;
      let m;
      let lastMatch;
      while ((m = re.exec(beforePath)) !== null) lastMatch = m;
      if (lastMatch) blockStart = lastMatch.index;
    }
    if (blockStart < 0) blockStart = content.lastIndexOf('{', pathIdx);
    // Never use pathIdx - 100: it can leave an unclosed template literal in the prefix.
    if (blockStart < 0) {
      console.warn(
        'patch-generated-client: could not find get-chat block start, skipping get-chat patch.'
      );
    } else {
      // End of get-chat block: next endpoint "\n  {" or "\n    {" or end of array "\n]);"
      const nextBlock2 = content.indexOf('\n  {', blockStart + 10);
      const nextBlock4 = content.indexOf('\n    {', blockStart + 10);
      const nextBlock =
        nextBlock2 !== -1 && (nextBlock4 === -1 || nextBlock2 <= nextBlock4)
          ? nextBlock2
          : nextBlock4;
      const endOfArray = content.indexOf('\n]);', blockStart);
      const blockEnd =
        nextBlock !== -1 && (endOfArray === -1 || nextBlock < endOfArray)
          ? nextBlock
          : endOfArray !== -1
            ? endOfArray
            : null;
      const prefix = content.slice(0, blockStart);
      const blockPrefix = content[blockStart] === '\n' ? '\n  ' : '';
      if (blockEnd !== null) {
        content =
          prefix +
          blockPrefix +
          fixedGetChatBlockContent.trimEnd() +
          '\n' +
          content.slice(blockEnd);
        console.log('patch-generated-client: patched get-chat endpoint (normalized).');
      } else {
        const suffix = `]);

export const api = new Zodios(endpoints);

export function createApiClient(baseUrl: string, options?: ZodiosOptions) {
  return new Zodios(baseUrl, endpoints, options);
}
`;
        content = prefix + blockPrefix + fixedGetChatBlockContent.trimEnd() + '\n' + suffix;
        console.log('patch-generated-client: patched get-chat endpoint (normalized, truncated).');
      }
    }
  }
}

if (content !== original) {
  fs.writeFileSync(clientPath, content);
  console.log('patch-generated-client: wrote patched client.ts');
} else {
  console.log('patch-generated-client: no patches applied.');
}
