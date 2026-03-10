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

// Fix: get-chat endpoint — codegen can produce unterminated template literal and merged send_mail body/response.
const fixedGetChatBlock = `{
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
]);`;

// Match 1: broken get-chat with send_mail_Body in the same block (merged output).
const getChatBrokenWithSendMail =
  /\{\s*method: 'get',\s*path: '\/chats\/:chatId'[\s\S]*?response: z\.`[\s\S]*?send_mail_Body[\s\S]*?\]\s*\)\s*;/;
// Match 2: broken get-chat with unterminated template literal, up to end of makeApi (]);).
const getChatBrokenAny =
  /\{\s*method: 'get',\s*path: '\/chats\/:chatId'[\s\S]*?response: z\.`[\s\S]*?\]\s*\)\s*;/;

if (getChatBrokenWithSendMail.test(content)) {
  content = content.replace(getChatBrokenWithSendMail, fixedGetChatBlock);
  console.log(
    'patch-generated-client: patched get-chat endpoint (unterminated template literal, with send_mail).'
  );
} else if (getChatBrokenAny.test(content)) {
  content = content.replace(getChatBrokenAny, fixedGetChatBlock);
  console.log('patch-generated-client: patched get-chat endpoint (unterminated template literal).');
} else if (content.includes("alias: 'get-chat'") && content.includes('response: z.`')) {
  // Fallback: broken get-chat with unterminated literal but regex didn't match (format variation)
  const makeApiStart = content.indexOf('const endpoints = makeApi([');
  if (makeApiStart !== -1) {
    const afterBracket = content.indexOf('\n  {', makeApiStart);
    const makeApiEnd = content.indexOf(']);', afterBracket);
    if (afterBracket !== -1 && makeApiEnd !== -1) {
      const before = content.slice(0, afterBracket);
      const after = content.slice(makeApiEnd);
      content =
        before +
        `
  {
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
` +
        after;
      console.log('patch-generated-client: patched get-chat endpoint (fallback).');
    }
  }
} else if (
  content.includes("path: '/chats/:chatId'") &&
  (content.includes('response: z.`') ||
    (content.includes('description: `Get chat (without its messages)') &&
      !content.includes('the request initiated from.`')))
) {
  // Fallback 2: broken get-chat (unterminated template or truncated file in CI)
  const getChatStart = content.indexOf("path: '/chats/:chatId'");
  if (getChatStart !== -1) {
    let blockStart = content.lastIndexOf('\n  {', getChatStart);
    if (blockStart < 0) blockStart = content.lastIndexOf('{\n    method:', getChatStart);
    if (blockStart < 0) blockStart = content.lastIndexOf('{', getChatStart);
    const suffix = `]);

export const api = new Zodios(endpoints);

export function createApiClient(baseUrl: string, options?: ZodiosOptions) {
  return new Zodios(baseUrl, endpoints, options);
}
`;
    const fixedBlockContent = `{
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
    const before = content.slice(0, blockStart >= 0 ? blockStart : getChatStart - 20);
    const closing = content.indexOf(']);', getChatStart);
    const after = closing !== -1 ? content.slice(closing) : suffix;
    content = before + (blockStart >= 0 ? '\n  ' : '') + fixedBlockContent.trimEnd() + '\n' + after;
    console.log(
      'patch-generated-client: patched get-chat endpoint (truncated/unterminated fallback).'
    );
  }
}

if (content !== original) {
  fs.writeFileSync(clientPath, content);
  console.log('patch-generated-client: wrote patched client.ts');
} else {
  console.log('patch-generated-client: no patches applied.');
}
