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

// Fix: get-chat endpoint — codegen can produce unterminated template literal and merged send_mail body/response
const brokenGetChat =
  /\{\s*method: 'get',\s*path: '\/chats\/:chatId',\s*alias: 'get-chat',\s*description: `[^`]*\\\\`, requestFormat: 'json', parameters: \[\s*\{\s*name: '\$select',[\s\S]*?\},\s*\{\s*name: '\$expand',[\s\S]*?\},\s*\], response: z\.`,\s*type: 'Body',\s*schema: send_mail_Body,\s*\},\s*\],\s*response: z\.void\(\)\s*\},\s*\]\s*\)\s*;/;
const fixedGetChat = `{
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

if (brokenGetChat.test(content)) {
  content = content.replace(brokenGetChat, fixedGetChat);
  console.log('patch-generated-client: patched get-chat endpoint (unterminated template literal).');
}

if (content !== original) {
  fs.writeFileSync(clientPath, content);
  console.log('patch-generated-client: wrote patched client.ts');
} else {
  console.log('patch-generated-client: no patches applied.');
}
