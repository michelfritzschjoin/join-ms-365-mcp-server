# Changelog

> **Repository:** https://github.com/michelfritzschjoin/join-ms-365-mcp-server

All notable changes to the Join Microsoft 365 MCP Server are documented here.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [1.7.0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.6.2...v1.7.0) (2026-01-27)

### Features

* improve calendar/email display with UTC times and quick summary lists ([8dbb573](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/8dbb57333e0213959b6010d9e42ec274fcce390d))

## [1.6.2](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.6.1...v1.6.2) (2026-01-27)

### Bug Fixes

* replace callEndpoint with callGraph wrapper for GraphClient compatibility ([e113d38](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/e113d38eb778f1a07f194564209e7cfea5ac01da))

## [1.6.1](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.6.0...v1.6.1) (2026-01-27)

### Bug Fixes

* add missing addThinkingToResponse export to thinking-process.ts ([52884fc](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/52884fcf9456b7d1136fa39a2911196511190566))

## [1.6.0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.5.1...v1.6.0) (2026-01-27)

### Features

* add search super-tool and read-only mode support ([4a3216d](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/4a3216d6e0d60056f6b7a741868e7e66e9db2739))
* add Super-Tools mode - consolidates 126+ tools into 10 unified tools ([d230300](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/d23030088970a6db6586ca0838ec07f0dd969dce))

## [1.5.1](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.5.0...v1.5.1) (2026-01-27)

### Bug Fixes

- robust schema recreation to prevent \_zod undefined crash in tools/list ([486129c](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/486129cb53b5a7fab3717442aaf5c433fa196c15))

### Documentation

- add version management rule and update readme with version info ([2026bdb](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/2026bdb6a93e499723a03bd99c8d07f5fd6b0873))

## [1.5.0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.4.0...v1.5.0) (2026-01-27)

### Features

- add thinking process for transparent reasoning in OpenWebUI ([4a334c3](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/4a334c306dc801b6d2e87435171c0537d2c3900b))

### Bug Fixes

- improve OAuth flow and fix tools/list schema crash ([588fc89](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/588fc89a4798eb125d30681769ececa64a7e300c))
- robust schema validation to prevent \_zod crash in tools/list ([3d7eb8b](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/3d7eb8b973b09b0c3013fd55a0c56df83145be3b))
- update tests to support registertool wrapper ([31bbe74](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/31bbe74c35278b437b75e2666787dca590c12ba7))

## [1.4.0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.3.0...v1.4.0) (2026-01-27)

### Features

- add query logging for mcp tool calls to dashboard ([f889935](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/f8899358d7e36775cfdcbc535e15cf8509757c1b))

## [1.3.0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.2.4...v1.3.0) (2026-01-27)

### Features

- return explicit error when dashboard is accessed without password ([346c270](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/346c270f7f81f1285cf076df75a64843cb74f2d4))

### Bug Fixes

- increase oauth code max length to 4000 chars ([1b582d0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/1b582d0e313d23223156a5183e148e45f7bf304b))

## [1.2.4](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.2.3...v1.2.4) (2026-01-27)

### Bug Fixes

- relax oauth callback parameter validation ([bd5fe85](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/bd5fe85e85a5523dd9928a26d9f7e493b904c32b))

## [1.2.3](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.2.2...v1.2.3) (2026-01-27)

### Bug Fixes

- replace onlinemeetings.read.all with onlinemeetings.read for delegated permissions ([d7ae4b1](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/d7ae4b171f068914cda0db2bb4d45ab886172d49))

## [1.2.2](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.2.1...v1.2.2) (2026-01-27)

### Bug Fixes

- allow unsafe-inline scripts for dashboard routes in csp ([98c063e](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/98c063eeec16f5db38835570f8c4884e56410139))

## [1.2.1](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.2.0...v1.2.1) (2026-01-27)

### Bug Fixes

- improve dashboard authentication for https with secure cookies and redirect ([40d22a9](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/40d22a9a920904e6d079384da511b2ef41a0ca8e))

## [1.2.0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.1.0...v1.2.0) (2026-01-27)

### Features

- add query logging and analytics dashboard ([97b84ef](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/97b84ef400cae9dc1a176c775239e8ea66b3b031))

### Bug Fixes

- improve dashboard authentication for https with secure cookies and redirect ([7a2c8c0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/7a2c8c0))

## [1.1.0](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/compare/v1.0.0...v1.1.0) (2026-01-26)

### Features

- add structured response formatting with server local time for calendar and mail ([cfd203f](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/cfd203fee0d881ab672a0690ca25194f6b039253))

### Bug Fixes

- set timezone to europe/berlin in docker for correct date handling ([bd69a4f](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/bd69a4fd7f290898e8bba63009866d19f199083f))

## 1.0.0 (2026-01-26)

### Features

- add ConsistencyLevel header for SharePoint sites search ([216a334](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/216a33400a7cb772aec1e7414d63a02476cce619))
- Add create-draft-email endpoint ([#60](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/60)) ([cdd6500](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/cdd65007e3aca0005086d930fa81757fea849ccb))
- add default date/time for calendar queries - defaults to now until tomorrow end of day with detailed content ([6ff4948](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/6ff4948ded6c2215a13ce127b346bb2d55fd7c10))
- Add destructiveHint and openWorldHint annotations to all tools ([#180](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/180)) ([09ff7b9](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/09ff7b91024abde05aaa540fee78afc384a379a2))
- add Docker Compose with Traefik support and Join AI solutions promotion ([8f6057b](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/8f6057b26aa8808696974e56506f8ecb89f5bee7))
- add endpoints for managing email attachments ([#112](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/112)) ([31bb63e](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/31bb63e54cb0273696ee03548ac1cacfae11fd35))
- add enhanced get-my-emails tool with rich formatting ([b93a5ab](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b93a5ab0630fc075708078f829d95087cc730769))
- add excludeResponse parameter to reduce token usage ([#136](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/136)) ([a842317](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/a8423178aa6321ce17b2c7850f223d29ffe1e66b))
- add experimental TOON output format for token-efficient responses ([#149](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/149)) ([6926c3a](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/6926c3a5d083e69f1471881747db9e7bfc0568bc))
- add extended timeframe search and permission awareness ([09299eb](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/09299eb77e4363df9966247b278e25156095dc61))
- add findMeetingTimes endpoint with Calendars.Read.Shared ORG-MODE scope ([#143](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/143)) ([0f5b0c6](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/0f5b0c61cc34e79540c521b8a955e2251f6dc1b9))
- add follow-up search with corrected names in search-everything ([4598a98](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/4598a9807fafa20333a320210d0982a89fe92560))
- add GET method support for MCP Streamable HTTP endpoint ([#108](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/108)) ([1b5210c](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/1b5210c9805616e136de8b7ffcc76a3fbdb379da))
- add get-teams-channel-posts compound tool for retrieving Teams channel messages ([9b36bdd](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/9b36bdda1830933034d49ed32171847fda5d1726))
- Add includeHeaders parameter to capture ETag values for Microsoft Graph operations ([#121](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/121)) ([1e29081](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/1e29081687eec28422a209dffe9b46cfa1ee17ff))
- add intelligent bilingual tools for Microsoft 365 queries ([b021768](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b02176814b9e05b93322e0846b2cc74d4008c66a))
- add meeting transcript tools and endpoints ([1aef477](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/1aef4774c73fe7b21de509299874ef9dfc333e1c))
- Add Microsoft 365 China (21Vianet) cloud support ([#184](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/184)) ([6932928](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/6932928073e10a27854b3c851f3981588e71887e))
- add option to select on what interface to bind http ([#170](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/170)) ([b7c14b1](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b7c14b12f0ea5068c96bf7a8ccad9989bcce25a7)), closes [#169](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/169)
- add per-chat memory for OpenWebUI sessions ([fadbf3e](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/fadbf3e35e087bec91615a12a9f78b05a63bd92c))
- add search functionality ([#105](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/105)) ([7213afb](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/7213afb26595f5d69c763d0efd6b468610594074))
- add shared mailbox access endpoints - also had to add /users for this to work ([#111](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/111)) ([412c3fd](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/412c3fddfb3f5b1760a0007798626ae2b7bd02db))
- **compound-tools:** add 8 business-focused intelligence tools ([491d590](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/491d590b25d603d0ca7b17fba9e871a011fe268c))
- enhance ask-microsoft-365 tool with 16 products and search-first strategy ([ccb6394](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/ccb63946280f2df6cd20b8fa4818a2491f1fc7e4))
- enhance get-my-emails with rich content summaries and statistics ([586a71a](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/586a71a72d84d5f60a062f512bc77d91b497cd52))
- enhance OpenAPI schema processing by pruning unused schemas and simplifying anyOf patterns! ([#92](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/92)) ([54c4371](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/54c437150fefcd64e4ca6ddc01ea636c443b5eac))
- enhance schema processing for requests and allOf items in OpenAPI - no more body as json! ([#109](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/109)) ([f0924bb](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/f0924bb83138ec3a7f86e4a83652c7aba147f34f))
- establish search-everything as universal fallback tool ([b097501](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b0975017148d72284f965b1e97cad0dedaf4d928))
- fix release build (hopefully) ([#150](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/150)) ([62ccff1](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/62ccff1ea4f21555b4c5288ae6c6996d61d29a05))
- Hello discovery tools and categories - bye bye error arrays ([#167](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/167)) ([bf13fab](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/bf13fab1afd6c930ffc3a58ae99c0416286b445d))
- hello hack; allow polymorphic attachment properties by updating object schema to passthrough for microsoft_graph_attachment ([#151](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/151)) ([75691bf](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/75691bffd1f7aeec2446413db01cd8cfd30c25a0))
- improve ask-microsoft-365 people search results ([648d316](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/648d3169e229deae3caf774e4ca1a91d51d73dd6))
- improve environment variable handling and add masked display on startup ([7ad166d](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/7ad166d9e15b0943cec43a2f665dcb356f5bd695))
- make calendar tools timezone aware ([#164](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/164)) ([f46b5ef](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/f46b5eff7a9418db2d2d74fba86ccbc2377e667d))
- **oauth:** implement RFC 7591 Dynamic Client Registration ([7fbb325](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/7fbb325ec9f52d759ebcb5a42a96cf35ba1f98b4))
- org-mode (work-mode) force modes! ([#97](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/97)) ([b03f08b](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b03f08b6d5a267f051aa921712a436cf640dd501))
- remove default date filters and add detailed content for all resource types ([db17c9d](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/db17c9d157d592993c3d39244bb2dbed0b5046c8))
- remove unused Excel session management code - legacy stuff ([#107](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/107)) ([715b58a](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/715b58ac75082ce44973a7f1b7c75765c6a5f58b))
- rewrite search-everything to use Microsoft Search API ([d0d428b](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/d0d428be419063dd5b13832ba09bc502fc356990))
- The “llmTip” field has been added to the endpoints.json to provide hints to the LLM. ([#163](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/163)) ([e4a3fae](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/e4a3fae0ef21af9acd1c1f692013b8a07a01d55b))

### Bug Fixes

- add date/time context and llm instructions for email/calendar responses ([fc0227d](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/fc0227d2dedc2a08724679210a2b15d9863ffd58))
- add descriptions to tools for improved clarity and usability ([#95](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/95)) ([27e81cb](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/27e81cb63b490903a4ab8816f905563ddb73790b))
- add linting and formatting checks to build process ([#98](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/98)) ([13fada3](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/13fada38524428d8f977b3da5a0accaca7be53e9))
- add new endpoints for updating planner tasks and their details ([#93](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/93)) ([7180369](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/71803698da6121b5a98c14b95ab0e4fbe1287456))
- add pagination support for Teams channel search ([241a630](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/241a63009ef2a02cc81540aa4d74765134cddfab))
- add queryParams support to GraphClient.performRequest ([82f8719](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/82f871931f6077050eb101d95992602425406023))
- Add quotes to file paths in execSync command ([#102](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/102)) ([8828229](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/88282294ede948c3ca47e4a479a28e0b67287c18))
- add support for auto-returning download URLs for specific endpoints helping agent not sending whole file to LLM ([#142](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/142)) ([5e37ad3](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/5e37ad38ea88930c4343842c9683b216fefb6c2e))
- auto-format search queries for list-users tool to use property:value format ([3d62cf8](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/3d62cf84ce67781fd0db8320ab6a20d230bc88c2))
- correct SharePoint file download endpoint path ([#139](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/139)) ([3fc0222](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/3fc0222d85986eb7943e2c5310351da736dc0a0e)), closes [#137](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/137)
- correct tsx argument parsing in dev:http script ([1dd587f](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/1dd587ff9ad017ca39d834653957cbd27ce80a3d))
- ensure OAuth tokens don't leak over to other users requests using AsyncLocalStorage ([#188](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/188)) ([b1bdc52](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b1bdc520ab034f707b283285ce113b1ef59e3e8a))
- ensure retried requests go through full error handling and parsing ([#135](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/135)) ([ca6cfbc](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/ca6cfbc2944ee99f21bd849346059fdc257fb8de))
- Fix calendarId parameter handling for calendar operations ([#116](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/116)) ([1c2cdae](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/1c2cdae3a846363ac137520b559561de8c331817))
- handle empty responses from Microsoft Graph API ([#94](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/94)) ([b63de33](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b63de332fc7b55d67e7911a8ebfa9adf5f0766cc))
- handle token rejected promise ([#99](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/99)) ([e22e59c](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/e22e59c4826dfad3e011fc1c6021cdac495bd936))
- improve Docker build resilience for ARM64 emulated environments ([42c1004](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/42c10042acdb1af83f984d07639ce532fd9738ae))
- improve people search results detection and merging ([5d62469](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/5d624694dc9c3f72347d9b4968583a5d95c8425b))
- lint! ([#100](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/100)) ([59e6fcf](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/59e6fcf08eaff2b4817bcd7a1119e2f41f49b45a))
- make chmod command cross-platform compatible ([4a435b3](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/4a435b3d3e77c3bb1d03ea02f8556c70d2f5761e))
- make client secret optional for public apps in Microsoft auth flow ([#172](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/172)) ([7f2a6e4](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/7f2a6e42ec532b20edfb94e99379790a82b20131))
- OAuth scope filtering to respect enabled tools filter ([#153](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/153)) ([09e4eb3](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/09e4eb3288fc2086954d83a618a19112e3f1a9e0))
- **oauth:** avoid auth router intercepting /token ([a64172b](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/a64172b081d85c3ed46d5c16564e68560e97d24b))
- **oauth:** improve token exchange error handling and logging ([c210838](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/c210838c4d5d6f183430a762efeb8b4706ab40e5))
- **oauth:** preserve state parameter for CSRF validation ([7a47276](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/7a47276cbd311f9bf0b77a68eb1607fc0d4e8318))
- **oauth:** resolve token exchange and rate limiting issues ([061eb85](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/061eb85a18839e15abfc957e6b4bd212e78d8817))
- prep for MCP Docker registry thing - adjust the docker file ([#159](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/159)) ([2ce965c](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/2ce965c4917ba1e71f3c44485987dd94177a401d))
- prep for MCP Docker registry thing - adjust the docker file ONCE MORE ([#160](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/160)) ([ab223e7](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/ab223e725647ee46270b5f8f06d87b594dd3abee))
- prep for MCP Docker registry thing - adjust the docker file ONCE MORE ([#161](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/161)) ([cb9bf70](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/cb9bf706b421b6caa82aa9e62ab6f5a457ae63ea))
- prep for MCP Docker registry thing - adjust the docker file ONCE MORE ([#162](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/162)) ([0a8b512](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/0a8b51242cf2cc143a93d4b281b7ac9850f0274c))
- prep for MCP Docker registry thing ([#158](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/158)) ([10db7ed](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/10db7ed173fc712da490b224ff2cc809850e9c69))
- prevent infinite recursion in mergeAllOfSchemas by tracking visited schemas - hope this works ([#132](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/132)) ([7bc3853](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/7bc3853a082dd395349eac60cd114801084e2f0a))
- remove postinstall script as this should definitely not run for end-users! ([#114](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/114)) ([2e4e979](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/2e4e979720f10246ae22c7abfd814ae8915c261b))
- replace OData contains filter with client-side filtering for calendar ([84bcff4](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/84bcff433acaaf28c0b9833872e67189f82213a1))
- resolve ARM64 Docker build failures by generating client code in CI ([67a888c](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/67a888cedf2a00103c0e3b7fc3ccd43a21cbe8ca))
- resolve search api 400 error for events and fix test suite issues ([8935fae](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/8935faea2d1527d6f94af206cd2aa194f4b4bdfd))
- **security:** harden file permissions and reduce logging exposure ([#176](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/176)) ([b559ef8](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/b559ef8db08ea6bb892297bb7ec4f2706ebbc85d))
- **security:** remove unauthenticated /register endpoint ([#178](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/178)) ([f119b53](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/f119b5371db126b80adadb16c5c6e54e4a2a8e07))
- semantic release! ([#89](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/89)) ([36e584c](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/36e584c4dffd1d7013478b5b48fee4a8299f616f))
- update dependencies to address security vulnerabilities ([#175](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/175)) ([00d5a51](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/00d5a51467c133daca95d884376dd56b747d2ca3))
- update dev:http script to use tsx --watch and add automatic generation check ([36305c5](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/36305c505cd4ea9b5a693f95705970933be7c042))
- update package description for clarity and specificity ([#101](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/101)) ([0033654](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/0033654e819af168d8ab55a51aadb9314eaa9274))
- update the parameter handling by adding auto-correction for nested fields in request body ([#131](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/131)) ([6559751](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/6559751708b4b9768953322ce44ddace03c9895a))
- use /oauth/callback as MCP client callback path per MCP OAuth spec ([4a7b979](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/4a7b9794a7e5b55c6ab849c86dac1f4ab21cad7c))
- use tsx watch command instead of --watch flag ([21edc53](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/21edc53c1559a5a4c4fa4bcb327113c3feaccf12))

### Performance Improvements

- performance optimization for startup sequence ([1afd3f5](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/1afd3f576fcb5d66b9e1a23f248988ffe338252a))

### Documentation

- add comprehensive README with all tools and features documented ([5c31e47](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/5c31e4729ab40dd14665cb75b752850766e2291f))
- add MS365_MCP_ANONYMIZE_PII environment variable documentation ([0509859](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/0509859f4b8dd0f0e228cb71d03d4642e5245314))
- Added details on how to configure redirect URIs when using OAuth mode with custom Azure credentials. ([#155](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/issues/155)) ([60a640a](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/60a640af868d4fb6b3453702740a6f9cbb2e13b1))
- remove contact email from README ([a8d6509](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/a8d65095cb93252c479c24ea3cd14720d4dfc8a3))
- update README for Docker-only deployment, remove softeria references ([d1d3e27](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/d1d3e27256da9688ef56f423008865bcee3f43f1))

### Code Refactoring

- Add new features and update project structure ([629cc41](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/629cc41a145b228d4fa3c62e790707788596f53b))
- centralize Microsoft Search API usage with ranking and relevance ([4ec3e09](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/commit/4ec3e0947f93e0a68cf9f4c744ada81f15b8c0e3))

<!-- This changelog is automatically updated by semantic-release -->
<!-- Do not manually edit the sections below this line -->

## [0.0.0-development] - 2026-01-23

### Features

- Initial MCP server implementation
- Microsoft Graph API integration
- OAuth 2.1 with PKCE authentication
- Device code flow for login
- Support for multiple Microsoft 365 services:
  - Mail (read, send, manage)
  - Calendar (events, meetings)
  - Contacts
  - OneDrive (files, folders)
  - To-Do tasks
  - OneNote
  - Teams (chats, channels)
  - SharePoint (sites, lists)
  - Planner
  - Excel operations
- HTTP server mode with Streamable HTTP
- SSE transport for legacy clients
- stdio mode for local MCP clients
- Read-only mode
- Tool presets for filtering
- Azure Key Vault integration
- Multi-cloud support (Global, China)
- Comprehensive error handling
- Rate limiting
- Security headers
- CORS support

### Documentation

- Comprehensive documentation in `/docs`
- Security guide with ISO 27001 and DSGVO compliance
- API reference documentation
- Development guide

### Security

- Secure token storage using system keychain
- Input validation with Zod schemas
- ISO 27001 compliance measures
- DSGVO/GDPR compliance measures

---

_For release details, see [GitHub Releases](https://github.com/michelfritzschjoin/join-ms-365-mcp-server/releases)._
