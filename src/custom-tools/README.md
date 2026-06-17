# Custom Graph endpoints

This folder is the extension point for **hand-authored Microsoft Graph
endpoints** — Graph endpoints that the codegen pipeline (`bin/generate-graph-client.mjs`)
can't produce because the upstream Graph OpenAPI spec is missing or
under-specifies them.

Custom tools run through the **same** execution pipeline as the auto-generated
ones (`executeGraphTool` in [`../graph-tools.ts`](../graph-tools.ts)). They
inherit every existing feature for free:

- OData parameter normalization (`$filter`, `$top`, …)
- `$top` clamping via `MS365_MCP_MAX_TOP`
- Path-parameter URL encoding (with `skipEncoding` for function-style calls)
- Multi-account token routing (`account` parameter, `getTokenForAccount`)
- `fetchAllPages` pagination across `@odata.nextLink`
- Sensitivity filtering for email and document responses
- `includeHeaders` / `excludeResponse` controls
- Calendar `timezone` / `expandExtendedProperties` parameters
- Custom `Content-Type` and `Accept` headers
- `returnDownloadUrl` for `/content` endpoints

You write the metadata and the parameter schemas; the existing pipeline does
everything else.

## Layout

```
src/custom-tools/
├── custom-endpoints.json   ← manifest: one entry per custom tool
├── manifest.ts             ← sync loader (used by auth.ts during scope build)
├── registry.ts             ← async registration into the McpServer
├── types.ts                ← CustomEndpointConfig, CustomToolParameter, …
└── tools/                  ← parameter modules, one TS file per tool
    └── <tool-name>.ts
```

## Adding a new tool — three steps

### 1. Add a manifest entry to `custom-endpoints.json`

Every JSON-side field has identical semantics to its `endpoints.json`
counterpart (see [`../graph-tools.ts`](../graph-tools.ts) `EndpointConfig`):

```jsonc
{
  "tools": [
    {
      "toolName": "list-mailbox-folders-extended",
      "pathPattern": "/me/mailFolders",
      "method": "GET",
      "scopes": ["Mail.Read"],
      "workScopes": [],
      "llmTip": "Use this when list-mail-folders is missing the includeHidden filter.",

      // The two extra fields, required only for custom tools:
      "description": "List the user's mail folders, including hidden ones.",
      "parametersModule": "./tools/list-mailbox-folders-extended.js"
    }
  ]
}
```

**Always end `parametersModule` in `.js`** — this matches the project's
NodeNext module resolution and works in both `npm run dev` (tsx) and the
built `dist/` output.

Optional fields you can set the same way you would in `endpoints.json`:
`disabled`, `readOnly`, `supportsTimezone`, `supportsExpandExtendedProperties`,
`skipEncoding`, `contentType`, `acceptType`, `returnDownloadUrl`.

### 2. Create the parameters module under `./tools/`

```ts
// src/custom-tools/tools/list-mailbox-folders-extended.ts
import { z } from 'zod';
import type { CustomToolParameters } from '../types.js';

export const parameters: CustomToolParameters = [
  {
    name: 'includeHidden',
    type: 'Query',
    description: 'When true, include hidden mail folders.',
    schema: z.boolean().optional().describe('Include hidden folders'),
  },
  {
    name: '$select',
    type: 'Query',
    description: 'Comma-separated fields to return',
    schema: z.string().optional(),
  },
];
```

The `type` field on each parameter (`'Path' | 'Query' | 'Body' | 'Header'`)
drives the dispatch in `executeGraphTool` — Path values substitute into the
URL template, Query values land in the query string, Body becomes the JSON
body, Header becomes a request header.

You can omit Path parameters that appear in `pathPattern` — they're picked up
by `executeGraphTool`'s fallback. But declaring them gives the LLM a typed
schema, which is preferred.

You don't need to declare `account`, `fetchAllPages`, `includeHeaders`,
`excludeResponse`, `timezone`, or `expandExtendedProperties` — those are
auto-injected by the same code path that handles them for generated tools.

### 3. Build and run

```bash
npm run build      # compiles your TS module + copies the manifest into dist/
npm run dev        # see the registration log: "Custom tool registration complete: 1 registered, …"
```

## Path patterns

Use OpenAPI-style `{name}` placeholders in `pathPattern`:

```
/me/messages/{message-id}/attachments
```

The registry converts these to `:name` style internally before passing to
`executeGraphTool`, matching what the generated client uses.

## Scopes

Add the OAuth scopes your endpoint needs to `scopes` (personal account) and/or
`workScopes` (organization account, requires `--org-mode`). They are merged
into the login scope set automatically — when a user runs `--login`, your
custom-tool scopes are requested alongside the built-in ones.

The scope-hierarchy collapse (e.g. `Mail.ReadWrite` subsumes `Mail.Read`) in
[`../auth.ts`](../auth.ts) applies to the merged set, so you don't need to
worry about declaring a redundant lower scope.

## Read-only mode

Your tool participates in `--read-only` filtering using the same rules as
`endpoints.json`:

- `GET` is always registered.
- `POST/PATCH/DELETE` are skipped, **unless** the manifest entry has
  `"readOnly": true` (used for POST endpoints that perform read-only
  operations like `/me/calendar/getSchedule`).

## Disabling a tool

Set `"disabled": true` in the manifest entry. The tool isn't registered, and
its scopes aren't added to the login set. Same semantics as `endpoints.json`.

## Filtering with `--enabled-tools`

The `--enabled-tools` regex filter applies to custom tools too. Match against
`toolName` to include or exclude specific custom tools.

## Verifying

After adding a tool:

1. `npm run build` — must succeed. Check `dist/custom-tools/tools/<name>.js`
   exists.
2. `npm run dev` — log line should read `Custom tool registration complete: N
   registered, 0 skipped, 0 failed`.
3. `npm run inspector` — your tool appears in the list with the description
   and parameter schema you defined.
4. Call the tool from the inspector — it issues a real Graph request through
   the same `GraphClient` the built-in tools use.
