/**
 * Custom-tools registry: loads `custom-endpoints.json`, dynamic-imports each
 * tool's parameters module, and registers a synthesized tool object with the
 * MCP server. The synthesized object matches `(typeof api.endpoints)[0]` from
 * the generated client, so `executeGraphTool` runs unchanged for both
 * generated and custom tools.
 *
 * See `types.ts` for the data shapes and `README.md` for the author guide.
 */
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import path from 'path';
import { fileURLToPath, pathToFileURL } from 'url';
import { z } from 'zod';
import logger from '../logger.js';
import GraphClient from '../graph-client.js';
import AuthManager from '../auth.js';
import {
  buildParamSchemaForTool,
  executeGraphTool,
  type EndpointConfig,
} from '../graph-tools.js';
import { loadCustomEndpointsSync } from './manifest.js';
import type { CustomEndpointConfig, CustomToolModule } from './types.js';

// Re-export so existing callers (tests, server.ts) can keep importing from registry.
export { loadCustomEndpointsSync } from './manifest.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

/**
 * Convert `{name}` placeholders in pathPattern to `:name` style, matching
 * what the generated client uses and what `executeGraphTool` substitutes
 * against (see graph-tools.ts:240-244 and the fallback at 283-298).
 */
function pathPatternToColonStyle(pathPattern: string): string {
  return pathPattern.replace(/\{([a-zA-Z][a-zA-Z0-9_-]*)\}/g, ':$1');
}

/**
 * Build the Zod parameter schema fragment from the parameters module export.
 * Mirrors the loop the generated client implicitly applies when populating
 * `tool.parameters[i].schema`.
 */
function synthesizeToolFromEntry(
  entry: CustomEndpointConfig,
  mod: CustomToolModule
): {
  alias: string;
  method: string;
  path: string;
  description: string;
  parameters: { name: string; type: 'Path' | 'Query' | 'Body' | 'Header'; schema: z.ZodTypeAny }[];
  errors: { description: string }[];
  requestFormat: 'json';
  response: z.ZodTypeAny;
} {
  return {
    alias: entry.toolName,
    method: entry.method,
    path: pathPatternToColonStyle(entry.pathPattern),
    description: entry.description,
    parameters: mod.parameters.map((p) => ({
      name: p.name,
      type: p.type,
      schema: p.schema,
    })),
    errors: mod.errors ?? [],
    requestFormat: 'json',
    response: z.any(),
  };
}

/**
 * Register all enabled custom Graph tools with the MCP server. Mirrors the
 * filter/registration cascade in `registerGraphTools` so behavior stays
 * symmetric (same flags, same logs, same skip reasons).
 *
 * Returns the number of tools that were successfully registered.
 */
export async function registerCustomTools(
  server: McpServer,
  graphClient: GraphClient,
  readOnly: boolean = false,
  enabledToolsPattern?: string,
  orgMode: boolean = false,
  authManager?: AuthManager,
  multiAccount: boolean = false,
  accountNames: string[] = []
): Promise<number> {
  const entries = loadCustomEndpointsSync();
  if (entries.length === 0) {
    return 0;
  }

  let enabledToolsRegex: RegExp | undefined;
  if (enabledToolsPattern) {
    try {
      enabledToolsRegex = new RegExp(enabledToolsPattern, 'i');
    } catch {
      // already logged by registerGraphTools — stay silent here
    }
  }

  let registeredCount = 0;
  let skippedCount = 0;
  let failedCount = 0;

  for (const entry of entries) {
    if (!orgMode && !entry.scopes && entry.workScopes) {
      logger.info(`Skipping work-account custom tool ${entry.toolName} - not in org mode`);
      skippedCount++;
      continue;
    }

    const method = entry.method.toUpperCase();
    if (readOnly && method !== 'GET') {
      // Allow POST endpoints explicitly marked readOnly (mirrors graph-tools.ts:629-639)
      if (!(method === 'POST' && entry.readOnly)) {
        logger.info(`Skipping write operation ${entry.toolName} in read-only mode`);
        skippedCount++;
        continue;
      }
    }

    if (enabledToolsRegex && !enabledToolsRegex.test(entry.toolName)) {
      logger.info(`Skipping custom tool ${entry.toolName} - doesn't match filter pattern`);
      skippedCount++;
      continue;
    }

    let mod: CustomToolModule;
    try {
      // Resolve the parameters module relative to this file. We convert to a
      // file:// URL so dynamic import works on Windows where bare relative
      // paths are not accepted by ESM's import().
      const modulePath = path.resolve(__dirname, entry.parametersModule);
      const moduleUrl = pathToFileURL(modulePath).href;
      const imported = (await import(moduleUrl)) as Partial<CustomToolModule>;
      if (!imported.parameters || !Array.isArray(imported.parameters)) {
        throw new Error(
          `Module ${entry.parametersModule} must export a 'parameters' array (got ${typeof imported.parameters})`
        );
      }
      mod = {
        parameters: imported.parameters,
        errors: imported.errors,
      };
    } catch (err) {
      logger.error(
        `Failed to load parameters module for custom tool ${entry.toolName}: ${(err as Error).message}`
      );
      failedCount++;
      continue;
    }

    const synthTool = synthesizeToolFromEntry(entry, mod);
    const endpointConfig: EndpointConfig = entry; // CustomEndpointConfig extends EndpointConfig

    const paramSchema = buildParamSchemaForTool(
      synthTool,
      endpointConfig,
      multiAccount,
      accountNames
    );

    let toolDescription =
      entry.description || `Execute ${method} request to ${entry.pathPattern}`;
    if (entry.llmTip) {
      toolDescription += `\n\n💡 TIP: ${entry.llmTip}`;
    }

    try {
      server.tool(
        synthTool.alias,
        toolDescription,
        paramSchema,
        {
          title: synthTool.alias,
          readOnlyHint: method === 'GET',
          destructiveHint: ['POST', 'PATCH', 'DELETE'].includes(method),
          openWorldHint: true,
        },
        async (params) =>
          executeGraphTool(
            synthTool as unknown as Parameters<typeof executeGraphTool>[0],
            endpointConfig,
            graphClient,
            params,
            authManager
          )
      );
      registeredCount++;
    } catch (err) {
      logger.error(
        `Failed to register custom tool ${entry.toolName}: ${(err as Error).message}`
      );
      failedCount++;
    }
  }

  logger.info(
    `Custom tool registration complete: ${registeredCount} registered, ${skippedCount} skipped, ${failedCount} failed`
  );
  return registeredCount;
}
