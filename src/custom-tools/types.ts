/**
 * Type definitions for the custom-tools extensibility layer.
 *
 * Custom tools are hand-authored Microsoft Graph endpoints that fill gaps the
 * codegen pipeline cannot cover (the Graph OpenAPI spec is incomplete for some
 * endpoints we need). They run through the *same* execution pipeline as the
 * generated tools — `executeGraphTool` from `graph-tools.ts` — so they inherit
 * every existing feature (OData normalization, multi-account routing,
 * pagination, sensitivity filtering, includeHeaders/excludeResponse, etc.).
 *
 * Two-file split per custom tool:
 *   1. JSON manifest entry  in   custom-endpoints.json   (sync-loadable, scope-discoverable)
 *   2. TS parameters module under  ./tools/<name>.ts     (Zod schemas, dynamic-imported)
 *
 * The split mirrors how the system already works: `endpoints.json` carries
 * metadata; `generated/client.ts` carries the parameter Zod schemas. We provide
 * the same separation for hand-authored endpoints.
 */
import type { z } from 'zod';
import type { EndpointConfig } from '../graph-tools.js';

/** A single parameter for a custom tool. Same shape as the parameters in the
 * generated client (`(typeof api.endpoints)[i].parameters[j]`), so the existing
 * Path/Query/Body/Header dispatch in `executeGraphTool` works unchanged. */
export interface CustomToolParameter {
  name: string;
  type: 'Path' | 'Query' | 'Body' | 'Header';
  description?: string;
  schema: z.ZodTypeAny;
}

export type CustomToolParameters = CustomToolParameter[];

/** Manifest entry — superset of `EndpointConfig`. JSON-serializable; loaded
 * synchronously at module load so OAuth scope resolution can see it before
 * the async tool-registration phase. */
export interface CustomEndpointConfig extends EndpointConfig {
  /** Description shown to the LLM as the tool description. The generated
   * client supplies this from the OpenAPI spec; for custom tools the author
   * provides it directly. */
  description: string;
  /** ESM module specifier for the parameters module, resolved relative to
   * `src/custom-tools/registry.ts`. Always end in `.js` to match the
   * project's NodeNext convention (works in dev under tsx and in built dist).
   * Example: `"./tools/list-mailbox-folders-extended.js"`. */
  parametersModule: string;
}

/** Shape that a parameters module is expected to export. */
export interface CustomToolModule {
  parameters: CustomToolParameters;
  /** Optional — influences media-content auto-detection in `executeGraphTool`
   * (graph-tools.ts:378-389). Most tools won't need it. */
  errors?: { description: string }[];
}

/** Top-level shape of `custom-endpoints.json`. */
export interface CustomToolsManifest {
  tools: CustomEndpointConfig[];
}
