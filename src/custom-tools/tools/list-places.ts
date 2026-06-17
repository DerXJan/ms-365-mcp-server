import { z } from 'zod';
import type { CustomToolParameters } from '../types.js';

/**
 * Parameters for `GET /places/{placeType}`.
 *
 * Spec: https://learn.microsoft.com/en-us/graph/api/place-list?view=graph-rest-1.0
 *
 * Notes for future maintainers:
 *   - $filter and $count=true are only supported when placeType is one of
 *     microsoft.graph.room, microsoft.graph.workspace, or microsoft.graph.roomList.
 *     For other place types Graph returns 400. The placeType-specific restriction
 *     is surfaced via the manifest's `llmTip` so the LLM sees it in the tool
 *     description (the OData override in `buildParamSchemaForTool` rewrites the
 *     per-parameter description, so we can't put it on $filter/$count directly).
 *   - We deliberately do not declare $orderby, $expand, or $search — they are
 *     not supported by this endpoint per the spec.
 *   - OData params are declared without the leading $ to match the project
 *     convention (MCP clients can't always send `$` in parameter names; the
 *     runtime normalizes both forms — see [[gotchas]]).
 */
export const parameters: CustomToolParameters = [
  {
    name: 'placeType',
    type: 'Path',
    description:
      'OData type literal of the place collection to list. Common values: ' +
      'microsoft.graph.room, microsoft.graph.workspace, microsoft.graph.roomList, ' +
      'microsoft.graph.building, microsoft.graph.desk.',
    schema: z
      .string()
      .describe(
        'OData type of the place collection (e.g. "microsoft.graph.room", ' +
          '"microsoft.graph.workspace", "microsoft.graph.roomList", ' +
          '"microsoft.graph.building", "microsoft.graph.desk").'
      ),
  },
  {
    name: 'select',
    type: 'Query',
    schema: z.string().optional(),
  },
  {
    name: 'top',
    type: 'Query',
    schema: z.number().optional(),
  },
  {
    name: 'skip',
    type: 'Query',
    schema: z.number().optional(),
  },
  {
    name: 'filter',
    type: 'Query',
    // Only valid for room/workspace/roomList — surfaced in the manifest llmTip.
    schema: z.string().optional(),
  },
  {
    name: 'count',
    type: 'Query',
    // Only valid for room/workspace/roomList — surfaced in the manifest llmTip.
    schema: z.boolean().optional(),
  },
];
