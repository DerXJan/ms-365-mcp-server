/**
 * Synchronous manifest loader for `custom-endpoints.json`.
 *
 * Lives in its own module (no Graph/auth imports) so `auth.ts` can call it
 * during synchronous scope-build without creating an import cycle through
 * `registry.ts` → `graph-tools.ts` → `auth.ts`.
 */
import { existsSync, readFileSync } from 'fs';
import path from 'path';
import { fileURLToPath } from 'url';
import logger from '../logger.js';
import type { CustomEndpointConfig, CustomToolsManifest } from './types.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const MANIFEST_PATH = path.join(__dirname, 'custom-endpoints.json');

let cachedEntries: CustomEndpointConfig[] | null = null;

/**
 * Load and cache the custom-endpoints manifest. Returns `[]` when the file is
 * missing or malformed; entries with `disabled: true` are filtered out (same
 * semantics as `endpoints.json`).
 */
export function loadCustomEndpointsSync(): CustomEndpointConfig[] {
  if (cachedEntries !== null) return cachedEntries;

  if (!existsSync(MANIFEST_PATH)) {
    cachedEntries = [];
    return cachedEntries;
  }

  try {
    const raw = readFileSync(MANIFEST_PATH, 'utf8');
    const parsed = JSON.parse(raw) as CustomToolsManifest;
    if (!parsed || !Array.isArray(parsed.tools)) {
      logger.warn(`custom-endpoints.json: expected { tools: [...] }, got ${typeof parsed}`);
      cachedEntries = [];
      return cachedEntries;
    }
    cachedEntries = parsed.tools.filter((e) => !e.disabled);
    if (cachedEntries.length > 0) {
      logger.info(
        `Loaded ${cachedEntries.length} custom Graph endpoint(s) from custom-endpoints.json`
      );
    }
    return cachedEntries;
  } catch (err) {
    logger.error(`Failed to read custom-endpoints.json: ${(err as Error).message}`);
    cachedEntries = [];
    return cachedEntries;
  }
}
