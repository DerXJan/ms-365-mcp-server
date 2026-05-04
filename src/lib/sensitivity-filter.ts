import logger from '../logger.js';

export interface SensitivityFilterConfig {
  blockedLabels: Set<string>;
  enabled: boolean;
}

interface LabelCatalogEntry {
  id: string;
  name: string;
}

interface GraphClientLike {
  makeRequest(endpoint: string, options?: { method?: string; accessToken?: string }): Promise<unknown>;
}

const EMAIL_TOOLS = new Set([
  'list-mail-messages',
  'get-mail-message',
  'list-mail-folder-messages',
  'list-shared-mailbox-messages',
  'get-shared-mailbox-message',
  'list-shared-mailbox-folder-messages',
]);

const DOCUMENT_TOOLS = new Set([
  'list-folder-files',
  'get-drive-item',
  'download-onedrive-file-content',
  'search-onedrive-files',
]);

let cachedConfig: SensitivityFilterConfig | undefined;
let labelCatalog: LabelCatalogEntry[] | undefined;
let blockedLabelIds: Set<string> | undefined;

export function loadSensitivityFilterConfig(): SensitivityFilterConfig {
  if (cachedConfig) return cachedConfig;

  const raw = process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS;
  if (!raw || raw.trim() === '') {
    cachedConfig = { blockedLabels: new Set(), enabled: false };
    return cachedConfig;
  }

  const labels = raw
    .split(',')
    .map((l) => l.trim().toLowerCase())
    .filter((l) => l.length > 0);

  cachedConfig = { blockedLabels: new Set(labels), enabled: labels.length > 0 };

  if (cachedConfig.enabled) {
    logger.info(
      `Sensitivity filter enabled — blocking labels: ${[...cachedConfig.blockedLabels].join(', ')}`
    );
  }

  return cachedConfig;
}

export function resetSensitivityFilterConfig(): void {
  cachedConfig = undefined;
  labelCatalog = undefined;
  blockedLabelIds = undefined;
}

export function getToolSensitivityContext(toolName: string): 'email' | 'document' | 'none' {
  if (EMAIL_TOOLS.has(toolName)) return 'email';
  if (DOCUMENT_TOOLS.has(toolName)) return 'document';
  return 'none';
}

async function fetchLabelCatalog(
  graphClient: GraphClientLike,
  accessToken?: string
): Promise<LabelCatalogEntry[]> {
  if (labelCatalog) return labelCatalog;

  // Try the v1.0 security namespace first, then fall back to the older beta-era path.
  const endpoints = [
    '/security/informationProtection/sensitivityLabels',
    '/me/informationProtection/policy/labels',
  ];

  for (const endpoint of endpoints) {
    try {
      const result = (await graphClient.makeRequest(endpoint, {
        method: 'GET',
        accessToken,
      })) as { value?: Array<{ id: string; name?: string }> };

      if (result?.value && Array.isArray(result.value) && result.value.length > 0) {
        labelCatalog = result.value.map((l) => ({ id: l.id, name: l.name || '' }));
        logger.info(
          `Sensitivity label catalog loaded (${endpoint}): ${labelCatalog.length} labels`
        );
        return labelCatalog;
      }
    } catch (err) {
      logger.info(
        `Label catalog endpoint ${endpoint} failed: ${(err as Error).message}`
      );
    }
  }

  logger.warn(
    'Could not fetch sensitivity label catalog from any endpoint. ' +
      'Ensure the app has InformationProtectionPolicy.Read permission. ' +
      'You can also specify label GUIDs directly in MS365_MCP_BLOCKED_SENSITIVITY_LABELS.'
  );
  labelCatalog = [];
  return labelCatalog;
}

async function resolveBlockedLabelIds(
  config: SensitivityFilterConfig,
  graphClient: GraphClientLike,
  accessToken?: string
): Promise<Set<string>> {
  if (blockedLabelIds) return blockedLabelIds;

  const catalog = await fetchLabelCatalog(graphClient, accessToken);
  blockedLabelIds = new Set<string>();

  for (const entry of catalog) {
    if (config.blockedLabels.has(entry.name.toLowerCase())) {
      blockedLabelIds.add(entry.id.toLowerCase());
    }
    if (config.blockedLabels.has(entry.id.toLowerCase())) {
      blockedLabelIds.add(entry.id.toLowerCase());
    }
  }

  // Also add any raw values that look like GUIDs directly
  for (const label of config.blockedLabels) {
    if (/^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/.test(label)) {
      blockedLabelIds.add(label);
    }
  }

  logger.info(
    `Resolved blocked label IDs: ${[...blockedLabelIds].join(', ') || '(none resolved)'}`
  );

  return blockedLabelIds;
}

async function extractItemLabels(
  graphClient: GraphClientLike,
  driveId: string,
  itemId: string,
  accessToken?: string
): Promise<string[]> {
  try {
    const result = (await graphClient.makeRequest(
      `/drives/${driveId}/items/${itemId}/extractSensitivityLabels`,
      { method: 'POST', accessToken }
    )) as { labels?: Array<{ sensitivityLabelId: string }> };

    if (result?.labels && Array.isArray(result.labels)) {
      return result.labels.map((l) => l.sensitivityLabelId.toLowerCase());
    }
  } catch (err) {
    const msg = (err as Error).message;
    if (msg.includes('403') || msg.includes('scope') || msg.includes('permission')) {
      logger.warn(
        `extractSensitivityLabels permission denied for ${itemId}. ` +
          'Ensure the app has Files.Read.All or Sites.Read.All permission.'
      );
    } else {
      logger.info(
        `extractSensitivityLabels failed for ${driveId}/${itemId}: ${msg}`
      );
    }
  }
  return [];
}

function extractDriveAndItemId(item: Record<string, unknown>): { driveId: string; itemId: string } | null {
  const itemId = item.id as string | undefined;
  if (!itemId) return null;

  const parentRef = item.parentReference as { driveId?: string } | undefined;
  if (parentRef?.driveId) {
    return { driveId: parentRef.driveId, itemId };
  }

  return null;
}

function isEmailBlocked(
  item: Record<string, unknown>,
  blockedLabels: Set<string>
): { blocked: boolean; label?: string } {
  const sensitivity = item.sensitivity;
  if (typeof sensitivity !== 'string' || sensitivity === '') return { blocked: false };

  if (blockedLabels.has(sensitivity.toLowerCase())) {
    return { blocked: true, label: sensitivity };
  }
  return { blocked: false };
}

export function filterListResponse(
  items: unknown[],
  context: 'email' | 'document',
  config: SensitivityFilterConfig
): { filtered: unknown[]; removedCount: number } {
  if (!config.enabled || context !== 'email') return { filtered: items, removedCount: 0 };

  const filtered: unknown[] = [];
  let removedCount = 0;

  for (const item of items) {
    if (item && typeof item === 'object') {
      const { blocked } = isEmailBlocked(item as Record<string, unknown>, config.blockedLabels);
      if (blocked) {
        removedCount++;
      } else {
        filtered.push(item);
      }
    } else {
      filtered.push(item);
    }
  }

  return { filtered, removedCount };
}

export function checkSingleItem(
  item: unknown,
  context: 'email' | 'document',
  config: SensitivityFilterConfig
): { blocked: boolean; label?: string } {
  if (!config.enabled || context !== 'email') return { blocked: false };
  if (!item || typeof item !== 'object') return { blocked: false };
  return isEmailBlocked(item as Record<string, unknown>, config.blockedLabels);
}

export async function filterDocumentListResponse(
  items: unknown[],
  config: SensitivityFilterConfig,
  graphClient: GraphClientLike,
  accessToken?: string
): Promise<{ filtered: unknown[]; removedCount: number }> {
  if (!config.enabled) return { filtered: items, removedCount: 0 };

  const resolved = await resolveBlockedLabelIds(config, graphClient, accessToken);
  if (resolved.size === 0) {
    logger.warn(
      'Sensitivity filter: no blocked label IDs could be resolved — allowing all documents through'
    );
    return { filtered: items, removedCount: 0 };
  }

  const filtered: unknown[] = [];
  let removedCount = 0;

  for (const item of items) {
    if (!item || typeof item !== 'object') {
      filtered.push(item);
      continue;
    }

    const ids = extractDriveAndItemId(item as Record<string, unknown>);
    if (!ids) {
      filtered.push(item);
      continue;
    }

    const labelIds = await extractItemLabels(graphClient, ids.driveId, ids.itemId, accessToken);
    const isBlocked = labelIds.some((lid) => resolved.has(lid));

    if (isBlocked) {
      removedCount++;
    } else {
      filtered.push(item);
    }
  }

  return { filtered, removedCount };
}

export async function checkSingleDocument(
  item: unknown,
  config: SensitivityFilterConfig,
  graphClient: GraphClientLike,
  accessToken?: string
): Promise<{ blocked: boolean; label?: string }> {
  if (!config.enabled) return { blocked: false };
  if (!item || typeof item !== 'object') return { blocked: false };

  const resolved = await resolveBlockedLabelIds(config, graphClient, accessToken);
  if (resolved.size === 0) return { blocked: false };

  const ids = extractDriveAndItemId(item as Record<string, unknown>);
  if (!ids) return { blocked: false };

  const labelIds = await extractItemLabels(graphClient, ids.driveId, ids.itemId, accessToken);
  const blockedId = labelIds.find((lid) => resolved.has(lid));

  if (blockedId) {
    const catalog = labelCatalog || [];
    const entry = catalog.find((e) => e.id.toLowerCase() === blockedId);
    return { blocked: true, label: entry?.name || blockedId };
  }

  return { blocked: false };
}
