import { beforeEach, afterEach, describe, expect, it, vi } from 'vitest';
import {
  loadSensitivityFilterConfig,
  resetSensitivityFilterConfig,
  getToolSensitivityContext,
  filterListResponse,
  checkSingleItem,
  filterDocumentListResponse,
  checkSingleDocument,
} from '../src/lib/sensitivity-filter.js';

vi.mock('../src/logger.js', () => ({
  default: {
    info: vi.fn(),
    error: vi.fn(),
    warn: vi.fn(),
  },
}));

describe('Sensitivity Filter', () => {
  beforeEach(() => {
    resetSensitivityFilterConfig();
    delete process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS;
  });

  afterEach(() => {
    resetSensitivityFilterConfig();
    delete process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS;
  });

  describe('loadSensitivityFilterConfig', () => {
    it('returns disabled config when env var is not set', () => {
      const config = loadSensitivityFilterConfig();
      expect(config.enabled).toBe(false);
      expect(config.blockedLabels.size).toBe(0);
    });

    it('returns disabled config when env var is empty', () => {
      process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS = '';
      const config = loadSensitivityFilterConfig();
      expect(config.enabled).toBe(false);
    });

    it('parses comma-separated labels', () => {
      process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS = 'confidential,private';
      const config = loadSensitivityFilterConfig();
      expect(config.enabled).toBe(true);
      expect(config.blockedLabels.has('confidential')).toBe(true);
      expect(config.blockedLabels.has('private')).toBe(true);
    });

    it('trims whitespace and lowercases labels', () => {
      process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS = ' Confidential , PRIVATE ';
      const config = loadSensitivityFilterConfig();
      expect(config.blockedLabels.has('confidential')).toBe(true);
      expect(config.blockedLabels.has('private')).toBe(true);
    });

    it('handles GUIDs', () => {
      process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS =
        'a1b2c3d4-e5f6-7890-abcd-ef1234567890';
      const config = loadSensitivityFilterConfig();
      expect(config.blockedLabels.has('a1b2c3d4-e5f6-7890-abcd-ef1234567890')).toBe(true);
    });

    it('handles mixed names and GUIDs', () => {
      process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS =
        'confidential,a1b2c3d4-e5f6-7890-abcd-ef1234567890,Highly Confidential';
      const config = loadSensitivityFilterConfig();
      expect(config.blockedLabels.size).toBe(3);
      expect(config.blockedLabels.has('confidential')).toBe(true);
      expect(config.blockedLabels.has('a1b2c3d4-e5f6-7890-abcd-ef1234567890')).toBe(true);
      expect(config.blockedLabels.has('highly confidential')).toBe(true);
    });

    it('caches config across calls', () => {
      process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS = 'confidential';
      const config1 = loadSensitivityFilterConfig();
      process.env.MS365_MCP_BLOCKED_SENSITIVITY_LABELS = 'private';
      const config2 = loadSensitivityFilterConfig();
      expect(config1).toBe(config2);
    });
  });

  describe('getToolSensitivityContext', () => {
    it('returns email for mail tools', () => {
      expect(getToolSensitivityContext('list-mail-messages')).toBe('email');
      expect(getToolSensitivityContext('get-mail-message')).toBe('email');
      expect(getToolSensitivityContext('list-mail-folder-messages')).toBe('email');
      expect(getToolSensitivityContext('list-shared-mailbox-messages')).toBe('email');
      expect(getToolSensitivityContext('get-shared-mailbox-message')).toBe('email');
      expect(getToolSensitivityContext('list-shared-mailbox-folder-messages')).toBe('email');
    });

    it('returns document for drive tools', () => {
      expect(getToolSensitivityContext('list-folder-files')).toBe('document');
      expect(getToolSensitivityContext('get-drive-item')).toBe('document');
      expect(getToolSensitivityContext('download-onedrive-file-content')).toBe('document');
      expect(getToolSensitivityContext('search-onedrive-files')).toBe('document');
    });

    it('returns none for unrelated tools', () => {
      expect(getToolSensitivityContext('list-calendar-events')).toBe('none');
      expect(getToolSensitivityContext('send-mail')).toBe('none');
      expect(getToolSensitivityContext('get-current-user')).toBe('none');
    });
  });

  describe('filterListResponse (email)', () => {
    it('returns all items when filter is disabled', () => {
      const config = { blockedLabels: new Set<string>(), enabled: false };
      const items = [{ id: '1', sensitivity: 'confidential' }];
      const { filtered, removedCount } = filterListResponse(items, 'email', config);
      expect(filtered).toHaveLength(1);
      expect(removedCount).toBe(0);
    });

    it('filters emails by sensitivity field', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const items = [
        { id: '1', subject: 'Public', sensitivity: 'normal' },
        { id: '2', subject: 'Secret', sensitivity: 'confidential' },
        { id: '3', subject: 'Private', sensitivity: 'private' },
      ];
      const { filtered, removedCount } = filterListResponse(items, 'email', config);
      expect(filtered).toHaveLength(2);
      expect(removedCount).toBe(1);
      expect(filtered).toEqual([
        { id: '1', subject: 'Public', sensitivity: 'normal' },
        { id: '3', subject: 'Private', sensitivity: 'private' },
      ]);
    });

    it('filters emails case-insensitively', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const items = [
        { id: '1', sensitivity: 'Confidential' },
        { id: '2', sensitivity: 'CONFIDENTIAL' },
        { id: '3', sensitivity: 'normal' },
      ];
      const { filtered, removedCount } = filterListResponse(items, 'email', config);
      expect(filtered).toHaveLength(1);
      expect(removedCount).toBe(2);
    });

    it('allows emails without sensitivity field (fail-open)', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const items = [
        { id: '1', subject: 'No sensitivity field' },
        { id: '2', subject: 'Has field', sensitivity: 'confidential' },
      ];
      const { filtered, removedCount } = filterListResponse(items, 'email', config);
      expect(filtered).toHaveLength(1);
      expect(removedCount).toBe(1);
      expect((filtered[0] as any).id).toBe('1');
    });

    it('does not filter documents (deferred to async method)', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const items = [{ id: '1', name: 'doc.docx' }];
      const { filtered, removedCount } = filterListResponse(items, 'document', config);
      expect(filtered).toHaveLength(1);
      expect(removedCount).toBe(0);
    });
  });

  describe('checkSingleItem (email)', () => {
    it('returns not blocked when filter is disabled', () => {
      const config = { blockedLabels: new Set<string>(), enabled: false };
      const item = { id: '1', sensitivity: 'confidential' };
      const result = checkSingleItem(item, 'email', config);
      expect(result.blocked).toBe(false);
    });

    it('blocks a single email with matching sensitivity', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const item = { id: '1', subject: 'Secret', sensitivity: 'confidential' };
      const result = checkSingleItem(item, 'email', config);
      expect(result.blocked).toBe(true);
      expect(result.label).toBe('confidential');
    });

    it('allows a single email without matching sensitivity', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const item = { id: '1', subject: 'Normal', sensitivity: 'normal' };
      const result = checkSingleItem(item, 'email', config);
      expect(result.blocked).toBe(false);
    });

    it('does not check documents (deferred to async method)', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const item = { id: '1', name: 'secret.docx' };
      const result = checkSingleItem(item, 'document', config);
      expect(result.blocked).toBe(false);
    });

    it('allows null/undefined items (fail-open)', () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      expect(checkSingleItem(null, 'email', config).blocked).toBe(false);
      expect(checkSingleItem(undefined, 'email', config).blocked).toBe(false);
    });
  });

  describe('filterDocumentListResponse', () => {
    function mockGraphClient(
      catalogLabels: Array<{ id: string; name: string }>,
      itemLabels: Record<string, string[]>
    ) {
      return {
        makeRequest: vi.fn(async (endpoint: string, options?: { method?: string }) => {
          if (endpoint.includes('/informationProtection/sensitivityLabels')) {
            return { value: catalogLabels };
          }
          if (options?.method === 'POST' && endpoint.includes('/extractSensitivityLabels')) {
            const itemId = endpoint.split('/items/')[1]?.split('/')[0];
            const labels = (itemLabels[itemId!] || []).map((id) => ({ sensitivityLabelId: id }));
            return { labels };
          }
          return {};
        }),
      };
    }

    it('returns all items when filter is disabled', async () => {
      const config = { blockedLabels: new Set<string>(), enabled: false };
      const client = mockGraphClient([], {});
      const items = [{ id: '1', parentReference: { driveId: 'd1' } }];
      const { filtered, removedCount } = await filterDocumentListResponse(items, config, client);
      expect(filtered).toHaveLength(1);
      expect(removedCount).toBe(0);
    });

    it('filters documents by label ID resolved from display name', async () => {
      const config = { blockedLabels: new Set(['confidential (non-encrypted)']), enabled: true };
      const client = mockGraphClient(
        [
          { id: 'label-guid-1', name: 'Confidential (Non-Encrypted)' },
          { id: 'label-guid-2', name: 'Public' },
        ],
        {
          item1: ['label-guid-1'],
          item2: ['label-guid-2'],
        }
      );
      const items = [
        { id: 'item1', name: 'secret.docx', parentReference: { driveId: 'd1' } },
        { id: 'item2', name: 'public.docx', parentReference: { driveId: 'd1' } },
      ];
      const { filtered, removedCount } = await filterDocumentListResponse(items, config, client);
      expect(removedCount).toBe(1);
      expect(filtered).toHaveLength(1);
      expect((filtered[0] as any).name).toBe('public.docx');
    });

    it('filters documents by GUID directly', async () => {
      const config = { blockedLabels: new Set(['label-guid-1']), enabled: true };
      const client = mockGraphClient(
        [{ id: 'label-guid-1', name: 'Some Label' }],
        { item1: ['label-guid-1'] }
      );
      const items = [
        { id: 'item1', name: 'file.xlsx', parentReference: { driveId: 'd1' } },
      ];
      const { filtered, removedCount } = await filterDocumentListResponse(items, config, client);
      expect(removedCount).toBe(1);
      expect(filtered).toHaveLength(0);
    });

    it('allows documents without labels (fail-open)', async () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const client = mockGraphClient(
        [{ id: 'label-guid-1', name: 'Confidential' }],
        { item1: [] }
      );
      const items = [
        { id: 'item1', name: 'unlabeled.docx', parentReference: { driveId: 'd1' } },
      ];
      const { filtered, removedCount } = await filterDocumentListResponse(items, config, client);
      expect(removedCount).toBe(0);
      expect(filtered).toHaveLength(1);
    });

    it('allows documents without parentReference (fail-open)', async () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const client = mockGraphClient(
        [{ id: 'label-guid-1', name: 'Confidential' }],
        {}
      );
      const items = [{ id: 'item1', name: 'orphan.docx' }];
      const { filtered, removedCount } = await filterDocumentListResponse(items, config, client);
      expect(removedCount).toBe(0);
      expect(filtered).toHaveLength(1);
    });
  });

  describe('checkSingleDocument', () => {
    function mockGraphClient(
      catalogLabels: Array<{ id: string; name: string }>,
      itemLabels: string[]
    ) {
      return {
        makeRequest: vi.fn(async (endpoint: string, options?: { method?: string }) => {
          if (endpoint.includes('/informationProtection/sensitivityLabels')) {
            return { value: catalogLabels };
          }
          if (options?.method === 'POST' && endpoint.includes('/extractSensitivityLabels')) {
            return { labels: itemLabels.map((id) => ({ sensitivityLabelId: id })) };
          }
          return {};
        }),
      };
    }

    it('returns not blocked when disabled', async () => {
      const config = { blockedLabels: new Set<string>(), enabled: false };
      const client = mockGraphClient([], []);
      const item = { id: 'item1', parentReference: { driveId: 'd1' } };
      const result = await checkSingleDocument(item, config, client);
      expect(result.blocked).toBe(false);
    });

    it('blocks a document with matching label', async () => {
      const config = { blockedLabels: new Set(['confidential (non-encrypted)']), enabled: true };
      const client = mockGraphClient(
        [{ id: 'label-guid-1', name: 'Confidential (Non-Encrypted)' }],
        ['label-guid-1']
      );
      const item = { id: 'item1', name: 'secret.docx', parentReference: { driveId: 'd1' } };
      const result = await checkSingleDocument(item, config, client);
      expect(result.blocked).toBe(true);
      expect(result.label).toBe('Confidential (Non-Encrypted)');
    });

    it('allows a document without matching label', async () => {
      const config = { blockedLabels: new Set(['confidential']), enabled: true };
      const client = mockGraphClient(
        [
          { id: 'label-guid-1', name: 'Confidential' },
          { id: 'label-guid-2', name: 'Public' },
        ],
        ['label-guid-2']
      );
      const item = { id: 'item1', name: 'public.docx', parentReference: { driveId: 'd1' } };
      const result = await checkSingleDocument(item, config, client);
      expect(result.blocked).toBe(false);
    });
  });
});
