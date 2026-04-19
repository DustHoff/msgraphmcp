import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerIntuneTools } from '../../src/tools/intune';

describe('Intune Tools', () => {
  let server: MockMcpServer;
  let graph: ReturnType<typeof createMockGraphClient>;

  beforeEach(() => {
    server = new MockMcpServer();
    graph = createMockGraphClient();
    registerIntuneTools(server as never, graph as never);
  });

  afterEach(() => jest.clearAllMocks());

  const EXPECTED_TOOLS = [
    'list_intune_apps', 'get_intune_app', 'create_intune_web_app', 'create_intune_store_app',
    'update_intune_app', 'delete_intune_app', 'list_intune_app_assignments',
    'assign_intune_app', 'get_intune_app_install_status',
    'list_device_configurations', 'get_device_configuration', 'create_device_configuration',
    'update_device_configuration', 'delete_device_configuration', 'assign_device_configuration',
    'get_device_configuration_assignments', 'get_device_configuration_device_status',
    'list_configuration_policies', 'get_configuration_policy', 'create_configuration_policy',
    'update_configuration_policy', 'delete_configuration_policy', 'assign_configuration_policy',
    'search_settings_catalog',
    'list_compliance_policies', 'get_compliance_policy', 'create_compliance_policy',
    'update_compliance_policy', 'delete_compliance_policy', 'assign_compliance_policy',
    'get_compliance_policy_device_status',
    'list_managed_devices', 'get_managed_device', 'sync_managed_device',
    'restart_managed_device', 'retire_managed_device', 'wipe_managed_device',
    'get_device_compliance_overview', 'list_intune_app_protection_policies',
  ];

  it('registers all Intune tools', () => {
    EXPECTED_TOOLS.forEach(name => expect(server.isRegistered(name)).toBe(true));
  });

  // ── App tests ─────────────────────────────────────────────────────────────

  describe('list_intune_apps', () => {
    it('calls getAll on mobileApps', async () => {
      graph.getAll.mockResolvedValue([{ id: 'a1', displayName: 'Office' }]);
      const result = await server.call('list_intune_apps', {});
      expect(graph.getAll).toHaveBeenCalledWith('/deviceAppManagement/mobileApps', expect.any(Object));
      expect(result.content[0].text).toContain('Office');
    });

    it('builds isof filter for appType', async () => {
      graph.getAll.mockResolvedValue([]);
      await server.call('list_intune_apps', { appType: '#microsoft.graph.webApp' });
      const [, params] = args(graph.getAll);
      expect(params.$filter).toContain("isof('#microsoft.graph.webApp')");
    });

    it('combines appType and custom filter with AND', async () => {
      graph.getAll.mockResolvedValue([]);
      await server.call('list_intune_apps', {
        appType: '#microsoft.graph.webApp',
        filter: 'isAssigned eq true',
      });
      const [, params] = args(graph.getAll);
      expect(params.$filter).toContain('isAssigned eq true');
      expect(params.$filter).toContain("isof('#microsoft.graph.webApp')");
    });
  });

  describe('create_intune_web_app', () => {
    it('posts WebApp type with appUrl', async () => {
      graph.post.mockResolvedValue({ id: 'app1' });
      await server.call('create_intune_web_app', {
        displayName: 'My Portal',
        publisher: 'Contoso',
        appUrl: 'https://portal.contoso.com',
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/deviceAppManagement/mobileApps');
      expect(body['@odata.type']).toBe('#microsoft.graph.webApp');
      expect(body.appUrl).toBe('https://portal.contoso.com');
    });
  });

  describe('create_intune_store_app', () => {
    it('maps iosStore to correct OData type', async () => {
      graph.post.mockResolvedValue({ id: 'app2' });
      await server.call('create_intune_store_app', {
        displayName: 'WhatsApp',
        publisher: 'Meta',
        storeType: 'iosStore',
        appStoreUrl: 'https://apps.apple.com/app/whatsapp',
      });
      const [, body] = args(graph.post);
      expect(body['@odata.type']).toBe('#microsoft.graph.iosStoreApp');
    });

    it('maps androidStore to correct OData type', async () => {
      graph.post.mockResolvedValue({ id: 'app3' });
      await server.call('create_intune_store_app', {
        displayName: 'Gmail',
        publisher: 'Google',
        storeType: 'androidStore',
        appStoreUrl: 'https://play.google.com/store/apps/details?id=com.google.android.gm',
      });
      const [, body] = args(graph.post);
      expect(body['@odata.type']).toBe('#microsoft.graph.androidStoreApp');
    });
  });

  describe('assign_intune_app', () => {
    it('posts mobileAppAssignments with correct intent', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('assign_intune_app', {
        appId: 'app1',
        assignments: [{ groupId: 'g1', intent: 'required' }],
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/deviceAppManagement/mobileApps/app1/assign');
      expect(body.mobileAppAssignments[0].intent).toBe('required');
      expect(body.mobileAppAssignments[0].target.groupId).toBe('g1');
    });

    it('requires at least one assignment', async () => {
      await expect(server.call('assign_intune_app', {
        appId: 'app1', assignments: [],
      })).rejects.toThrow();
    });
  });

  // ── Device configuration tests ────────────────────────────────────────────

  describe('create_device_configuration', () => {
    it('merges odata type and settings into body', async () => {
      graph.post.mockResolvedValue({ id: 'cfg1' });
      await server.call('create_device_configuration', {
        displayName: 'Win10 Password Policy',
        odataType: '#microsoft.graph.windows10GeneralConfiguration',
        settings: { passwordRequired: true, passwordMinimumLength: 8 },
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/deviceManagement/deviceConfigurations');
      expect(body['@odata.type']).toBe('#microsoft.graph.windows10GeneralConfiguration');
      expect(body.passwordRequired).toBe(true);
      expect(body.displayName).toBe('Win10 Password Policy');
    });
  });

  describe('assign_device_configuration', () => {
    it('builds include target for include intent', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('assign_device_configuration', {
        configId: 'cfg1',
        assignments: [{ groupId: 'g1', intent: 'include' }],
      });
      const [, body] = args(graph.post);
      expect(body.assignments[0].target['@odata.type']).toContain('groupAssignmentTarget');
    });

    it('builds exclusion target for exclude intent', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('assign_device_configuration', {
        configId: 'cfg1',
        assignments: [{ groupId: 'g2', intent: 'exclude' }],
      });
      const [, body] = args(graph.post);
      expect(body.assignments[0].target['@odata.type']).toContain('exclusionGroup');
    });
  });

  describe('get_configuration_policy', () => {
    it('fetches policy and settings', async () => {
      graph.get.mockResolvedValue({ id: 'p1', name: 'Test' });
      graph.getAll.mockResolvedValue([{ id: 's1' }]);
      const result = await server.call('get_configuration_policy', { policyId: 'p1' });
      expect(graph.get).toHaveBeenCalledWith('/deviceManagement/configurationPolicies/p1');
      expect(graph.getAll).toHaveBeenCalledWith('/deviceManagement/configurationPolicies/p1/settings');
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.policy).toBeDefined();
      expect(parsed.settings).toBeDefined();
    });
  });

  // ── Compliance policy tests ───────────────────────────────────────────────

  describe('create_compliance_policy', () => {
    it('includes default scheduledActionsForRule when not provided', async () => {
      graph.post.mockResolvedValue({ id: 'pol1' });
      await server.call('create_compliance_policy', {
        displayName: 'Windows Compliance',
        odataType: '#microsoft.graph.windows10CompliancePolicy',
        settings: { osMinimumVersion: '10.0.19041' },
      });
      const [, body] = args(graph.post);
      expect(body.scheduledActionsForRule).toBeDefined();
      expect(body.scheduledActionsForRule[0].scheduledActionConfigurations[0].actionType).toBe('block');
    });

    it('accepts custom scheduledActionsForRule', async () => {
      graph.post.mockResolvedValue({ id: 'pol2' });
      await server.call('create_compliance_policy', {
        displayName: 'iOS Compliance',
        odataType: '#microsoft.graph.iosCompliancePolicy',
        settings: {},
        scheduledActionsForRule: [{
          ruleName: 'DeviceLock',
          scheduledActionConfigurations: [{ actionType: 'retire', gracePeriodHours: 24 }],
        }],
      });
      const [, body] = args(graph.post);
      expect(body.scheduledActionsForRule[0].ruleName).toBe('DeviceLock');
      expect(body.scheduledActionsForRule[0].scheduledActionConfigurations[0].gracePeriodHours).toBe(24);
    });
  });

  // ── Managed device tests ──────────────────────────────────────────────────

  describe('list_managed_devices', () => {
    it('supports OData filter', async () => {
      graph.get.mockResolvedValue({ value: [] });
      await server.call('list_managed_devices', { filter: "operatingSystem eq 'Windows'" });
      const [, params] = args(graph.get);
      expect(params.$filter).toBe("operatingSystem eq 'Windows'");
    });
  });

  describe('sync_managed_device', () => {
    it('posts to syncDevice action', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('sync_managed_device', { deviceId: 'dev1' });
      expect(graph.post).toHaveBeenCalledWith(
        '/deviceManagement/managedDevices/dev1/syncDevice', {},
      );
    });
  });

  describe('wipe_managed_device', () => {
    it('posts to wipe with keepEnrollmentData flags', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('wipe_managed_device', {
        deviceId: 'dev1',
        keepEnrollmentData: true,
        keepUserData: false,
      });
      const [url, body] = args(graph.post);
      expect(url).toContain('wipe');
      expect(body.keepEnrollmentData).toBe(true);
      expect(body.keepUserData).toBe(false);
    });
  });

  describe('list_intune_app_protection_policies', () => {
    it('fetches both iOS and Android when platform=all', async () => {
      graph.getAll.mockResolvedValue([]);
      await server.call('list_intune_app_protection_policies', { platform: 'all' });
      expect(graph.getAll).toHaveBeenCalledTimes(2);
    });

    it('fetches only iOS when platform=ios', async () => {
      graph.getAll.mockResolvedValue([{ id: 'p1' }]);
      await server.call('list_intune_app_protection_policies', { platform: 'ios' });
      expect(graph.getAll).toHaveBeenCalledTimes(1);
      const [url] = args(graph.getAll);
      expect(url).toContain('ios');
    });
  });
});
