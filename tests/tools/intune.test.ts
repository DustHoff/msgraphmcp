import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import AdmZip from 'adm-zip';
import { MockMcpServer } from '../helpers/MockMcpServer';
import { args, createMockGraphClient } from '../helpers/mockGraphClient';
import { registerIntuneTools } from '../../src/tools/intune';

// ── Synthetic .intunewin builder ──────────────────────────────────────────
// IntuneWinAppUtil-produced packages have Detection.xml + an encrypted blob.
// Tests need only the XML; the blob is replaced with a placeholder because
// the upload step is bypassed before any byte is touched.
const MSI_DETECTION_XML = `<?xml version="1.0" encoding="utf-8"?>
<ApplicationInfo ToolVersion="1.8.4.0">
  <FileName>IntunePackage.intunewin</FileName>
  <UnencryptedContentSize>1024</UnencryptedContentSize>
  <EncryptionInfo>
    <EncryptionKey>AAAA</EncryptionKey>
    <MacKey>BBBB</MacKey>
    <InitializationVector>CCCC</InitializationVector>
    <Mac>DDDD</Mac>
    <FileDigest>EEEE</FileDigest>
    <FileDigestAlgorithm>SHA256</FileDigestAlgorithm>
    <ProfileIdentifier>ProfileVersion1</ProfileIdentifier>
  </EncryptionInfo>
  <MsiInfo>
    <MsiProductCode>{12345678-1234-1234-1234-123456789012}</MsiProductCode>
    <MsiProductVersion>1.2.3</MsiProductVersion>
    <MsiPackageCode>{AAAAAAAA-AAAA-AAAA-AAAA-AAAAAAAAAAAA}</MsiPackageCode>
    <MsiUpgradeCode>{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}</MsiUpgradeCode>
    <MsiExecutionContext>System</MsiExecutionContext>
    <MsiRequiresReboot>false</MsiRequiresReboot>
    <MsiIsMachineInstall>true</MsiIsMachineInstall>
    <MsiIsUserInstall>false</MsiIsUserInstall>
    <MsiPublisher>Test Corp</MsiPublisher>
  </MsiInfo>
</ApplicationInfo>`;

const EXE_DETECTION_XML = `<?xml version="1.0" encoding="utf-8"?>
<ApplicationInfo ToolVersion="1.8.4.0">
  <FileName>IntunePackage.intunewin</FileName>
  <UnencryptedContentSize>1024</UnencryptedContentSize>
  <EncryptionInfo>
    <EncryptionKey>AAAA</EncryptionKey>
    <MacKey>BBBB</MacKey>
    <InitializationVector>CCCC</InitializationVector>
    <Mac>DDDD</Mac>
    <FileDigest>EEEE</FileDigest>
    <FileDigestAlgorithm>SHA256</FileDigestAlgorithm>
    <ProfileIdentifier>ProfileVersion1</ProfileIdentifier>
  </EncryptionInfo>
</ApplicationInfo>`;

function makeIntuneWin(detectionXml: string): string {
  const zip = new AdmZip();
  zip.addFile('IntuneWinPackage/Metadata/Detection.xml', Buffer.from(detectionXml));
  zip.addFile('IntuneWinPackage/Contents/IntunePackage.intunewin', Buffer.from('dummy-payload'));
  const tmpPath = path.join(
    os.tmpdir(),
    `intunewin-test-${Date.now()}-${Math.random().toString(36).slice(2)}.intunewin`,
  );
  zip.writeZip(tmpPath);
  return tmpPath;
}

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
    'list_notification_templates', 'get_notification_template', 'create_notification_template',
    'update_notification_template', 'delete_notification_template',
    'add_notification_template_message', 'send_notification_template_test',
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
      graph.beta.get.mockResolvedValue({ id: 'p1', name: 'Test' });
      graph.beta.getAll.mockResolvedValue([{ id: 's1' }]);
      const result = await server.call('get_configuration_policy', { policyId: 'p1' });
      expect(graph.beta.get).toHaveBeenCalledWith('/deviceManagement/configurationPolicies/p1');
      expect(graph.beta.getAll).toHaveBeenCalledWith('/deviceManagement/configurationPolicies/p1/settings');
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

  // ── Notification template tests ───────────────────────────────────────────

  describe('create_notification_template', () => {
    it('posts template with required fields via beta (v1.0 StatelessNotificationFEService rejects writes)', async () => {
      graph.beta.post.mockResolvedValue({ id: 'tmpl1', displayName: 'Test' });
      const result = await server.call('create_notification_template', {
        displayName: 'Device Non-Compliance',
        defaultLocale: 'en-US',
        brandingOptions: 'includeCompanyName',
      });
      const [url, body] = args(graph.beta.post);
      expect(url).toBe('/deviceManagement/notificationMessageTemplates');
      expect(body['@odata.type']).toBe('#microsoft.graph.notificationMessageTemplate');
      expect(body.displayName).toBe('Device Non-Compliance');
      expect(body.brandingOptions).toBe('includeCompanyName');
      expect(result.content[0].text).toContain('tmpl1');
    });
  });

  describe('get_notification_template', () => {
    it('fetches template and localized messages', async () => {
      graph.get.mockResolvedValue({ id: 'tmpl1' });
      graph.getAll.mockResolvedValue([{ locale: 'en-US', subject: 'Subject' }]);
      const result = await server.call('get_notification_template', { templateId: 'tmpl1' });
      expect(graph.get).toHaveBeenCalledWith('/deviceManagement/notificationMessageTemplates/tmpl1');
      expect(graph.getAll).toHaveBeenCalledWith(
        '/deviceManagement/notificationMessageTemplates/tmpl1/localizedNotificationMessages'
      );
      const parsed = JSON.parse(result.content[0].text);
      expect(parsed.template).toBeDefined();
      expect(parsed.localizedMessages).toBeDefined();
    });
  });

  describe('add_notification_template_message', () => {
    it('posts localized message via v1.0 with correct body', async () => {
      graph.post.mockResolvedValue({ id: 'msg1', locale: 'de-DE' });
      await server.call('add_notification_template_message', {
        templateId: 'tmpl1',
        locale: 'de-DE',
        subject: 'Gerät nicht konform',
        messageTemplate: 'Ihr Gerät erfüllt nicht die Anforderungen.',
        isDefault: false,
      });
      const [url, body] = args(graph.post);
      expect(url).toBe('/deviceManagement/notificationMessageTemplates/tmpl1/localizedNotificationMessages');
      expect(body['@odata.type']).toBe('#microsoft.graph.localizedNotificationMessage');
      expect(body.locale).toBe('de-DE');
      expect(body.isDefault).toBe(false);
    });
  });

  describe('send_notification_template_test', () => {
    it('posts to sendTestMessage action via v1.0', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('send_notification_template_test', { templateId: 'tmpl1' });
      const [url] = args(graph.post);
      expect(url).toBe('/deviceManagement/notificationMessageTemplates/tmpl1/sendTestMessage');
    });
  });

  // ── Bug-fix regression tests ──────────────────────────────────────────────

  describe('collect_device_diagnostics', () => {
    it('sends templateType as a complex object with @odata.type', async () => {
      graph.beta.post.mockResolvedValue({ id: 'req1' });
      await server.call('collect_device_diagnostics', { deviceId: 'dev1' });
      const [, body] = args(graph.beta.post);
      expect(body.templateType).toEqual({
        '@odata.type': 'microsoft.graph.deviceLogCollectionRequest',
        templateType: 'predefined',
      });
    });

    it("wraps deviceId in single-quoted key-segment and escapes embedded quotes", async () => {
      graph.beta.post.mockResolvedValue({ id: 'req1' });
      await server.call('collect_device_diagnostics', { deviceId: "d'1" });
      const [url] = args(graph.beta.post);
      expect(url).toBe(
        "/deviceManagement/managedDevices('d''1')/createDeviceLogCollectionRequest"
      );
    });
  });

  describe('get_intune_app_install_status', () => {
    it('fetches the full app (no $select) so @odata.type is preserved', async () => {
      graph.get.mockResolvedValue({
        id: 'app1',
        displayName: 'Office',
        '@odata.type': '#microsoft.graph.win32LobApp',
        publishingState: 'published',
      });
      graph.post.mockResolvedValue({ Schema: [], Values: [] });
      graph.beta.get.mockResolvedValue({ installedDeviceCount: 5 });

      await server.call('get_intune_app_install_status', { appId: 'app1' });

      const [appUrl, appParams] = args(graph.get);
      expect(appUrl).toBe('/deviceAppManagement/mobileApps/app1');
      expect(appParams).toBeUndefined();
    });

    it('skips the installSummary call when @odata.type is missing', async () => {
      graph.get.mockResolvedValue({ id: 'app1', displayName: 'Foo' }); // no @odata.type
      graph.post.mockResolvedValue({ Schema: [], Values: [] });

      await server.call('get_intune_app_install_status', { appId: 'app1' });

      expect(graph.beta.get).not.toHaveBeenCalled();
    });
  });

  // ── Win32 LOB upload: detection rules at create time (issue #3) ──────────

  describe('upload_win32_lob_app', () => {
    // Tracks tmp .intunewin files we built so we can clean up after each case.
    let tmpFiles: string[] = [];
    afterEach(() => {
      for (const p of tmpFiles) { try { fs.unlinkSync(p); } catch { /* ignore */ } }
      tmpFiles = [];
    });
    const makeFixture = (xml: string): string => {
      const p = makeIntuneWin(xml);
      tmpFiles.push(p);
      return p;
    };

    // We bail out of the upload chain by failing the second `graph.post`
    // (the contentVersions POST inside uploadAppContentVersion). The first
    // POST captures the create body — which is the thing under test —
    // before the failure, and the handler's catch block cleans up the
    // shell app via `graph.delete`. This avoids mocking the full
    // blob-upload + polling chain just to inspect the create body.
    const armBailOut = () => {
      graph.post
        .mockResolvedValueOnce({ id: 'app-new' })
        .mockRejectedValueOnce(new Error('test bail-out before upload'));
      graph.delete.mockResolvedValueOnce(undefined);
    };

    const baseArgs = (filePath: string) => ({
      filePath,
      displayName: 'Test App',
      publisher: 'Test Corp',
      installCommandLine: 'msiexec /i app.msi /qn',
      uninstallCommandLine: 'msiexec /x {12345678-1234-1234-1234-123456789012} /qn',
      setupFilePath: 'app.msi',
    });

    it('auto-injects a ProductCode detection rule for MSI packages when rules omitted', async () => {
      armBailOut();
      const fixture = makeFixture(MSI_DETECTION_XML);

      await expect(server.call('upload_win32_lob_app', baseArgs(fixture)))
        .rejects.toThrow('test bail-out before upload');

      const [url, body] = args(graph.post, 0);
      expect(url).toBe('/deviceAppManagement/mobileApps');
      expect(body['@odata.type']).toBe('#microsoft.graph.win32LobApp');
      expect(body.rules).toEqual([{
        '@odata.type': '#microsoft.graph.win32LobAppProductCodeRule',
        ruleType: 'detection',
        productCode: '{12345678-1234-1234-1234-123456789012}',
        productVersionOperator: 'notConfigured',
        productVersion: null,
      }]);
      expect(body.msiInformation).toMatchObject({
        productCode: '{12345678-1234-1234-1234-123456789012}',
        productVersion: '1.2.3',
        upgradeCode: '{BBBBBBBB-BBBB-BBBB-BBBB-BBBBBBBBBBBB}',
        requiresReboot: false,
        packageType: 'perMachine',
        publisher: 'Test Corp',
      });
    });

    it('uses caller-supplied rules over the MSI auto-rule', async () => {
      armBailOut();
      const fixture = makeFixture(MSI_DETECTION_XML);

      const callerRule = {
        '@odata.type': '#microsoft.graph.win32LobAppFileSystemRule',
        ruleType: 'detection' as const,
        path: 'C:\\Program Files\\MyApp',
        fileOrFolderName: 'app.exe',
        operationType: 'exists',
        operator: 'notConfigured',
        check32BitOn64System: false,
      };

      await expect(server.call('upload_win32_lob_app', {
        ...baseArgs(fixture),
        rules: [callerRule],
      })).rejects.toThrow('test bail-out before upload');

      const [, body] = args(graph.post, 0);
      expect(body.rules).toHaveLength(1);
      expect(body.rules[0]['@odata.type']).toBe('#microsoft.graph.win32LobAppFileSystemRule');
      expect(body.rules[0].path).toBe('C:\\Program Files\\MyApp');
    });

    it('fails fast for EXE packages when no rules are provided', async () => {
      // No bail-out needed — handler throws before any POST.
      const fixture = makeFixture(EXE_DETECTION_XML);

      await expect(server.call('upload_win32_lob_app', baseArgs(fixture)))
        .rejects.toThrow(/No detection rules provided and package is not an MSI/);

      expect(graph.post).not.toHaveBeenCalled();
    });

    it('accepts EXE packages when caller supplies rules', async () => {
      armBailOut();
      const fixture = makeFixture(EXE_DETECTION_XML);

      const callerRule = {
        '@odata.type': '#microsoft.graph.win32LobAppRegistryRule',
        ruleType: 'detection' as const,
        keyPath: 'HKLM\\SOFTWARE\\TestCorp\\MyApp',
        valueName: 'Version',
        operationType: 'string',
        operator: 'equal',
        comparisonValue: '1.2.3',
        check32BitOn64System: false,
      };

      await expect(server.call('upload_win32_lob_app', {
        ...baseArgs(fixture),
        rules: [callerRule],
      })).rejects.toThrow('test bail-out before upload');

      const [, body] = args(graph.post, 0);
      expect(body.rules).toHaveLength(1);
      expect(body.rules[0]['@odata.type']).toBe('#microsoft.graph.win32LobAppRegistryRule');
      // EXE packages have no MsiInfo block — msiInformation must not be set
      // to an empty/stub object that would overwrite anything in Intune.
      expect(body.msiInformation).toBeUndefined();
    });

    it('passes caller returnCodes, displayVersion, roleScopeTagIds and maxRunTimeInMinutes through', async () => {
      armBailOut();
      const fixture = makeFixture(MSI_DETECTION_XML);

      await expect(server.call('upload_win32_lob_app', {
        ...baseArgs(fixture),
        displayVersion: '1.2.3',
        roleScopeTagIds: ['tag-1', 'tag-2'],
        maxRunTimeInMinutes: 90,
        returnCodes: [
          { returnCode: 0, type: 'success' as const },
          { returnCode: 42, type: 'retry' as const },
        ],
      })).rejects.toThrow('test bail-out before upload');

      const [, body] = args(graph.post, 0);
      expect(body.displayVersion).toBe('1.2.3');
      expect(body.roleScopeTagIds).toEqual(['tag-1', 'tag-2']);
      expect(body.installExperience.maxRunTimeInMinutes).toBe(90);
      expect(body.returnCodes).toHaveLength(2);
      expect(body.returnCodes[1]).toEqual({ returnCode: 42, type: 'retry' });
    });
  });

  describe('URL-encoding of opaque ids', () => {
    it('encodes appId in mobileApps path', async () => {
      graph.get.mockResolvedValue({ id: 'a/1' });
      await server.call('get_intune_app', { appId: 'a/1' });
      const [url] = args(graph.get);
      expect(url).toBe('/deviceAppManagement/mobileApps/a%2F1');
    });

    it('encodes deviceId in managedDevices actions', async () => {
      graph.post.mockResolvedValue(undefined);
      await server.call('sync_managed_device', { deviceId: 'd/1' });
      expect(graph.post).toHaveBeenCalledWith(
        '/deviceManagement/managedDevices/d%2F1/syncDevice', {},
      );
    });

    it('encodes policyId in deviceCompliancePolicies path', async () => {
      graph.get.mockResolvedValue({ id: 'p/1' });
      await server.call('get_compliance_policy', { policyId: 'p/1' });
      expect(graph.get).toHaveBeenCalledWith('/deviceManagement/deviceCompliancePolicies/p%2F1');
    });

    it('encodes configId in deviceConfigurations path', async () => {
      graph.get.mockResolvedValue({ id: 'c/1' });
      await server.call('get_device_configuration', { configId: 'c/1' });
      expect(graph.get).toHaveBeenCalledWith('/deviceManagement/deviceConfigurations/c%2F1');
    });
  });
});
