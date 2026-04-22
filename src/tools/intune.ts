import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { GraphClient } from '../graph/GraphClient';

// ─── shared helpers ──────────────────────────────────────────────────────────

const groupAssignmentSchema = z.object({
  groupId: z.string().describe('Azure AD group object id'),
  intent: z.enum(['available', 'required', 'uninstall', 'availableWithoutEnrollment'])
    .default('required')
    .describe('Assignment intent (apps only)'),
  filterMode: z.enum(['include', 'exclude']).optional(),
  filterId: z.string().optional().describe('Assignment filter id'),
});

function buildAssignTarget(groupId: string) {
  return {
    '@odata.type': '#microsoft.graph.groupAssignmentTarget',
    groupId,
  };
}

// ─── Intune Application Management ──────────────────────────────────────────

export function registerIntuneTools(server: McpServer, graph: GraphClient) {

  // ── Mobile Apps ────────────────────────────────────────────────────────────

  server.tool(
    'list_intune_apps',
    'List Intune managed apps (mobileApps). Supports OData filter/select.',
    {
      filter: z.string().optional().describe(
        "OData filter, e.g. \"isAssigned eq true\" or \"contains(displayName,'Office')\""
      ),
      select: z.string().optional().describe(
        "Comma-separated fields, e.g. 'id,displayName,publisher,appAvailability'"
      ),
      top: z.number().int().min(1).max(999).default(50),
      appType: z.string().optional().describe(
        "Filter by OData type, e.g. '#microsoft.graph.windowsStoreApp', '#microsoft.graph.iosStoreApp', " +
        "'#microsoft.graph.androidStoreApp', '#microsoft.graph.webApp', '#microsoft.graph.win32LobApp'"
      ),
    },
    async ({ filter, select, top, appType }) => {
      const params: Record<string, unknown> = { $top: top };
      const filters: string[] = [];
      if (filter) filters.push(filter);
      if (appType) filters.push(`isof('${appType}')`);
      if (filters.length) params['$filter'] = filters.join(' and ');
      if (select) params['$select'] = select;
      const apps = await graph.getAll('/deviceAppManagement/mobileApps', params);
      return { content: [{ type: 'text', text: JSON.stringify(apps, null, 2) }] };
    }
  );

  server.tool(
    'get_intune_app',
    'Get a specific Intune managed app by id.',
    {
      appId: z.string(),
      select: z.string().optional(),
    },
    async ({ appId, select }) => {
      const app = await graph.get(
        `/deviceAppManagement/mobileApps/${appId}`,
        select ? { $select: select } : undefined
      );
      return { content: [{ type: 'text', text: JSON.stringify(app, null, 2) }] };
    }
  );

  server.tool(
    'create_intune_web_app',
    'Create a Web App in Intune (shortcut URL published to managed devices).',
    {
      displayName: z.string(),
      publisher: z.string(),
      appUrl: z.string().url().describe('Target URL of the web app'),
      description: z.string().optional(),
      useManagedBrowser: z.boolean().default(false),
    },
    async ({ displayName, publisher, appUrl, description, useManagedBrowser }) => {
      const body: Record<string, unknown> = {
        '@odata.type': '#microsoft.graph.webApp',
        displayName,
        publisher,
        appUrl,
        useManagedBrowser,
      };
      if (description) body.description = description;
      const app = await graph.post('/deviceAppManagement/mobileApps', body);
      return { content: [{ type: 'text', text: JSON.stringify(app, null, 2) }] };
    }
  );

  server.tool(
    'create_intune_store_app',
    'Add a store app to Intune (Windows Store, iOS App Store, Google Play).',
    {
      displayName: z.string(),
      publisher: z.string(),
      storeType: z.enum(['windowsStore', 'iosStore', 'androidStore'])
        .describe('App store platform'),
      appStoreUrl: z.string().url().describe('Store URL for the app'),
      description: z.string().optional(),
      bundleId: z.string().optional().describe('Bundle id (iOS) or package name (Android)'),
      minimumSupportedOperatingSystem: z.record(z.boolean()).optional()
        .describe('e.g. {"v10_0": true} for Windows 10, {"v12_0": true} for iOS 12'),
    },
    async ({ displayName, publisher, storeType, appStoreUrl, description, bundleId, minimumSupportedOperatingSystem }) => {
      const typeMap: Record<string, string> = {
        windowsStore: '#microsoft.graph.windowsStoreApp',
        iosStore: '#microsoft.graph.iosStoreApp',
        androidStore: '#microsoft.graph.androidStoreApp',
      };
      const body: Record<string, unknown> = {
        '@odata.type': typeMap[storeType],
        displayName,
        publisher,
        appStoreUrl,
      };
      if (description) body.description = description;
      if (bundleId) body.bundleId = bundleId;
      if (minimumSupportedOperatingSystem) body.minimumSupportedOperatingSystem = minimumSupportedOperatingSystem;

      const app = await graph.post('/deviceAppManagement/mobileApps', body);
      return { content: [{ type: 'text', text: JSON.stringify(app, null, 2) }] };
    }
  );

  server.tool(
    'update_intune_app',
    'Update properties of an existing Intune app.',
    {
      appId: z.string(),
      displayName: z.string().optional(),
      publisher: z.string().optional(),
      description: z.string().optional(),
      isFeatured: z.boolean().optional(),
      privacyInformationUrl: z.string().url().optional(),
      informationUrl: z.string().url().optional(),
      notes: z.string().optional(),
    },
    async ({ appId, ...props }) => {
      const body = Object.fromEntries(Object.entries(props).filter(([, v]) => v !== undefined));
      await graph.patch(`/deviceAppManagement/mobileApps/${appId}`, body);
      return { content: [{ type: 'text', text: `App ${appId} updated.` }] };
    }
  );

  server.tool(
    'delete_intune_app',
    'Delete an Intune managed app.',
    { appId: z.string() },
    async ({ appId }) => {
      await graph.delete(`/deviceAppManagement/mobileApps/${appId}`);
      return { content: [{ type: 'text', text: `App ${appId} deleted.` }] };
    }
  );

  server.tool(
    'list_intune_app_assignments',
    'List group assignments of an Intune app.',
    { appId: z.string() },
    async ({ appId }) => {
      const assignments = await graph.getAll(`/deviceAppManagement/mobileApps/${appId}/assignments`);
      return { content: [{ type: 'text', text: JSON.stringify(assignments, null, 2) }] };
    }
  );

  server.tool(
    'assign_intune_app',
    'Assign an Intune app to Azure AD groups. Replaces all existing assignments.',
    {
      appId: z.string(),
      assignments: z.array(groupAssignmentSchema).min(1),
    },
    async ({ appId, assignments }) => {
      const body = {
        mobileAppAssignments: assignments.map((a) => ({
          '@odata.type': '#microsoft.graph.mobileAppAssignment',
          intent: a.intent,
          target: buildAssignTarget(a.groupId),
          ...(a.filterId ? { settings: { deviceAndAppManagementAssignmentFilterId: a.filterId, deviceAndAppManagementAssignmentFilterType: a.filterMode ?? 'include' } } : {}),
        })),
      };
      await graph.post(`/deviceAppManagement/mobileApps/${appId}/assign`, body);
      return { content: [{ type: 'text', text: `App ${appId} assigned to ${assignments.length} group(s).` }] };
    }
  );

  server.tool(
    'get_intune_app_install_status',
    'Get device/user install status for an Intune app.',
    {
      appId: z.string(),
      top: z.number().int().min(1).max(200).default(25),
    },
    async ({ appId, top }) => {
      const statuses = await graph.get(
        `/deviceAppManagement/mobileApps/${appId}/deviceStatuses`,
        { $top: top }
      );
      return { content: [{ type: 'text', text: JSON.stringify(statuses, null, 2) }] };
    }
  );

  // ── Device Configuration Policies ─────────────────────────────────────────

  server.tool(
    'list_device_configurations',
    'List Intune device configuration profiles (legacy profiles, e.g. Windows 10, iOS, Android).',
    {
      filter: z.string().optional().describe("OData filter, e.g. \"contains(displayName,'VPN')\""),
      select: z.string().optional(),
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ filter, select, top }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      if (select) params['$select'] = select;
      const configs = await graph.getAll('/deviceManagement/deviceConfigurations', params);
      return { content: [{ type: 'text', text: JSON.stringify(configs, null, 2) }] };
    }
  );

  server.tool(
    'get_device_configuration',
    'Get a specific device configuration profile.',
    { configId: z.string() },
    async ({ configId }) => {
      const config = await graph.get(`/deviceManagement/deviceConfigurations/${configId}`);
      return { content: [{ type: 'text', text: JSON.stringify(config, null, 2) }] };
    }
  );

  server.tool(
    'create_device_configuration',
    'Create an Intune device configuration profile. The body must include the @odata.type that identifies the profile type.',
    {
      displayName: z.string(),
      description: z.string().optional(),
      odataType: z.string().describe(
        'Profile type, e.g. "#microsoft.graph.windows10GeneralConfiguration", ' +
        '"#microsoft.graph.iosGeneralDeviceConfiguration", ' +
        '"#microsoft.graph.androidGeneralDeviceConfiguration", ' +
        '"#microsoft.graph.windows10EndpointProtectionConfiguration"'
      ),
      settings: z.record(z.unknown()).describe(
        'Platform-specific settings object. Keys depend on the @odata.type. ' +
        'Example for windows10GeneralConfiguration: {"passwordRequired":true,"passwordMinimumLength":8}'
      ),
    },
    async ({ displayName, description, odataType, settings }) => {
      const body: Record<string, unknown> = {
        '@odata.type': odataType,
        displayName,
        ...settings,
      };
      if (description) body.description = description;
      const config = await graph.post('/deviceManagement/deviceConfigurations', body);
      return { content: [{ type: 'text', text: JSON.stringify(config, null, 2) }] };
    }
  );

  server.tool(
    'update_device_configuration',
    'Update an existing device configuration profile.',
    {
      configId: z.string(),
      displayName: z.string().optional(),
      description: z.string().optional(),
      settings: z.record(z.unknown()).optional().describe('Platform-specific settings to update'),
    },
    async ({ configId, displayName, description, settings }) => {
      const body: Record<string, unknown> = { ...settings };
      if (displayName) body.displayName = displayName;
      if (description) body.description = description;
      const config = await graph.patch(`/deviceManagement/deviceConfigurations/${configId}`, body);
      return { content: [{ type: 'text', text: JSON.stringify(config, null, 2) }] };
    }
  );

  server.tool(
    'delete_device_configuration',
    'Delete a device configuration profile.',
    { configId: z.string() },
    async ({ configId }) => {
      await graph.delete(`/deviceManagement/deviceConfigurations/${configId}`);
      return { content: [{ type: 'text', text: `Configuration ${configId} deleted.` }] };
    }
  );

  server.tool(
    'assign_device_configuration',
    'Assign a device configuration profile to groups. Replaces existing assignments.',
    {
      configId: z.string(),
      assignments: z.array(z.object({
        groupId: z.string(),
        intent: z.enum(['include', 'exclude']).default('include'),
      })).min(1),
    },
    async ({ configId, assignments }) => {
      const body = {
        assignments: assignments.map((a) => ({
          target: a.intent === 'exclude'
            ? { '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget', groupId: a.groupId }
            : buildAssignTarget(a.groupId),
        })),
      };
      await graph.post(`/deviceManagement/deviceConfigurations/${configId}/assign`, body);
      return { content: [{ type: 'text', text: `Configuration ${configId} assigned to ${assignments.length} group(s).` }] };
    }
  );

  server.tool(
    'get_device_configuration_assignments',
    'List group assignments of a device configuration profile.',
    { configId: z.string() },
    async ({ configId }) => {
      const assignments = await graph.getAll(`/deviceManagement/deviceConfigurations/${configId}/assignments`);
      return { content: [{ type: 'text', text: JSON.stringify(assignments, null, 2) }] };
    }
  );

  server.tool(
    'get_device_configuration_device_status',
    'Get per-device deployment status for a configuration profile.',
    {
      configId: z.string(),
      top: z.number().int().min(1).max(200).default(25),
    },
    async ({ configId, top }) => {
      const statuses = await graph.get(
        `/deviceManagement/deviceConfigurations/${configId}/deviceStatuses`,
        { $top: top }
      );
      return { content: [{ type: 'text', text: JSON.stringify(statuses, null, 2) }] };
    }
  );

  // ── Settings Catalog Policies (new-style) ─────────────────────────────────

  server.tool(
    'list_configuration_policies',
    'List Intune Settings Catalog policies (new policy format).',
    {
      filter: z.string().optional(),
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ filter, top }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      const policies = await graph.beta.getAll('/deviceManagement/configurationPolicies', params);
      return { content: [{ type: 'text', text: JSON.stringify(policies, null, 2) }] };
    }
  );

  server.tool(
    'get_configuration_policy',
    'Get a Settings Catalog policy with its settings.',
    { policyId: z.string() },
    async ({ policyId }) => {
      const [policy, settings] = await Promise.all([
        graph.beta.get(`/deviceManagement/configurationPolicies/${policyId}`),
        graph.beta.getAll(`/deviceManagement/configurationPolicies/${policyId}/settings`),
      ]);
      return { content: [{ type: 'text', text: JSON.stringify({ policy, settings }, null, 2) }] };
    }
  );

  server.tool(
    'create_configuration_policy',
    'Create an Intune Settings Catalog policy.',
    {
      name: z.string(),
      description: z.string().optional(),
      platforms: z.enum(['windows10', 'macOS', 'iOS', 'android', 'androidEnterprise', 'linux'])
        .describe('Target platform'),
      technologies: z.string().default('mdm').describe(
        'Comma-separated technologies, e.g. "mdm", "mdm,windows10XManagement"'
      ),
      settings: z.array(z.object({
        id: z.string().describe('Setting definition id from the Settings Catalog'),
        settingInstance: z.record(z.unknown()).describe('Setting value instance object'),
      })).describe('Array of settings from the Settings Catalog'),
    },
    async ({ name, description, platforms, technologies, settings }) => {
      const body: Record<string, unknown> = {
        name,
        platforms,
        technologies,
        settings: settings.map((s, idx) => ({ id: String(idx), settingInstance: s.settingInstance })),
      };
      if (description) body.description = description;
      const policy = await graph.beta.post('/deviceManagement/configurationPolicies', body);
      return { content: [{ type: 'text', text: JSON.stringify(policy, null, 2) }] };
    }
  );

  server.tool(
    'update_configuration_policy',
    'Update a Settings Catalog policy (name/description only — settings require recreation).',
    {
      policyId: z.string(),
      name: z.string().optional(),
      description: z.string().optional(),
    },
    async ({ policyId, name, description }) => {
      const body: Record<string, unknown> = {};
      if (name) body.name = name;
      if (description) body.description = description;
      const policy = await graph.beta.patch(`/deviceManagement/configurationPolicies/${policyId}`, body);
      return { content: [{ type: 'text', text: JSON.stringify(policy, null, 2) }] };
    }
  );

  server.tool(
    'delete_configuration_policy',
    'Delete a Settings Catalog policy.',
    { policyId: z.string() },
    async ({ policyId }) => {
      await graph.beta.delete(`/deviceManagement/configurationPolicies/${policyId}`);
      return { content: [{ type: 'text', text: `Policy ${policyId} deleted.` }] };
    }
  );

  server.tool(
    'assign_configuration_policy',
    'Assign a Settings Catalog policy to groups.',
    {
      policyId: z.string(),
      assignments: z.array(z.object({
        groupId: z.string(),
        intent: z.enum(['include', 'exclude']).default('include'),
      })).min(1),
    },
    async ({ policyId, assignments }) => {
      const body = {
        assignments: assignments.map((a) => ({
          target: a.intent === 'exclude'
            ? { '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget', groupId: a.groupId }
            : buildAssignTarget(a.groupId),
        })),
      };
      await graph.beta.post(`/deviceManagement/configurationPolicies/${policyId}/assign`, body);
      return { content: [{ type: 'text', text: `Policy ${policyId} assigned.` }] };
    }
  );

  server.tool(
    'search_settings_catalog',
    'Search available settings in the Intune Settings Catalog (to find setting definition ids).',
    {
      keyword: z.string().describe('Search term, e.g. "BitLocker", "firewall", "password"'),
      platform: z.enum(['windows10', 'macOS', 'iOS', 'android']).optional(),
      top: z.number().int().min(1).max(100).default(20),
    },
    async ({ keyword, platform, top }) => {
      const params: Record<string, unknown> = {
        $top: top,
        $search: `"${keyword}"`,
      };
      if (platform) params['$filter'] = `platforms has '${platform}'`;
      const settings = await graph.beta.get('/deviceManagement/configurationSettings', params);
      return { content: [{ type: 'text', text: JSON.stringify(settings, null, 2) }] };
    }
  );

  // ── Compliance Policies ────────────────────────────────────────────────────

  server.tool(
    'list_compliance_policies',
    'List Intune device compliance policies.',
    {
      filter: z.string().optional(),
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ filter, top }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      const policies = await graph.getAll('/deviceManagement/deviceCompliancePolicies', params);
      return { content: [{ type: 'text', text: JSON.stringify(policies, null, 2) }] };
    }
  );

  server.tool(
    'get_compliance_policy',
    'Get a specific compliance policy.',
    { policyId: z.string() },
    async ({ policyId }) => {
      const policy = await graph.get(`/deviceManagement/deviceCompliancePolicies/${policyId}`);
      return { content: [{ type: 'text', text: JSON.stringify(policy, null, 2) }] };
    }
  );

  server.tool(
    'create_compliance_policy',
    'Create an Intune device compliance policy.',
    {
      displayName: z.string(),
      description: z.string().optional(),
      odataType: z.string().describe(
        'Platform type, e.g. "#microsoft.graph.windows10CompliancePolicy", ' +
        '"#microsoft.graph.iosCompliancePolicy", ' +
        '"#microsoft.graph.androidCompliancePolicy", ' +
        '"#microsoft.graph.macOSCompliancePolicy"'
      ),
      settings: z.record(z.unknown()).describe(
        'Platform-specific compliance settings. ' +
        'Example for windows10CompliancePolicy: {"passwordRequired":true,"osMinimumVersion":"10.0.19041"}'
      ),
      scheduledActionsForRule: z.array(z.object({
        ruleName: z.string().default('PasswordRequired'),
        scheduledActionConfigurations: z.array(z.object({
          actionType: z.enum(['block', 'retire', 'wipe', 'notification', 'pushNotification']).default('block'),
          gracePeriodHours: z.number().int().min(0).default(0),
        })).default([{ actionType: 'block', gracePeriodHours: 0 }]),
      })).optional().describe('Non-compliance actions (defaults to block immediately)'),
    },
    async ({ displayName, description, odataType, settings, scheduledActionsForRule }) => {
      const body: Record<string, unknown> = {
        '@odata.type': odataType,
        displayName,
        ...settings,
        scheduledActionsForRule: scheduledActionsForRule ?? [{
          ruleName: 'PasswordRequired',
          scheduledActionConfigurations: [{ actionType: 'block', gracePeriodHours: 0 }],
        }],
      };
      if (description) body.description = description;
      const policy = await graph.post('/deviceManagement/deviceCompliancePolicies', body);
      return { content: [{ type: 'text', text: JSON.stringify(policy, null, 2) }] };
    }
  );

  server.tool(
    'update_compliance_policy',
    'Update a compliance policy.',
    {
      policyId: z.string(),
      displayName: z.string().optional(),
      description: z.string().optional(),
      settings: z.record(z.unknown()).optional(),
    },
    async ({ policyId, displayName, description, settings }) => {
      const body: Record<string, unknown> = { ...settings };
      if (displayName) body.displayName = displayName;
      if (description) body.description = description;
      const policy = await graph.patch(`/deviceManagement/deviceCompliancePolicies/${policyId}`, body);
      return { content: [{ type: 'text', text: JSON.stringify(policy, null, 2) }] };
    }
  );

  server.tool(
    'delete_compliance_policy',
    'Delete a compliance policy.',
    { policyId: z.string() },
    async ({ policyId }) => {
      await graph.delete(`/deviceManagement/deviceCompliancePolicies/${policyId}`);
      return { content: [{ type: 'text', text: `Compliance policy ${policyId} deleted.` }] };
    }
  );

  server.tool(
    'assign_compliance_policy',
    'Assign a compliance policy to groups.',
    {
      policyId: z.string(),
      assignments: z.array(z.object({
        groupId: z.string(),
        intent: z.enum(['include', 'exclude']).default('include'),
      })).min(1),
    },
    async ({ policyId, assignments }) => {
      const body = {
        assignments: assignments.map((a) => ({
          target: a.intent === 'exclude'
            ? { '@odata.type': '#microsoft.graph.exclusionGroupAssignmentTarget', groupId: a.groupId }
            : buildAssignTarget(a.groupId),
        })),
      };
      await graph.post(`/deviceManagement/deviceCompliancePolicies/${policyId}/assign`, body);
      return { content: [{ type: 'text', text: `Compliance policy ${policyId} assigned.` }] };
    }
  );

  server.tool(
    'get_compliance_policy_device_status',
    'Get per-device compliance status for a policy.',
    {
      policyId: z.string(),
      top: z.number().int().min(1).max(200).default(25),
    },
    async ({ policyId, top }) => {
      const statuses = await graph.get(
        `/deviceManagement/deviceCompliancePolicies/${policyId}/deviceStatuses`,
        { $top: top }
      );
      return { content: [{ type: 'text', text: JSON.stringify(statuses, null, 2) }] };
    }
  );

  // ── Managed Devices ────────────────────────────────────────────────────────

  server.tool(
    'list_managed_devices',
    'List devices enrolled in Intune.',
    {
      filter: z.string().optional().describe(
        "OData filter, e.g. \"operatingSystem eq 'Windows'\", \"complianceState eq 'compliant'\""
      ),
      select: z.string().optional().describe(
        "Fields to return, e.g. 'id,deviceName,operatingSystem,complianceState,lastSyncDateTime'"
      ),
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ filter, select, top }) => {
      const params: Record<string, unknown> = { $top: top };
      if (filter) params['$filter'] = filter;
      if (select) params['$select'] = select;
      const devices = await graph.get('/deviceManagement/managedDevices', params);
      return { content: [{ type: 'text', text: JSON.stringify(devices, null, 2) }] };
    }
  );

  server.tool(
    'get_managed_device',
    'Get details of a specific managed device.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      const device = await graph.get(`/deviceManagement/managedDevices/${deviceId}`);
      return { content: [{ type: 'text', text: JSON.stringify(device, null, 2) }] };
    }
  );

  server.tool(
    'sync_managed_device',
    'Trigger an immediate Intune policy sync on a managed device.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/syncDevice`, {});
      return { content: [{ type: 'text', text: `Sync triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'restart_managed_device',
    'Reboot a managed Windows device.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/rebootNow`, {});
      return { content: [{ type: 'text', text: `Reboot triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'retire_managed_device',
    'Retire (unenroll) a managed device from Intune. Removes corporate data but keeps personal data.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/retire`, {});
      return { content: [{ type: 'text', text: `Device ${deviceId} retired.` }] };
    }
  );

  server.tool(
    'wipe_managed_device',
    'Factory-reset a managed device. ALL data will be erased. Use with caution.',
    {
      deviceId: z.string(),
      keepEnrollmentData: z.boolean().default(false)
        .describe('Keep Intune enrollment data after wipe (re-enrolls automatically)'),
      keepUserData: z.boolean().default(false)
        .describe('Keep user data on the device (supported on some platforms)'),
    },
    async ({ deviceId, keepEnrollmentData, keepUserData }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/wipe`, {
        keepEnrollmentData,
        keepUserData,
      });
      return { content: [{ type: 'text', text: `Wipe initiated for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'shutdown_managed_device',
    'Shut down a managed Windows device.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/shutDown`, {});
      return { content: [{ type: 'text', text: `Shutdown triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'lock_managed_device',
    'Remotely lock a managed device. The device will require a PIN/password to unlock.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/remoteLock`, {});
      return { content: [{ type: 'text', text: `Remote lock triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'set_managed_device_name',
    'Rename a managed Windows device. The new name must be ≤15 characters and contain only letters, numbers, and hyphens.',
    {
      deviceId: z.string(),
      deviceName: z.string().max(15).describe('New device name (max 15 chars, letters/numbers/hyphens)'),
    },
    async ({ deviceId, deviceName }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/setDeviceName`, { deviceName });
      return { content: [{ type: 'text', text: `Device rename to "${deviceName}" triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'windows_defender_scan',
    'Trigger a Windows Defender antivirus scan on a managed Windows device.',
    {
      deviceId: z.string(),
      quickScan: z.boolean().default(true).describe('true = quick scan, false = full scan (takes longer)'),
    },
    async ({ deviceId, quickScan }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/windowsDefenderScan`, { quickScan });
      return { content: [{ type: 'text', text: `Windows Defender ${quickScan ? 'quick' : 'full'} scan triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'windows_defender_update_signatures',
    'Force a Windows Defender signature/definition update on a managed Windows device.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/windowsDefenderUpdateSignatures`, {});
      return { content: [{ type: 'text', text: `Defender signature update triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'rotate_bitlocker_keys',
    'Rotate the BitLocker recovery key for a managed Windows device. The new key will be escrowed to Entra ID / Intune.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/rotateBitLockerKeys`, {});
      return { content: [{ type: 'text', text: `BitLocker key rotation triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'rotate_local_admin_password',
    'Rotate the local administrator password (LAPS) for a managed Windows device. Requires Windows LAPS configured in Intune.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/rotateLocalAdminPassword`, {});
      return { content: [{ type: 'text', text: `Local admin password rotation triggered for device ${deviceId}.` }] };
    }
  );

  server.tool(
    'send_device_notification',
    'Send a custom push notification to the Company Portal app on a managed device.',
    {
      deviceId: z.string(),
      notificationTitle: z.string().describe('Notification title'),
      notificationBody: z.string().describe('Notification body text'),
    },
    async ({ deviceId, notificationTitle, notificationBody }) => {
      await graph.post(
        `/deviceManagement/managedDevices/${deviceId}/sendCustomNotificationToCompanyPortal`,
        { notificationTitle, notificationBody }
      );
      return { content: [{ type: 'text', text: `Notification sent to device ${deviceId}.` }] };
    }
  );

  server.tool(
    'disable_managed_device',
    'Disable a managed device in Intune. The device will lose access to corporate resources but remains enrolled.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/disable`, {});
      return { content: [{ type: 'text', text: `Device ${deviceId} disabled.` }] };
    }
  );

  server.tool(
    'reenable_managed_device',
    'Re-enable a previously disabled managed device in Intune.',
    { deviceId: z.string() },
    async ({ deviceId }) => {
      await graph.post(`/deviceManagement/managedDevices/${deviceId}/reenable`, {});
      return { content: [{ type: 'text', text: `Device ${deviceId} re-enabled.` }] };
    }
  );

  server.tool(
    'trigger_proactive_remediation',
    'Trigger on-demand proactive remediation on a managed Windows device for a specific remediation script.',
    {
      deviceId: z.string(),
      scriptId: z.string().describe('Proactive remediation script ID'),
    },
    async ({ deviceId, scriptId }) => {
      await graph.post(
        `/deviceManagement/managedDevices/${deviceId}/initiateOnDemandProactiveRemediation`,
        { scriptPolicyId: scriptId }
      );
      return { content: [{ type: 'text', text: `Proactive remediation triggered on device ${deviceId} for script ${scriptId}.` }] };
    }
  );

  server.tool(
    'collect_device_diagnostics',
    'Trigger the "Collect diagnostics" remote action on an Intune-managed device (beta API). Returns a log collection request ID — use list_device_diagnostics to poll status and download_device_diagnostics to get the ZIP download URL once completed.',
    { deviceId: z.string().describe('Intune managed device ID') },
    async ({ deviceId }) => {
      const result = await graph.beta.post(
        `/deviceManagement/managedDevices('${deviceId}')/createDeviceLogCollectionRequest`,
        {
          templateType: {
            '@odata.type': '#microsoft.graph.deviceLogCollectionRequest',
            templateType: 'predefined',
          },
        }
      );
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.tool(
    'list_device_diagnostics',
    'List log collection requests (diagnostic packages) for a specific Intune-managed device. Each entry contains the status (pending, completed, failed) and can be used with download_device_diagnostics once completed.',
    {
      deviceId: z.string().describe('Intune managed device ID'),
      requestId: z.string().optional().describe('Specific log collection request ID to check status of a single request'),
    },
    async ({ deviceId, requestId }) => {
      const url = requestId
        ? `/deviceManagement/managedDevices('${deviceId}')/logCollectionRequests('${requestId}')`
        : `/deviceManagement/managedDevices('${deviceId}')/logCollectionRequests`;
      const result = await graph.beta.get(url);
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.tool(
    'download_device_diagnostics',
    'Get a time-limited SAS download URL for a completed diagnostic log package. Returns { value: "<url>" } — use the URL to download the ZIP archive directly.',
    {
      deviceId: z.string().describe('Intune managed device ID'),
      requestId: z.string().describe('Log collection request ID from collect_device_diagnostics or list_device_diagnostics'),
    },
    async ({ deviceId, requestId }) => {
      const result = await graph.beta.post(
        `/deviceManagement/managedDevices('${deviceId}')/logCollectionRequests('${requestId}')/createDownloadUrl`,
        {}
      );
      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
    }
  );

  server.tool(
    'get_device_compliance_overview',
    'Get a tenant-wide compliance state overview across all managed devices.',
    {},
    async () => {
      const overview = await graph.beta.get('/deviceManagement/deviceComplianceOverview');
      return { content: [{ type: 'text', text: JSON.stringify(overview, null, 2) }] };
    }
  );

  server.tool(
    'list_intune_app_protection_policies',
    'List app protection policies (MAM policies) for iOS and Android.',
    {
      platform: z.enum(['ios', 'android', 'all']).default('all'),
    },
    async ({ platform }) => {
      const results: unknown[] = [];
      if (platform === 'ios' || platform === 'all') {
        const ios = await graph.getAll('/deviceAppManagement/iosManagedAppProtections');
        results.push(...ios);
      }
      if (platform === 'android' || platform === 'all') {
        const android = await graph.getAll('/deviceAppManagement/androidManagedAppProtections');
        results.push(...android);
      }
      return { content: [{ type: 'text', text: JSON.stringify(results, null, 2) }] };
    }
  );

  // ── Notification Message Templates ────────────────────────────────────────

  server.tool(
    'list_notification_templates',
    'List Intune notification message templates.',
    {
      top: z.number().int().min(1).max(999).default(50),
    },
    async ({ top }) => {
      const templates = await graph.getAll('/deviceManagement/notificationMessageTemplates', { $top: top });
      return { content: [{ type: 'text', text: JSON.stringify(templates, null, 2) }] };
    }
  );

  server.tool(
    'get_notification_template',
    'Get a notification message template by id (including its localized messages).',
    { templateId: z.string() },
    async ({ templateId }) => {
      const [template, messages] = await Promise.all([
        graph.get(`/deviceManagement/notificationMessageTemplates/${templateId}`),
        graph.getAll(`/deviceManagement/notificationMessageTemplates/${templateId}/localizedNotificationMessages`),
      ]);
      return { content: [{ type: 'text', text: JSON.stringify({ template, localizedMessages: messages }, null, 2) }] };
    }
  );

  server.tool(
    'create_notification_template',
    'Create an Intune notification message template.',
    {
      displayName: z.string().describe('Display name of the template'),
      description: z.string().optional(),
      defaultLocale: z.string().default('en-US').describe('Fallback locale, e.g. "en-US"'),
      brandingOptions: z.string().default('none').describe(
        'Comma-separated branding flags, e.g. "includeCompanyLogo,includeCompanyName,includeContactInformation". ' +
        'Possible values: none, includeCompanyLogo, includeCompanyName, includeContactInformation, ' +
        'includeCompanyPortalLink, includeDeviceDetails'
      ),
      roleScopeTagIds: z.array(z.string()).optional().describe('Scope tag ids to assign to this template'),
    },
    async ({ displayName, description, defaultLocale, brandingOptions, roleScopeTagIds }) => {
      const body: Record<string, unknown> = {
        '@odata.type': '#microsoft.graph.notificationMessageTemplate',
        displayName,
        defaultLocale,
      };
      if (brandingOptions && brandingOptions !== 'none') body.brandingOptions = brandingOptions;
      if (description) body.description = description;
      if (roleScopeTagIds) body.roleScopeTagIds = roleScopeTagIds;
      // v1.0 routes to StatelessNotificationFEService (api-version=2023-12-04) which rejects all
      // write ops with 400; beta uses a different internal routing that accepts them.
      const template = await graph.beta.post('/deviceManagement/notificationMessageTemplates', body);
      return { content: [{ type: 'text', text: JSON.stringify(template, null, 2) }] };
    }
  );

  server.tool(
    'update_notification_template',
    'Update an Intune notification message template.',
    {
      templateId: z.string(),
      displayName: z.string().optional(),
      description: z.string().optional(),
      defaultLocale: z.string().optional(),
      brandingOptions: z.string().optional().describe(
        'Comma-separated branding flags, e.g. "includeCompanyLogo,includeCompanyName". ' +
        'Possible values: none, includeCompanyLogo, includeCompanyName, includeContactInformation, ' +
        'includeCompanyPortalLink, includeDeviceDetails'
      ),
      roleScopeTagIds: z.array(z.string()).optional().describe('Scope tag ids to assign to this template'),
    },
    async ({ templateId, displayName, description, defaultLocale, brandingOptions, roleScopeTagIds }) => {
      const body: Record<string, unknown> = {};
      if (displayName) body.displayName = displayName;
      if (description) body.description = description;
      if (defaultLocale) body.defaultLocale = defaultLocale;
      if (brandingOptions) body.brandingOptions = brandingOptions;
      if (roleScopeTagIds) body.roleScopeTagIds = roleScopeTagIds;
      const template = await graph.patch(`/deviceManagement/notificationMessageTemplates/${templateId}`, body);
      return { content: [{ type: 'text', text: JSON.stringify(template, null, 2) }] };
    }
  );

  server.tool(
    'delete_notification_template',
    'Delete an Intune notification message template.',
    { templateId: z.string() },
    async ({ templateId }) => {
      await graph.delete(`/deviceManagement/notificationMessageTemplates/${templateId}`);
      return { content: [{ type: 'text', text: `Notification template ${templateId} deleted.` }] };
    }
  );

  server.tool(
    'add_notification_template_message',
    'Add or update a localized message for a notification template.',
    {
      templateId: z.string(),
      locale: z.string().describe('Locale tag, e.g. "en-US", "de-DE"'),
      subject: z.string().describe('Email subject line'),
      messageTemplate: z.string().describe('Email body (plain text or HTML)'),
      isDefault: z.boolean().default(false).describe('Set as default/fallback locale'),
    },
    async ({ templateId, locale, subject, messageTemplate, isDefault }) => {
      const body = {
        '@odata.type': '#microsoft.graph.localizedNotificationMessage',
        locale,
        subject,
        messageTemplate,
        isDefault,
      };
      const msg = await graph.post(
        `/deviceManagement/notificationMessageTemplates/${templateId}/localizedNotificationMessages`,
        body
      );
      return { content: [{ type: 'text', text: JSON.stringify(msg, null, 2) }] };
    }
  );

  server.tool(
    'send_notification_template_test',
    'Send a test notification email using the template (uses the default locale).',
    { templateId: z.string() },
    async ({ templateId }) => {
      await graph.post(`/deviceManagement/notificationMessageTemplates/${templateId}/sendTestMessage`, {});
      return { content: [{ type: 'text', text: `Test message sent for template ${templateId}.` }] };
    }
  );
}
