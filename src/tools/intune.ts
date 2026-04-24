import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import * as fs from 'node:fs';
import * as os from 'node:os';
import * as path from 'node:path';
import * as net from 'node:net';
import AdmZip from 'adm-zip';
import axios from 'axios';
import { GraphClient } from '../graph/GraphClient';
import { odataQuote } from './shared';

// Upper bound for server-side downloads of .intunewin packages. Guards against
// an attacker (or misconfigured URL) draining the /tmp filesystem.
const MAX_DOWNLOAD_BYTES = 2 * 1024 * 1024 * 1024; // 2 GB — large LOB apps fit

/**
 * SSRF guard: reject URLs that target the host loopback, link-local metadata
 * services (e.g. 169.254.169.254), or private RFC1918 ranges. Only http/https
 * schemes are allowed — blocks file://, gopher://, ftp://, etc.
 */
function assertSafeDownloadUrl(rawUrl: string): void {
  let url: URL;
  try {
    url = new URL(rawUrl);
  } catch {
    throw new Error(`Invalid fileUrl: ${rawUrl}`);
  }

  if (url.protocol !== 'http:' && url.protocol !== 'https:') {
    throw new Error(`fileUrl scheme not allowed: ${url.protocol} — use http:// or https://`);
  }

  const host = url.hostname;
  const ipFamily = net.isIP(host); // 0 = not an IP literal

  // IP-literal check — block loopback, link-local, and private RFC1918 ranges
  if (ipFamily !== 0) {
    if (isPrivateOrLoopbackIp(host, ipFamily)) {
      throw new Error(`fileUrl host is not allowed (private/loopback IP): ${host}`);
    }
    return;
  }

  // Hostname — reject the obvious internal names. DNS rebinding is not
  // completely mitigated by name checks alone, but blocking these covers the
  // common SSRF vectors without also blocking legitimate cloud hostnames.
  const lower = host.toLowerCase();
  if (lower === 'localhost' || lower.endsWith('.localhost') || lower === 'metadata.google.internal') {
    throw new Error(`fileUrl host is not allowed: ${host}`);
  }
}

function isPrivateOrLoopbackIp(ip: string, family: number): boolean {
  if (family === 4) {
    const parts = ip.split('.').map(p => parseInt(p, 10));
    if (parts.length !== 4 || parts.some(p => Number.isNaN(p))) return false;
    const [a, b] = parts;
    return (
      a === 10 ||                          // 10.0.0.0/8
      (a === 172 && b >= 16 && b <= 31) || // 172.16.0.0/12
      (a === 192 && b === 168) ||          // 192.168.0.0/16
      a === 127 ||                         // 127.0.0.0/8 loopback
      (a === 169 && b === 254) ||          // 169.254.0.0/16 link-local (cloud metadata)
      a === 0 ||                           // 0.0.0.0/8
      a >= 224                             // 224.0.0.0/4 multicast + reserved
    );
  }
  // IPv6: reject loopback, link-local, unique-local, unspecified
  const lower = ip.toLowerCase();
  if (lower === '::1' || lower === '::') return true;
  if (lower.startsWith('fe80:') || lower.startsWith('fc') || lower.startsWith('fd')) return true;
  return false;
}

// ─── Win32 LOB upload helpers ─────────────────────────────────────────────────

interface IntuneWinInfo {
  fileName: string;
  unencryptedSize: number;
  encryptedContent: Buffer;
  encryptionKey: string;
  macKey: string;
  initializationVector: string;
  mac: string;
  fileDigest: string;
  fileDigestAlgorithm: string;
}

function xmlValue(xml: string, tag: string): string {
  return xml.match(new RegExp(`<${tag}[^>]*>([^<]*)<\/${tag}>`))?.[1]?.trim() ?? '';
}

function parseIntuneWin(filePath: string): IntuneWinInfo {
  if (!fs.existsSync(filePath)) throw new Error(`File not found: ${filePath}`);
  const zip = new AdmZip(filePath);

  const detectionEntry = zip.getEntry('IntuneWinPackage/Metadata/Detection.xml');
  if (!detectionEntry) throw new Error('.intunewin missing IntuneWinPackage/Metadata/Detection.xml');
  const xml = detectionEntry.getData().toString('utf8');

  const contentEntry = zip.getEntries().find(
    e => e.entryName.startsWith('IntuneWinPackage/Contents/') && e.entryName.endsWith('.intunewin'),
  );
  if (!contentEntry) throw new Error('.intunewin missing encrypted content in IntuneWinPackage/Contents/');

  return {
    fileName: xmlValue(xml, 'FileName'),
    unencryptedSize: parseInt(xmlValue(xml, 'UnencryptedContentSize'), 10),
    encryptedContent: contentEntry.getData(),
    encryptionKey: xmlValue(xml, 'EncryptionKey'),
    macKey: xmlValue(xml, 'MacKey'),
    initializationVector: xmlValue(xml, 'InitializationVector'),
    mac: xmlValue(xml, 'Mac'),
    fileDigest: xmlValue(xml, 'FileDigest'),
    fileDigestAlgorithm: xmlValue(xml, 'FileDigestAlgorithm') || 'SHA256',
  };
}

async function downloadToTempFile(url: string): Promise<string> {
  assertSafeDownloadUrl(url);
  const tmpPath = path.join(os.tmpdir(), `intunewin-${Date.now()}-${Math.random().toString(36).slice(2)}.intunewin`);
  const response = await axios.get<NodeJS.ReadableStream>(url, {
    responseType: 'stream',
    maxRedirects: 5,
    // Cap the download — axios enforces this after reading the response;
    // the stream-level byte counter below enforces the same cap mid-stream.
    maxContentLength: MAX_DOWNLOAD_BYTES,
  });

  try {
    await new Promise<void>((resolve, reject) => {
      const dest = fs.createWriteStream(tmpPath);
      const stream = response.data as NodeJS.ReadableStream;
      let bytes = 0;
      stream.on('data', (chunk: Buffer) => {
        bytes += chunk.length;
        if (bytes > MAX_DOWNLOAD_BYTES) {
          stream.pause();
          dest.destroy();
          reject(new Error(`Download exceeded ${MAX_DOWNLOAD_BYTES} bytes`));
        }
      });
      stream.on('error', reject);
      dest.on('finish', resolve);
      dest.on('error', reject);
      stream.pipe(dest);
    });
    return tmpPath;
  } catch (err) {
    fs.unlink(tmpPath, () => {});
    throw err;
  }
}

const BLOB_BLOCK_SIZE = 4 * 1024 * 1024; // 4 MB per block

async function uploadBlobBlocks(sasUri: string, data: Buffer): Promise<void> {
  const blockIds: string[] = [];
  const blockCount = Math.ceil(data.length / BLOB_BLOCK_SIZE);

  for (let i = 0; i < blockCount; i++) {
    const blockId = Buffer.from(String(i).padStart(6, '0')).toString('base64');
    blockIds.push(blockId);
    await axios.put(
      `${sasUri}&comp=block&blockid=${encodeURIComponent(blockId)}`,
      data.subarray(i * BLOB_BLOCK_SIZE, Math.min((i + 1) * BLOB_BLOCK_SIZE, data.length)),
      { headers: { 'Content-Type': 'application/octet-stream' }, maxBodyLength: Infinity },
    );
  }

  const blockListXml = `<?xml version="1.0" encoding="utf-8"?><BlockList>${
    blockIds.map(b => `<Latest>${b}</Latest>`).join('')
  }</BlockList>`;
  await axios.put(`${sasUri}&comp=blocklist`, blockListXml, {
    headers: { 'Content-Type': 'application/xml' },
  });
}

async function pollContentFile(
  graph: GraphClient,
  appId: string,
  versionId: string,
  fileId: string,
  until: (state: string) => boolean,
  timeoutMs = 180_000,
): Promise<Record<string, unknown>> {
  const FAIL_STATES = ['Failed', 'TimedOut', 'Error'];
  const url = `/deviceAppManagement/mobileApps/${appId}/microsoft.graph.win32LobApp/contentVersions/${versionId}/files/${fileId}`;
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    await new Promise(r => setTimeout(r, 3000));
    const file = await graph.get<Record<string, unknown>>(url);
    const state = (file.uploadState as string) ?? '';
    if (until(state)) return file;
    if (FAIL_STATES.some(s => state.includes(s))) throw new Error(`Upload failed: uploadState=${state}`);
  }
  throw new Error('Timed out waiting for file upload state');
}

// ─── shared helpers ──────────────────────────────────────────────────────────

// Win32 LOB app rules — field name in Graph API is "rules" (GET and PATCH),
// each entry needs ruleType ("detection"|"requirement") plus type-specific fields.
// Correct @odata.type values use the suffix "Rule", not "Detection".
const win32LobAppRuleSchema = z.object({
  '@odata.type': z.string().describe(
    'Rule type — one of: ' +
    '#microsoft.graph.win32LobAppFileSystemRule | ' +
    '#microsoft.graph.win32LobAppRegistryRule | ' +
    '#microsoft.graph.win32LobAppProductCodeRule | ' +
    '#microsoft.graph.win32LobAppPowerShellScriptRule'
  ),
  ruleType: z.enum(['detection', 'requirement']).default('detection'),
  // file system
  path: z.string().optional().describe('Folder path (file system rule)'),
  fileOrFolderName: z.string().optional().describe('File or folder name (file system rule)'),
  // registry
  keyPath: z.string().optional().describe('Registry key path (registry rule)'),
  valueName: z.string().optional().describe('Registry value name (registry rule)'),
  // shared by file system + registry
  check32BitOn64System: z.boolean().optional(),
  operationType: z.string().optional().describe(
    'e.g. notConfigured | exists | doesNotExist | string | integer | version | sizeInMB'
  ),
  operator: z.string().optional().describe(
    'e.g. notConfigured | equal | notEqual | greaterThan | greaterThanOrEqual | lessThan | lessThanOrEqual'
  ),
  comparisonValue: z.string().optional(),
  // MSI product code
  productCode: z.string().optional().describe('MSI product GUID'),
  productVersionOperator: z.string().optional(),
  productVersion: z.string().optional(),
  // PowerShell script
  enforceSignatureCheck: z.boolean().optional(),
  runAs32Bit: z.boolean().optional(),
  scriptContent: z.string().optional().describe('Base64-encoded PowerShell script content'),
}).passthrough(); // keep any extra fields the caller provides (e.g. detectionType, detectionValue)

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
    'Update properties of an existing Intune app. For Win32 LOB apps use "rules" to replace detection/requirement rules.',
    {
      appId: z.string(),
      displayName: z.string().optional(),
      publisher: z.string().optional(),
      description: z.string().optional(),
      isFeatured: z.boolean().optional(),
      privacyInformationUrl: z.string().url().optional(),
      informationUrl: z.string().url().optional(),
      notes: z.string().optional(),
      rules: z.array(win32LobAppRuleSchema).optional()
        .describe('Detection/requirement rules for Win32 LOB apps (replaces all existing rules). Graph API field name is "rules".'),
    },
    async ({ appId, rules, ...props }) => {
      const body: Record<string, unknown> = Object.fromEntries(
        Object.entries(props).filter(([, v]) => v !== undefined)
      );
      if (rules !== undefined) {
        body.rules = rules;
        // Graph API requires @odata.type on the app body when patching win32LobApp-specific fields
        body['@odata.type'] = '#microsoft.graph.win32LobApp';
      }
      if (Object.keys(body).length === 0) {
        return { content: [{ type: 'text', text: 'No fields provided — nothing to update.' }] };
      }
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
    'upload_win32_lob_app',
    'Upload a Win32 LOB app (.intunewin file) to Intune. ' +
    'Provide either filePath (absolute path on the server, e.g. /data/myapp.intunewin) ' +
    'or fileUrl (HTTP/HTTPS URL from which the server will download the file). ' +
    'After upload, set detection rules via update_intune_app and assign groups via assign_intune_app.',
    {
      filePath: z.string().optional()
        .describe('Absolute path to the .intunewin file on the server, e.g. /data/myapp.intunewin'),
      fileUrl: z.string().url().optional()
        .describe('HTTP/HTTPS URL of the .intunewin file — the server will download it before uploading to Intune'),
      displayName: z.string(),
      publisher: z.string(),
      description: z.string().optional(),
      installCommandLine: z.string()
        .describe('Install command, e.g. "setup.exe /S" or "msiexec /i app.msi /qn"'),
      uninstallCommandLine: z.string()
        .describe('Uninstall command, e.g. "msiexec /x {GUID} /qn"'),
      setupFilePath: z.string()
        .describe('Name of the main installer file inside the package, e.g. "setup.exe" or "app.msi"'),
      applicableArchitectures: z.string().default('x64')
        .describe('Comma-separated architectures, e.g. "x64" or "x86,x64"'),
      minimumSupportedWindowsRelease: z.string().default('1607')
        .describe('Minimum Windows 10/11 feature release, e.g. "1607" or "21H2"'),
      runAsAccount: z.enum(['system', 'user']).default('system'),
      deviceRestartBehavior: z.enum(['allow', 'basedOnReturnCode', 'suppress', 'force']).default('allow'),
    },
    async ({
      filePath, fileUrl, displayName, publisher, description,
      installCommandLine, uninstallCommandLine, setupFilePath,
      applicableArchitectures, minimumSupportedWindowsRelease,
      runAsAccount, deviceRestartBehavior,
    }) => {
      if (!filePath && !fileUrl) throw new Error('Provide either filePath or fileUrl.');
      if (filePath && fileUrl) throw new Error('Provide only one of filePath or fileUrl, not both.');

      // ① Download from URL if needed, then parse .intunewin
      let tmpPath: string | undefined;
      if (fileUrl) {
        tmpPath = await downloadToTempFile(fileUrl);
        filePath = tmpPath;
      }
      const info = parseIntuneWin(filePath!);

      // ② Create Win32LobApp metadata entry
      const app = await graph.post<{ id: string }>('/deviceAppManagement/mobileApps', {
        '@odata.type': '#microsoft.graph.win32LobApp',
        displayName,
        publisher,
        ...(description && { description }),
        fileName: info.fileName,
        installCommandLine,
        uninstallCommandLine,
        setupFilePath,
        applicableArchitectures,
        minimumSupportedWindowsRelease,
        installExperience: {
          '@odata.type': '#microsoft.graph.win32LobAppInstallExperience',
          runAsAccount,
          deviceRestartBehavior,
        },
        returnCodes: [
          { returnCode: 0, type: 'success' },
          { returnCode: 1707, type: 'success' },
          { returnCode: 3010, type: 'softReboot' },
          { returnCode: 1641, type: 'hardReboot' },
          { returnCode: 1618, type: 'retry' },
        ],
        rules: [],
      });

      try {
        // ③ Create content version
        const version = await graph.post<{ id: string }>(
          `/deviceAppManagement/mobileApps/${app.id}/microsoft.graph.win32LobApp/contentVersions`,
          {},
        );

        // ④ Create file entry — triggers Azure Storage URI allocation
        const fileEntry = await graph.post<{ id: string }>(
          `/deviceAppManagement/mobileApps/${app.id}/microsoft.graph.win32LobApp/contentVersions/${version.id}/files`,
          {
            '@odata.type': '#microsoft.graph.mobileAppContentFile',
            name: info.fileName,
            size: info.unencryptedSize,
            sizeEncrypted: info.encryptedContent.length,
            manifest: null,
            isDependency: false,
          },
        );

        // ⑤ Poll until azureStorageUri is ready
        const fileReady = await pollContentFile(
          graph, app.id, version.id, fileEntry.id,
          s => s === 'azureStorageUriRequestSuccess',
        );
        const sasUri = fileReady.azureStorageUri as string;
        if (!sasUri) throw new Error('Azure Storage URI not provided by Graph API');

        // ⑥ Upload encrypted content to Azure Blob in 4 MB blocks
        await uploadBlobBlocks(sasUri, info.encryptedContent);

        // ⑦ Commit file with encryption metadata
        await graph.post(
          `/deviceAppManagement/mobileApps/${app.id}/microsoft.graph.win32LobApp/contentVersions/${version.id}/files/${fileEntry.id}/commit`,
          {
            fileEncryptionInfo: {
              encryptionKey: info.encryptionKey,
              initializationVector: info.initializationVector,
              mac: info.mac,
              macKey: info.macKey,
              profileIdentifier: 'ProfileVersion1',
              fileDigest: info.fileDigest,
              fileDigestAlgorithm: info.fileDigestAlgorithm,
            },
          },
        );

        // ⑧ Poll until commit is confirmed
        await pollContentFile(
          graph, app.id, version.id, fileEntry.id,
          s => s === 'commitFileSuccess',
          300_000, // 5 min — large files take time
        );

        // ⑨ Link committed content version to the app
        await graph.patch(`/deviceAppManagement/mobileApps/${app.id}`, {
          '@odata.type': '#microsoft.graph.win32LobApp',
          committedContentVersion: version.id,
        });

        return {
          content: [{
            type: 'text' as const,
            text: JSON.stringify({
              appId: app.id,
              contentVersionId: version.id,
              fileName: info.fileName,
              message: `"${displayName}" uploaded. Next: set detection rules via update_intune_app, then assign groups via assign_intune_app.`,
            }, null, 2),
          }],
        };
      } catch (err) {
        // Remove the shell app entry so there is no orphaned app without content
        await graph.delete(`/deviceAppManagement/mobileApps/${app.id}`).catch(() => {});
        throw err;
      } finally {
        if (tmpPath) fs.unlink(tmpPath, () => {});
      }
    }
  );

  server.tool(
    'list_intune_app_relationships',
    'List supersedence and dependency relationships of an Intune app.',
    { appId: z.string() },
    async ({ appId }) => {
      const relationships = await graph.getAll(`/deviceAppManagement/mobileApps/${appId}/relationships`);
      return { content: [{ type: 'text', text: JSON.stringify(relationships, null, 2) }] };
    }
  );

  server.tool(
    'set_intune_app_relationships',
    'Set supersedence and/or dependency relationships for an Intune app. Replaces all existing relationships.',
    {
      appId: z.string(),
      relationships: z.array(z.object({
        '@odata.type': z.enum([
          '#microsoft.graph.mobileAppSupersedence',
          '#microsoft.graph.mobileAppDependency',
        ]),
        targetId: z.string().describe('Object id of the target app'),
        supersedenceType: z.enum(['update', 'replace']).optional()
          .describe('Required for supersedence: "update" keeps the old app, "replace" uninstalls it'),
        dependencyType: z.enum(['detect', 'autoInstall']).optional()
          .describe('Required for dependency: "detect" checks presence, "autoInstall" installs automatically'),
      })).min(1),
    },
    async ({ appId, relationships }) => {
      const body = {
        relationships: relationships.map((r) => ({
          '@odata.type': r['@odata.type'],
          targetId: r.targetId,
          ...(r.supersedenceType ? { supersedenceType: r.supersedenceType } : {}),
          ...(r.dependencyType ? { dependencyType: r.dependencyType } : {}),
        })),
      };
      await graph.post(`/deviceAppManagement/mobileApps/${appId}/updateRelationships`, body);
      return { content: [{ type: 'text', text: `App ${appId}: ${relationships.length} relationship(s) set.` }] };
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
    'Get per-device install status for an Intune app (all types: Win32, MSI, MSIX, iOS, Android, macOS, WinGet, webApp, etc.). ' +
    'Uses the Intune Reports API (v1.0 retrieveDeviceAppInstallationStatusReport action) which works for all app types. ' +
    'Response contains deviceInstallStatusReport with Schema + Values (column-based rows) including DeviceName, InstallState, InstallStateDetail, ErrorCode. ' +
    'Optionally returns per-user aggregate statuses. ' +
    'Also attempts to fetch an installSummary (aggregate counts) where available.',
    {
      appId: z.string(),
      top: z.number().int().min(1).max(200).default(25),
      includeUserStatuses: z.boolean().default(false)
        .describe('Also return per-user aggregate install states'),
    },
    async ({ appId, top, includeUserStatuses }) => {
      const app = await graph.get<Record<string, unknown>>(
        `/deviceAppManagement/mobileApps/${appId}`,
        { $select: 'id,displayName,publishingState' }
      );
      const odataType = (app['@odata.type'] as string) ?? '';

      const result: Record<string, unknown> = {
        app: {
          id: app.id,
          displayName: app.displayName,
          odataType,
          publishingState: app.publishingState,
        },
      };

      // Reports API (v1.0) works for all app types: win32, MSIX, iOS, Android, macOS, WinGet, etc.
      // Response shape: { Schema: [{Name, Type}], Values: [[...row values]] }
      result.deviceInstallStatusReport = await graph.post<Record<string, unknown>>(
        '/deviceManagement/reports/microsoft.graph.retrieveDeviceAppInstallationStatusReport',
        {
          filter: `(ApplicationId eq '${odataQuote(appId)}')`,
          select: [
            'DeviceId', 'DeviceName', 'Platform', 'UserName', 'UserPrincipalName',
            'InstallState', 'InstallStateDetail', 'ErrorCode',
            'LastModifiedDateTime', 'AppVersion',
          ],
          skip: 0,
          top,
          orderBy: [],
        }
      );

      if (includeUserStatuses) {
        result.userStatuses = await graph.beta.get(
          `/deviceAppManagement/mobileApps/${appId}/userStatuses`,
          { $top: top }
        ).catch(() => null);
      }

      // installSummary — type-cast beta path; not available for all types, skip silently
      const typeSegment = odataType.replace('#', '');
      try {
        result.installSummary = await graph.beta.get(
          `/deviceAppManagement/mobileApps/${appId}/${typeSegment}/installSummary`
        );
      } catch { /* not available for all app types */ }

      return { content: [{ type: 'text', text: JSON.stringify(result, null, 2) }] };
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
      // OData $search on configurationSettings expects a double-quoted term.
      // Stripping embedded double quotes is safer than trying to escape them —
      // the API has no documented escape sequence for `"` inside $search.
      const safeKeyword = keyword.replace(/"/g, '');
      const params: Record<string, unknown> = {
        $top: top,
        $search: `"${safeKeyword}"`,
      };
      if (platform) params['$filter'] = `platforms has '${odataQuote(platform)}'`;
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
          templateType: { templateType: 'predefined' },
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
