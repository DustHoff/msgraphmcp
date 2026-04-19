# Security Notice

**Reviewed:** 2026-04-19  
**Scope:** source code, dependencies, container, Kubernetes manifests  
**npm audit result:** 0 vulnerabilities (info / low / moderate / high / critical)

---

## Table of Contents

1. [Findings Fixed in This Review](#1-findings-fixed-in-this-review)
2. [Dependency Status](#2-dependency-status)
3. [Accepted Risks](#3-accepted-risks)
4. [Operational Security Guidance](#4-operational-security-guidance)

---

## 1. Findings Fixed in This Review

### 1.1 HTTP request body — no size limit (DoS / OOM) — FIXED

**Severity:** Medium  
**File:** `src/index.ts` — `parseBody()`

The HTTP server (Kubernetes mode) accumulated all request data in memory without
an upper bound.  A malicious or misbehaving client could exhaust the pod's heap
by sending an arbitrarily large body.

**Fix:** `MAX_BODY_BYTES = 4 MB` guard added.  Requests that exceed the limit
are immediately destroyed before any allocation happens.

---

### 1.2 Token cache written with world-readable permissions (0644) — FIXED

**Severity:** Medium  
**File:** `src/auth/TokenManager.ts` — `afterCacheAccess()`

The MSAL token cache (`tokens.json`) was written using Node.js default file
permissions (0644).  Any OS user who can read the file system path could read
the file and obtain the MSAL refresh token.

**Fix:** `fs.writeFileSync(..., { mode: 0o600 })` — only the owning user
(`node`, UID 1000 in the container) may read or write the file.

---

## 2. Dependency Status

### 2.1 Production dependencies

| Package | Installed | Latest | Status | Risk |
|---|---|---|---|---|
| `@azure/msal-node` | 2.16.3 | 5.1.3 | ⚠️ 3 major versions behind | **Medium** — see §2.1.1 |
| `@modelcontextprotocol/sdk` | 1.29.0 | 1.29.0 | ✅ up to date | None |
| `axios` | 1.15.0 | 1.15.0 | ✅ up to date | None |
| `zod` | 3.25.76 | 4.3.6 | ⚠️ 1 major version behind | **Low** — see §2.1.2 |

#### 2.1.1 `@azure/msal-node` 2.x → 5.x — Medium Risk

`@azure/msal-node` v5 is the current Microsoft-supported release.  Version 2.x
is three major releases behind.

**Cannot update without breaking changes:**

- `PublicClientApplication` and `ConfidentialClientApplication` constructor
  signatures changed in v3/v4.
- `ICachePlugin` interface changed in v4 (async hooks now use a different context
  type).
- `acquireTokenByDeviceCode()` callback format changed.
- `SilentFlowRequest` type was refactored.

Full migration requires rewriting `src/auth/TokenManager.ts`.

**Risk assessment:**

- `npm audit` reports 0 CVEs for v2.16.3 as of the review date.
- Microsoft publishes security advisories via the
  [MSRC](https://msrc.microsoft.com/) and the npm advisory database.
- Microsoft historically backports critical security fixes to older major
  versions of MSAL for a limited period; however, v2 is not guaranteed to
  receive future patches indefinitely.
- The device code flow and token caching logic in v2 is mature and well-tested.

**Action:** Track [MSAL Node release notes](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-node/CHANGELOG.md)
and [GitHub Security Advisories](https://github.com/AzureAD/microsoft-authentication-library-for-js/security/advisories).
Plan a migration to v5 as part of the next major development cycle.

#### 2.1.2 `zod` 3.x → 4.x — Low Risk

Zod is used exclusively for MCP tool input validation (all validation runs
server-side before any Graph API call).  It has no network exposure of its own.

**Cannot update without breaking changes:**

- `z.object().strip()` / `.strict()` / `.passthrough()` default behaviour
  changed.
- Error message formats changed.
- Several internal type utilities renamed.

All tool schemas in `src/tools/*.ts` would require review and testing.

**Risk assessment:** Low.  A Zod vulnerability would, at worst, allow a
malformed MCP tool argument to bypass schema validation and reach the Graph API,
which enforces its own server-side validation.  No known CVEs.

---

### 2.2 Development-only dependencies (no production exposure)

| Package | Installed | Latest | Status | Risk |
|---|---|---|---|---|
| `jest` | 29.7.0 | 30.3.0 | ⚠️ 1 major behind | **None** (dev only) |
| `@types/jest` | 29.5.14 | 30.0.0 | ⚠️ 1 major behind | **None** (dev only) |
| `typescript` | 5.9.3 | 6.0.3 | ⚠️ 1 major behind | **None** (compile-time only) |
| `@types/node` | 22.19.17 | 25.6.0 | ⚠️ behind | **None** (type declarations) |
| `ts-jest` | 29.4.9 | 29.4.9 | ✅ up to date | None |
| `tsx` | 4.21.0 | current | ✅ up to date | None |

These packages are excluded from the production Docker image (`npm ci --omit=dev`
in the runtime stage).  They carry no production security risk.

**Note on jest 29.x transitive warnings:** `inflight@1.0.6` and `glob@7.2.3`
are flagged as deprecated by npm.  These are jest-internal dependencies; no
user-facing code relies on them.  They will be resolved when jest 30.x is stable
and `ts-jest` supports it.

---

## 3. Accepted Risks

### 3.1 MCP `/mcp` endpoint — no application-level authentication

**Severity:** Low (mitigated by network controls)

The HTTP server (Kubernetes mode) does not authenticate incoming MCP connections
at the application layer.  Any client that can reach `POST /mcp` can issue Graph
API calls under the authenticated user's identity.

**Mitigations in place:**

- Kubernetes Ingress restricts access to RFC 1918 ranges only
  (`nginx.ingress.kubernetes.io/whitelist-source-range`).
- The pod's `ClusterIP` service is not exposed outside the cluster.
- Kubernetes network policies can further restrict which pods may reach the
  service (not included — cluster-specific).

**Recommendation:** Apply a Kubernetes `NetworkPolicy` to restrict ingress to
the `msgraphmcp` pod to known consumer namespaces only.

---

### 3.2 MCP sessions — no idle timeout

**Severity:** Low

Sessions created in HTTP mode are stored in an in-memory `Map` and are only
removed when the client sends a `DELETE /mcp` request or the server restarts.
Clients that disconnect without sending `DELETE` (e.g., crash, network loss)
leave orphaned session entries.

**Impact:** Memory growth proportional to the number of abandoned sessions.
For a small internal deployment with controlled clients this is acceptable.
Each idle session consumes < 1 KB of overhead.

**Recommendation:** For high-traffic environments, add a session idle-timeout
(e.g., 4 hours) with a background cleanup interval.

---

### 3.3 MSAL refresh token stored in plaintext on disk

**Severity:** Low (mitigated)

MSAL's file-based token cache (`tokens.json`) stores the refresh token in
cleartext JSON.  The refresh token grants Graph API access for up to 90 days
without re-authentication.

**Mitigations in place:**

- Fixed in this review: file is now written with `mode: 0o600` (owner-only).
- In the container the file is owned by UID 1000 (`node` user); the container
  runs as that user with `allowPrivilegeEscalation: false`.
- The PVC is only accessible within the Kubernetes cluster.
- Docker images do not include the token file (`VOLUME ["/data"]` ensures the
  path is always a mount, never baked into the image layer).

**Recommendation:** For high-security environments consider an encrypted
token cache (MSAL `ICachePlugin` can be replaced with an implementation that
encrypts the serialised cache, e.g., using Node.js `crypto.createCipheriv`).

---

### 3.4 Graph API URL paths — user-controlled path segments

**Severity:** Low

Several tools in `src/tools/files.ts` embed user-supplied `filePath` and
`itemPath` values directly in Graph API URL paths (e.g.,
`/drive/root:${filePath}:/content`).  No explicit path normalisation is applied.

**Why the risk is low:**

- The Graph API scopes all `root:/path:` segments to the authenticated user's
  own OneDrive.  Cross-user access via path traversal is structurally impossible
  within the Graph API.
- `userId` is always encoded with `encodeURIComponent()`, preventing header
  injection or path escaping at the user-segment level.
- All parameters pass through Zod `z.string()` schema validation (rejects
  non-string inputs and enforces max length via overall payload size limit).

**Recommendation:** Add a `z.string().regex(/^\/[^<>:"|?*]+$/)` refine to
`filePath`/`itemPath` parameters to reject clearly malformed paths early and
improve error messages for MCP clients.

---

### 3.5 Graph API OData `$filter` injection

**Severity:** Informational

Tools that accept `filter` parameters (e.g., `list_users`, `list_messages`)
pass the value verbatim to the Graph API as `$filter`.  A crafted filter string
cannot escape to another API endpoint or execute code; the worst-case outcome
is reading data the authenticated user already has permission to access (which
is the intended behaviour of the tool).

The Graph API enforces its own OData query validation server-side.

---

## 4. Operational Security Guidance

| Topic | Recommendation |
|---|---|
| **Secrets management** | Store `AZURE_CLIENT_ID`, `AZURE_TENANT_ID`, and `AZURE_CLIENT_SECRET` in Kubernetes Secrets or an external vault (e.g., HashiCorp Vault with ESO). Never commit `.env` files with real values. |
| **Token cache backup** | Do not back up the PVC snapshot with the token cache to untrusted storage. The refresh token is as sensitive as a password. |
| **Scope minimisation** | Set `GRAPH_SCOPES` to only the permissions required for your use case.  Avoid granting `DeviceManagement*` or `Directory.ReadWrite.All` if Intune/directory management is not needed. |
| **Log redaction** | Graph API URLs are logged at `info` level. URLs may contain User Principal Names (UPNs), which are PII in some jurisdictions. Ensure log storage complies with your data classification policy.  Set `LOG_LEVEL=warn` to suppress URL logging in production. |
| **Token rotation** | The MSAL refresh token expires after 90 days (configurable in Azure AD). Automate re-authentication alerts via monitoring on the `authentication failed` log event. |
| **Network policy** | Apply a Kubernetes `NetworkPolicy` to allow ingress to the `msgraphmcp` pod only from authorised namespaces (e.g., the namespace running your MCP clients). |
| **Image pinning** | In production, pin the container image to a specific digest (`ghcr.io/dusthoff/msgraphmcp@sha256:…`) rather than `:latest` to prevent unexpected updates. |
| **Wipe device caution** | The `wipe_managed_device` tool performs an irreversible factory reset. Restrict MCP client access to trusted operators and consider adding a confirmation wrapper prompt. |
