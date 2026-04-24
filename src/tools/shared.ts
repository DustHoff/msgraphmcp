// Shared helpers used by tool modules.

/**
 * Microsoft Graph path prefix for a user.
 *
 * The literal string "me" is NOT a valid user id or userPrincipalName —
 * `/users/me` returns 404. For the signed-in user the Graph API uses the
 * dedicated `/me` endpoint. This helper routes the `'me'` sentinel to
 * `/me` and encodes any other id/UPN for safe URL inclusion.
 *
 * @example userPath('me')                    → '/me'
 * @example userPath('alice@contoso.com')     → '/users/alice%40contoso.com'
 */
export function userPath(userId: string): string {
  return userId === 'me' ? '/me' : `/users/${encodeURIComponent(userId)}`;
}

/**
 * Escapes a string for safe inclusion inside an OData single-quoted string literal.
 * Per the OData spec single quotes are escaped by doubling: `it's` → `it''s`.
 * Call this on any user-supplied value inserted between quotes in an OData
 * expression such as `search(q='...')` or `$filter=displayName eq '...'`.
 */
export function odataQuote(value: string): string {
  return value.replace(/'/g, "''");
}

/**
 * Returns true when a set of OData query parameters requires the
 * `ConsistencyLevel: eventual` request header on directory-object collections
 * (`/users`, `/groups`, `/directoryObjects`, etc).
 *
 * Required for: `$search`, `$count=true`, and advanced `$filter` usages like
 * `endsWith`, `not`, `$orderby` + `$filter`.
 * See https://learn.microsoft.com/en-us/graph/aad-advanced-queries
 */
export function needsEventualConsistency(params: Record<string, unknown>): boolean {
  return '$search' in params || params['$count'] === true;
}

/**
 * Percent-encode an opaque entity id for safe inclusion as a single path
 * segment in a Graph API URL. Prevents callers from breaking out of the
 * intended path with characters like `/`, `?`, `#`, or whitespace — any of
 * which would change the target resource or let a tool argument smuggle
 * additional URL parts into the Graph request.
 *
 * Use this for GUIDs, object ids, directory-entity ids, device ids, etc.
 * For OneDrive/SharePoint paths that must keep internal `/` separators, use
 * a dedicated path encoder instead.
 */
export function encodeId(id: string): string {
  return encodeURIComponent(id);
}

/**
 * Encode a OneDrive/SharePoint path for inclusion between `root:` and `:`
 * markers. The path is encoded segment-by-segment so that `/` separators
 * stay intact while spaces, `#`, `?`, `%`, etc. are percent-encoded.
 * Ensures the path is absolute (leading `/`).
 */
export function encodeDrivePath(path: string): string {
  const normalized = path.startsWith('/') ? path : `/${path}`;
  return normalized.split('/').map(encodeURIComponent).join('/');
}

/**
 * Escapes a value for inclusion inside an HTML text node. Encodes the
 * five HTML-significant characters: & < > " '. Use this when reflecting
 * untrusted data (e.g. OAuth `error_description` query params) into a
 * response body so a crafted value cannot break out of the surrounding
 * markup or re-introduce a tag via HTML numeric entities.
 */
export function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
