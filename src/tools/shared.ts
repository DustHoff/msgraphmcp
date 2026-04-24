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
