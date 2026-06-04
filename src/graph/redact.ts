// Redaction of sensitive values before they reach debug logs.
//
// GRAPH_DEBUG logs full Graph request/response bodies; this strips credentials,
// tokens, and encryption material first (audit LOG-2/LOG-5).

// Keys whose VALUE must never appear in logs. Case-insensitive, substring match,
// so nested fields like `passwordProfile.password`, `clientSecret`, `accessToken`,
// `encryptionKey`, `macKey` are all caught.
const SENSITIVE_KEY_PATTERN =
  /password|secret|token|credential|private[-_]?key|apikey|api[-_]key|encryptionkey|mackey/i;

// Keys whose ENTIRE subtree must be redacted. The individual nested field names are
// not obviously sensitive (e.g. `fileEncryptionInfo.{mac,initializationVector}`) but
// together they reveal the Intune AES content-encryption material (audit LOG-2).
const REDACT_SUBTREE_PATTERN = /^fileEncryptionInfo$/i;

const REDACTED = '***REDACTED***';

// Bound the recursion so a deeply nested or cyclic body cannot blow the stack or
// amplify logging work (audit LOG-5). `seen` detects cycles along the current path.
const MAX_REDACT_DEPTH = 16;

export function redactSensitive(value: unknown, depth = 0, seen = new WeakSet<object>()): unknown {
  if (value === null || value === undefined) return value;
  if (typeof value !== 'object') return value;
  if (depth >= MAX_REDACT_DEPTH) return '***TRUNCATED***';

  const obj = value as object;
  if (seen.has(obj)) return '***CIRCULAR***';
  seen.add(obj);
  try {
    if (Array.isArray(value)) {
      return value.map((v) => redactSensitive(v, depth + 1, seen));
    }

    const out: Record<string, unknown> = {};
    for (const [key, val] of Object.entries(value as Record<string, unknown>)) {
      if (REDACT_SUBTREE_PATTERN.test(key)) {
        out[key] = REDACTED; // redact the whole subtree (keys + nested values)
      } else if (SENSITIVE_KEY_PATTERN.test(key)) {
        out[key] =
          typeof val === 'string' || typeof val === 'number'
            ? REDACTED
            : redactSensitive(val, depth + 1, seen);
      } else {
        out[key] = redactSensitive(val, depth + 1, seen);
      }
    }
    return out;
  } finally {
    // Backtrack so a sibling that legitimately references the same object isn't
    // mis-flagged as circular — only ancestor cycles are caught.
    seen.delete(obj);
  }
}
