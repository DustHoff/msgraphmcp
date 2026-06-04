import { redactSensitive } from '../../src/graph/redact';

describe('redactSensitive', () => {
  it('redacts secret/token/password-like keys by name', () => {
    const out = redactSensitive({ clientSecret: 'x', accessToken: 'y', displayName: 'Alice' }) as Record<string, unknown>;
    expect(out.clientSecret).toBe('***REDACTED***');
    expect(out.accessToken).toBe('***REDACTED***');
    expect(out.displayName).toBe('Alice');
  });

  it('redacts the entire fileEncryptionInfo subtree (LOG-2)', () => {
    const out = redactSensitive({
      fileEncryptionInfo: { encryptionKey: 'k', mac: 'm', initializationVector: 'iv', macKey: 'mk' },
    }) as Record<string, unknown>;
    expect(out.fileEncryptionInfo).toBe('***REDACTED***');
  });

  it('redacts top-level encryptionKey / macKey by name', () => {
    const out = redactSensitive({ encryptionKey: 'k', macKey: 'mk' }) as Record<string, unknown>;
    expect(out.encryptionKey).toBe('***REDACTED***');
    expect(out.macKey).toBe('***REDACTED***');
  });

  it('recurses into nested objects and arrays', () => {
    const out = redactSensitive({ a: { password: 'p' }, list: [{ token: 't' }] }) as {
      a: { password: string };
      list: Array<{ token: string }>;
    };
    expect(out.a.password).toBe('***REDACTED***');
    expect(out.list[0].token).toBe('***REDACTED***');
  });

  it('truncates beyond the max depth (LOG-5)', () => {
    let deep: Record<string, unknown> = {};
    let cur = deep;
    for (let i = 0; i < 25; i++) {
      cur.next = {};
      cur = cur.next as Record<string, unknown>;
    }
    expect(JSON.stringify(redactSensitive(deep))).toContain('TRUNCATED');
  });

  it('handles circular references without throwing (LOG-5)', () => {
    const a: Record<string, unknown> = { name: 'x' };
    a.self = a;
    let out: Record<string, unknown> = {};
    expect(() => {
      out = redactSensitive(a) as Record<string, unknown>;
    }).not.toThrow();
    expect(out.self).toBe('***CIRCULAR***');
  });

  it('passes primitives through unchanged', () => {
    expect(redactSensitive('hi')).toBe('hi');
    expect(redactSensitive(42)).toBe(42);
    expect(redactSensitive(null)).toBeNull();
  });
});
