import { RateLimiter } from '../../src/http/rateLimiter';

describe('RateLimiter', () => {
  it('allows up to max within a window, then blocks', () => {
    const rl = new RateLimiter({ windowMs: 1000, max: 3, maxKeys: 100 });
    const now = 1_000_000;
    expect(rl.take('a', now).allowed).toBe(true);
    expect(rl.take('a', now).allowed).toBe(true);
    expect(rl.take('a', now).allowed).toBe(true);
    const blocked = rl.take('a', now);
    expect(blocked.allowed).toBe(false);
    expect(blocked.retryAfterSec).toBeGreaterThan(0);
  });

  it('resets after the window elapses', () => {
    const rl = new RateLimiter({ windowMs: 1000, max: 1, maxKeys: 100 });
    expect(rl.take('a', 0).allowed).toBe(true);
    expect(rl.take('a', 500).allowed).toBe(false);
    expect(rl.take('a', 1000).allowed).toBe(true); // new window (resetAt <= now)
  });

  it('tracks keys independently', () => {
    const rl = new RateLimiter({ windowMs: 1000, max: 1, maxKeys: 100 });
    expect(rl.take('a', 0).allowed).toBe(true);
    expect(rl.take('b', 0).allowed).toBe(true);
    expect(rl.take('a', 0).allowed).toBe(false);
  });

  it('bounds memory when maxKeys is reached', () => {
    const rl = new RateLimiter({ windowMs: 1000, max: 5, maxKeys: 2 });
    rl.take('a', 0);
    rl.take('b', 0);
    // third distinct key at capacity with all buckets live → table cleared, still allowed
    expect(rl.take('c', 0).allowed).toBe(true);
  });
});
