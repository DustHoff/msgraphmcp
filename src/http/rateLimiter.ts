// Lightweight in-process fixed-window rate limiter (audit DOS-1).
//
// Bounds how many requests a single client key (IP) may make per window, so an
// unauthenticated flood cannot exhaust the session pool or spam the auth check.
// This is a single-replica, best-effort control — distributed floods still need
// an edge/WAF limit. No external dependency, predictable memory via maxKeys.

export interface RateLimitConfig {
  /** Window length in milliseconds. */
  windowMs: number;
  /** Max requests permitted per key per window. */
  max: number;
  /** Upper bound on distinct keys tracked (memory guard). */
  maxKeys: number;
}

export interface RateLimitResult {
  allowed: boolean;
  /** Seconds until the window resets (for the Retry-After header). */
  retryAfterSec: number;
}

export class RateLimiter {
  private readonly windowMs: number;
  private readonly max: number;
  private readonly maxKeys: number;
  private readonly buckets = new Map<string, { count: number; resetAt: number }>();

  constructor(cfg: RateLimitConfig) {
    this.windowMs = cfg.windowMs;
    this.max = cfg.max;
    this.maxKeys = cfg.maxKeys;
  }

  /** Records a hit for `key` at time `now` (ms) and reports whether it is allowed. */
  take(key: string, now: number): RateLimitResult {
    let bucket = this.buckets.get(key);

    if (!bucket || bucket.resetAt <= now) {
      // Opportunistic cleanup before tracking a new key, bounding memory.
      if (!bucket && this.buckets.size >= this.maxKeys) this.evictExpired(now);
      bucket = { count: 0, resetAt: now + this.windowMs };
      this.buckets.set(key, bucket);
    }

    bucket.count++;
    if (bucket.count > this.max) {
      return { allowed: false, retryAfterSec: Math.max(1, Math.ceil((bucket.resetAt - now) / 1000)) };
    }
    return { allowed: true, retryAfterSec: 0 };
  }

  private evictExpired(now: number): void {
    for (const [key, bucket] of this.buckets) {
      if (bucket.resetAt <= now) this.buckets.delete(key);
    }
    // Hard safety bound: if everything is still live and we are at capacity,
    // drop the whole table rather than grow unbounded. Worst case this resets
    // counters for active clients — acceptable for a flood-protection control.
    if (this.buckets.size >= this.maxKeys) this.buckets.clear();
  }
}
