const LEVELS = { debug: 0, info: 1, warn: 2, error: 3 } as const;
type Level = keyof typeof LEVELS;

const configured = (process.env.LOG_LEVEL ?? 'info').toLowerCase() as Level;
const minLevel: number = LEVELS[configured] ?? LEVELS.info;

function emit(level: Level, msg: string, data?: Record<string, unknown>): void {
  if (LEVELS[level] < minLevel) return;
  const entry: Record<string, unknown> = {
    ts: new Date().toISOString(),
    level,
    msg,
    ...data,
  };
  process.stderr.write(JSON.stringify(entry) + '\n');
}

export const logger = {
  debug: (msg: string, data?: Record<string, unknown>) => emit('debug', msg, data),
  info:  (msg: string, data?: Record<string, unknown>) => emit('info',  msg, data),
  warn:  (msg: string, data?: Record<string, unknown>) => emit('warn',  msg, data),
  error: (msg: string, data?: Record<string, unknown>) => emit('error', msg, data),
};
