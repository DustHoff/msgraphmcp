describe('logger', () => {
  let stderrSpy: jest.SpyInstance;

  beforeEach(() => {
    stderrSpy = jest.spyOn(process.stderr, 'write').mockImplementation(() => true);
    jest.resetModules();
  });

  afterEach(() => {
    stderrSpy.mockRestore();
    delete process.env.LOG_LEVEL;
  });

  async function loadLogger() {
    const mod = await import('../src/logger');
    return mod.logger;
  }

  it('emits info messages as JSON lines to stderr', async () => {
    const logger = await loadLogger();
    logger.info('test message', { foo: 'bar' });

    expect(stderrSpy).toHaveBeenCalledTimes(1);
    const raw = String((stderrSpy.mock.calls[0] as [string])[0]);
    const parsed = JSON.parse(raw.trim());
    expect(parsed.level).toBe('info');
    expect(parsed.msg).toBe('test message');
    expect(parsed.foo).toBe('bar');
    expect(parsed.ts).toMatch(/^\d{4}-\d{2}-\d{2}T/);
  });

  it('emits error messages', async () => {
    const logger = await loadLogger();
    logger.error('boom', { code: 500 });
    const raw = String((stderrSpy.mock.calls[0] as [string])[0]);
    const parsed = JSON.parse(raw.trim());
    expect(parsed.level).toBe('error');
    expect(parsed.code).toBe(500);
  });

  it('suppresses debug messages when LOG_LEVEL=info (default)', async () => {
    const logger = await loadLogger();
    logger.debug('hidden');
    expect(stderrSpy).not.toHaveBeenCalled();
  });

  it('emits debug messages when LOG_LEVEL=debug', async () => {
    process.env.LOG_LEVEL = 'debug';
    const logger = await loadLogger();
    logger.debug('visible');
    expect(stderrSpy).toHaveBeenCalledTimes(1);
  });

  it('suppresses messages below configured level', async () => {
    process.env.LOG_LEVEL = 'error';
    const logger = await loadLogger();
    logger.info('suppressed');
    logger.warn('also suppressed');
    expect(stderrSpy).not.toHaveBeenCalled();
  });

  it('emits at error level when LOG_LEVEL=error', async () => {
    process.env.LOG_LEVEL = 'error';
    const logger = await loadLogger();
    logger.error('shown');
    expect(stderrSpy).toHaveBeenCalledTimes(1);
  });

  it('writes one JSON entry per call followed by newline', async () => {
    const logger = await loadLogger();
    logger.info('first');
    logger.info('second');
    expect(stderrSpy).toHaveBeenCalledTimes(2);
    const lines = stderrSpy.mock.calls.map((c) => String((c as [string])[0]).trim());
    lines.forEach((line) => expect(() => JSON.parse(line)).not.toThrow());
  });
});
