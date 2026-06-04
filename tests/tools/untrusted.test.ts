import { wrapUntrustedText, annotateToolResult, guardToolOutput } from '../../src/tools/untrusted';

describe('wrapUntrustedText', () => {
  it('wraps text in an untrusted-data envelope', () => {
    const out = wrapUntrustedText('hello');
    expect(out).toContain('UNTRUSTED TOOL OUTPUT');
    expect(out).toContain('hello');
    expect(out).toContain('END UNTRUSTED TOOL OUTPUT');
  });
});

describe('annotateToolResult', () => {
  it('wraps text content items', () => {
    const r = annotateToolResult({ content: [{ type: 'text', text: 'graph-data' }] });
    const text = r.content![0].text as string;
    expect(text).toContain('graph-data');
    expect(text).toContain('UNTRUSTED');
  });

  it('leaves non-text items untouched', () => {
    const img = { content: [{ type: 'image', data: 'xx' }] };
    expect(annotateToolResult(img)).toEqual(img);
  });

  it('passes through shapes without a content array', () => {
    expect(annotateToolResult({} as Record<string, unknown>)).toEqual({});
  });

  it('preserves isError and sibling fields', () => {
    const r = annotateToolResult({ content: [{ type: 'text', text: 'oops' }], isError: true });
    expect(r.isError).toBe(true);
    expect(r.content![0].text as string).toContain('oops');
  });
});

describe('guardToolOutput', () => {
  it('wraps the output of registered tool handlers', async () => {
    let registeredHandler: ((...a: unknown[]) => unknown) | undefined;
    const fakeServer = {
      tool: (...args: unknown[]) => {
        registeredHandler = args[args.length - 1] as (...a: unknown[]) => unknown;
      },
    };
    guardToolOutput(fakeServer as never);
    (fakeServer as { tool: (...a: unknown[]) => unknown }).tool('t', 'desc', {}, async () => ({
      content: [{ type: 'text', text: 'secret-data' }],
    }));

    const result = (await registeredHandler!()) as { content: Array<{ text: string }> };
    expect(result.content[0].text).toContain('secret-data');
    expect(result.content[0].text).toContain('UNTRUSTED');
  });

  it('passes through registrations with no handler function unchanged', () => {
    const calls: unknown[][] = [];
    const fakeServer = { tool: (...args: unknown[]) => calls.push(args) };
    guardToolOutput(fakeServer as never);
    (fakeServer as { tool: (...a: unknown[]) => unknown }).tool('t', 'desc');
    expect(calls[0]).toEqual(['t', 'desc']);
  });
});
