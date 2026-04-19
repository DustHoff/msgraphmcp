import { GraphClient } from '../../src/graph/GraphClient';

export type MockGraphClient = {
  [K in keyof GraphClient]: jest.MockedFunction<GraphClient[K]>;
};

export function createMockGraphClient(): MockGraphClient {
  return {
    get: jest.fn(),
    getAll: jest.fn(),
    post: jest.fn(),
    patch: jest.fn(),
    put: jest.fn(),
    delete: jest.fn(),
  } as unknown as MockGraphClient;
}

/** Returns a successful Graph list response envelope */
export function listResponse<T>(items: T[]) {
  return { value: items, '@odata.count': items.length };
}

/**
 * Extract call arguments from a jest.Mock as a plain `any[]` tuple.
 * Avoids TS2339 "property does not exist on type unknown" when
 * inspecting mock.calls in strict-mode test files.
 *
 * @example
 *   const [url, body] = args(graph.post);
 *   expect(body.displayName).toBe('Alice');
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export function args(fn: { mock: { calls: unknown[][] } }, callIndex = 0): any[] {
  return fn.mock.calls[callIndex] as any[];
}
