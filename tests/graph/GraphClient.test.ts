import axios from 'axios';
import { GraphClient } from '../../src/graph/GraphClient';
import { TokenManager } from '../../src/auth/TokenManager';

jest.mock('axios');
jest.mock('../../src/auth/TokenManager');

const mockedAxios = jest.mocked(axios);

describe('GraphClient', () => {
  let mockHttp: {
    get: jest.Mock;
    post: jest.Mock;
    patch: jest.Mock;
    put: jest.Mock;
    delete: jest.Mock;
    interceptors: { request: { use: jest.Mock }; response: { use: jest.Mock } };
  };
  let client: GraphClient;
  let requestInterceptor: (config: Record<string, unknown>) => Promise<Record<string, unknown>>;
  let responseErrorInterceptor: (error: unknown) => Promise<unknown>;

  beforeEach(() => {
    mockHttp = {
      get: jest.fn(),
      post: jest.fn(),
      patch: jest.fn(),
      put: jest.fn(),
      delete: jest.fn(),
      interceptors: {
        request: { use: jest.fn() },
        response: { use: jest.fn() },
      },
    };

    (mockedAxios.create as jest.Mock).mockReturnValue(mockHttp);

    const mockTokenManager = new (TokenManager as jest.MockedClass<typeof TokenManager>)();
    (mockTokenManager.getAccessToken as jest.Mock).mockResolvedValue('test-bearer-token');

    client = new GraphClient(mockTokenManager);

    // Capture interceptors to test them
    requestInterceptor = mockHttp.interceptors.request.use.mock.calls[0][0];
    const [, onError] = mockHttp.interceptors.response.use.mock.calls[0];
    responseErrorInterceptor = onError;
  });

  afterEach(() => jest.clearAllMocks());

  describe('request interceptor', () => {
    it('attaches Bearer token to every request', async () => {
      const config: Record<string, unknown> = { headers: {} };
      const result = await requestInterceptor(config);
      expect(result.headers).toMatchObject({ Authorization: 'Bearer test-bearer-token' });
    });
  });

  describe('get()', () => {
    it('calls GET and returns response data', async () => {
      mockHttp.get.mockResolvedValue({ data: { id: '1', name: 'Alice' } });
      const result = await client.get('/users/1');
      expect(mockHttp.get).toHaveBeenCalledWith('/users/1', { params: undefined });
      expect(result).toEqual({ id: '1', name: 'Alice' });
    });

    it('passes query params', async () => {
      mockHttp.get.mockResolvedValue({ data: { value: [] } });
      await client.get('/users', { $top: 10 });
      expect(mockHttp.get).toHaveBeenCalledWith('/users', { params: { $top: 10 } });
    });
  });

  describe('getAll()', () => {
    it('follows @odata.nextLink to collect all pages', async () => {
      mockHttp.get
        .mockResolvedValueOnce({ data: { value: [{ id: '1' }], '@odata.nextLink': 'https://graph.microsoft.com/v1.0/users?$skiptoken=abc' } })
        .mockResolvedValueOnce({ data: { value: [{ id: '2' }] } });

      const result = await client.getAll('/users');
      expect(result).toEqual([{ id: '1' }, { id: '2' }]);
      expect(mockHttp.get).toHaveBeenCalledTimes(2);
    });

    it('returns empty array when value is missing', async () => {
      mockHttp.get.mockResolvedValue({ data: {} });
      const result = await client.getAll('/empty');
      expect(result).toEqual([]);
    });
  });

  describe('post()', () => {
    it('sends POST and returns data', async () => {
      mockHttp.post.mockResolvedValue({ data: { id: 'new-id' } });
      const result = await client.post('/users', { displayName: 'Bob' });
      expect(mockHttp.post).toHaveBeenCalledWith('/users', { displayName: 'Bob' }, undefined);
      expect(result).toEqual({ id: 'new-id' });
    });
  });

  describe('patch()', () => {
    it('sends PATCH and returns data', async () => {
      mockHttp.patch.mockResolvedValue({ data: { id: '1', displayName: 'Updated' } });
      const result = await client.patch('/users/1', { displayName: 'Updated' });
      expect(mockHttp.patch).toHaveBeenCalledWith('/users/1', { displayName: 'Updated' });
      expect(result).toEqual({ id: '1', displayName: 'Updated' });
    });
  });

  describe('delete()', () => {
    it('sends DELETE request', async () => {
      mockHttp.delete.mockResolvedValue({ data: undefined });
      await client.delete('/users/1');
      expect(mockHttp.delete).toHaveBeenCalledWith('/users/1');
    });
  });

  describe('error handling', () => {
    it('surfaces Graph API error message', async () => {
      const graphError = {
        response: {
          status: 400,
          data: { error: { message: 'Invalid request' } },
          config: {},
        },
        config: { headers: {} },
        message: 'Request failed',
      };
      await expect(responseErrorInterceptor(graphError)).rejects.toThrow('Graph API error 400: Invalid request');
    });
  });
});
