import axios, { AxiosInstance, AxiosRequestConfig, AxiosResponse, InternalAxiosRequestConfig } from 'axios';
import { TokenManager, AuthRequiredError } from '../auth/TokenManager';
import { logger } from '../logger';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_BETA = 'https://graph.microsoft.com/beta';

const DEBUG = process.env.GRAPH_DEBUG !== 'false';

// Relevant MS Graph response headers for diagnostics
const DEBUG_RESPONSE_HEADERS = ['request-id', 'client-request-id', 'x-ms-ags-diagnostic', 'odata-version'];

interface TimedRequestConfig extends InternalAxiosRequestConfig {
  _startMs?: number;
  _retried?: boolean;
  _user?: string;
}

function pickHeaders(headers: Record<string, unknown> | undefined, keys: string[]): Record<string, unknown> {
  if (!headers) return {};
  const result: Record<string, unknown> = {};
  for (const key of keys) {
    if (headers[key] !== undefined) result[key] = headers[key];
  }
  return result;
}

function createAxiosInstance(
  baseURL: string,
  label: string,
  tokenManager: TokenManager,
  getLoginUrl?: () => string,
): AxiosInstance {
  const http = axios.create({ baseURL, maxRedirects: 0 });

  http.interceptors.request.use(async (config: TimedRequestConfig) => {
    let token: string;
    try {
      token = await tokenManager.getAccessToken();
    } catch (err) {
      if (err instanceof AuthRequiredError && getLoginUrl) {
        throw new AuthRequiredError(`Not authenticated — visit ${getLoginUrl()} to sign in with Microsoft`);
      }
      throw err;
    }
    const accountInfo = await tokenManager.getAccountInfo().catch(() => null);
    config.headers.Authorization = `Bearer ${token}`;
    config.headers['Content-Type'] = 'application/json';
    config._startMs = Date.now();
    config._user = accountInfo?.upn ?? 'unknown';

    if (DEBUG) {
      logger.info(`${label} request`, {
        user: config._user,
        method: config.method?.toUpperCase(),
        url: config.url,
        ...(config.params && { params: config.params }),
        ...(config.data !== undefined && { body: config.data }),
      });
    }

    return config;
  });

  http.interceptors.response.use(
    (res: AxiosResponse) => {
      const cfg = res.config as TimedRequestConfig;
      const duration = cfg._startMs !== undefined ? Date.now() - cfg._startMs : undefined;

      logger.info(`${label} ok`, {
        user: cfg._user,
        method: cfg.method?.toUpperCase(),
        url: cfg.url,
        status: res.status,
        ...(duration !== undefined && { durationMs: duration }),
      });

      if (DEBUG) {
        logger.info(`${label} response`, {
          user: cfg._user,
          method: cfg.method?.toUpperCase(),
          url: cfg.url,
          status: res.status,
          headers: pickHeaders(res.headers as Record<string, unknown>, DEBUG_RESPONSE_HEADERS),
          body: res.data,
        });
      }

      return res;
    },
    async (error) => {
      const cfg = error.config as TimedRequestConfig | undefined;
      const duration = cfg?._startMs !== undefined ? Date.now() - cfg._startMs : undefined;
      const status: number | undefined = error.response?.status;
      const msg: string = error.response?.data?.error?.message || error.message;

      if (status === 401 && !cfg?._retried) {
        logger.warn(`${label} 401 — retrying with fresh token`, {
          method: cfg?.method?.toUpperCase(),
          url: cfg?.url,
          ...(duration !== undefined && { durationMs: duration }),
        });
        const token = await tokenManager.getAccessToken();
        error.config.headers.Authorization = `Bearer ${token}`;
        error.config._retried = true;
        return http.request(error.config);
      }

      // Capture actual response URL to detect silent proxy redirects
      const finalUrl: string | undefined = (error.request as { res?: { responseUrl?: string } } | undefined)
        ?.res?.responseUrl;

      logger.error(`${label} error`, {
        user: cfg?._user,
        method: cfg?.method?.toUpperCase(),
        url: cfg?.url,
        ...(finalUrl && finalUrl !== `${baseURL}${cfg?.url}` && { finalUrl }),
        status,
        message: msg,
        ...(duration !== undefined && { durationMs: duration }),
        ...(DEBUG && error.response?.data && { responseBody: error.response.data }),
        ...(DEBUG && error.response?.headers && {
          headers: pickHeaders(error.response.headers as Record<string, unknown>, DEBUG_RESPONSE_HEADERS),
        }),
      });
      throw new Error(`Graph API error ${status}: ${msg}`);
    }
  );

  return http;
}

export class GraphClient {
  private http: AxiosInstance;
  readonly beta: BetaClient;

  constructor(tokenManager: TokenManager, getLoginUrl?: () => string) {
    this.http = createAxiosInstance(GRAPH_BASE, 'graph', tokenManager, getLoginUrl);
    this.beta = new BetaClient(tokenManager, getLoginUrl);
  }

  async get<T = unknown>(url: string, params?: Record<string, unknown>): Promise<T> {
    const res: AxiosResponse<T> = await this.http.get(url, { params });
    return res.data;
  }

  async getAll<T = unknown>(url: string, params?: Record<string, unknown>): Promise<T[]> {
    const results: T[] = [];
    let nextUrl: string | undefined = url;
    let queryParams: Record<string, unknown> | undefined = params;

    while (nextUrl) {
      const data: { value: T[]; '@odata.nextLink'?: string } = await this.get(nextUrl, queryParams);
      results.push(...(data.value ?? []));
      nextUrl = data['@odata.nextLink'];
      queryParams = undefined;
    }

    return results;
  }

  async post<T = unknown>(url: string, body: unknown, config?: AxiosRequestConfig): Promise<T> {
    const res: AxiosResponse<T> = await this.http.post(url, body, config);
    return res.data;
  }

  async patch<T = unknown>(url: string, body: unknown): Promise<T> {
    const res: AxiosResponse<T> = await this.http.patch(url, body);
    return res.data;
  }

  async put<T = unknown>(url: string, body: unknown, config?: AxiosRequestConfig): Promise<T> {
    const res: AxiosResponse<T> = await this.http.put(url, body, config);
    return res.data;
  }

  async delete(url: string): Promise<void> {
    await this.http.delete(url);
  }
}

class BetaClient {
  private http: AxiosInstance;

  constructor(tokenManager: TokenManager, getLoginUrl?: () => string) {
    this.http = createAxiosInstance(GRAPH_BETA, 'graph(beta)', tokenManager, getLoginUrl);
  }

  async get<T = unknown>(url: string, params?: Record<string, unknown>): Promise<T> {
    const res: AxiosResponse<T> = await this.http.get(url, { params });
    return res.data;
  }

  async getAll<T = unknown>(url: string, params?: Record<string, unknown>): Promise<T[]> {
    const results: T[] = [];
    let nextUrl: string | undefined = url;
    let queryParams: Record<string, unknown> | undefined = params;

    while (nextUrl) {
      const data: { value: T[]; '@odata.nextLink'?: string } = await this.get(nextUrl, queryParams);
      results.push(...(data.value ?? []));
      nextUrl = data['@odata.nextLink'];
      queryParams = undefined;
    }

    return results;
  }

  async post<T = unknown>(url: string, body: unknown, config?: AxiosRequestConfig): Promise<T> {
    const res: AxiosResponse<T> = await this.http.post(url, body, config);
    return res.data;
  }

  async patch<T = unknown>(url: string, body: unknown): Promise<T> {
    const res: AxiosResponse<T> = await this.http.patch(url, body);
    return res.data;
  }

  async delete(url: string): Promise<void> {
    await this.http.delete(url);
  }
}
