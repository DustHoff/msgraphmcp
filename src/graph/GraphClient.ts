import axios, { AxiosInstance, AxiosRequestConfig, AxiosResponse, InternalAxiosRequestConfig } from 'axios';
import { TokenManager } from '../auth/TokenManager';
import { logger } from '../logger';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const GRAPH_BETA = 'https://graph.microsoft.com/beta';

// Extend InternalAxiosRequestConfig to carry request start time for duration logging
interface TimedRequestConfig extends InternalAxiosRequestConfig {
  _startMs?: number;
  _retried?: boolean;
}

export class GraphClient {
  private http: AxiosInstance;
  private tokenManager: TokenManager;
  readonly beta: BetaClient;

  constructor(tokenManager: TokenManager) {
    this.tokenManager = tokenManager;
    this.http = axios.create({ baseURL: GRAPH_BASE });
    this.beta = new BetaClient(tokenManager);

    // ── Auth interceptor ──────────────────────────────────────────────────────
    this.http.interceptors.request.use(async (config: TimedRequestConfig) => {
      const token = await this.tokenManager.getAccessToken();
      config.headers.Authorization = `Bearer ${token}`;
      config.headers['Content-Type'] = 'application/json';
      config._startMs = Date.now();
      return config;
    });

    // ── Logging + success response ────────────────────────────────────────────
    this.http.interceptors.response.use(
      (res: AxiosResponse) => {
        const cfg = res.config as TimedRequestConfig;
        const duration = cfg._startMs !== undefined ? Date.now() - cfg._startMs : undefined;
        logger.info('graph ok', {
          method: cfg.method?.toUpperCase(),
          url: cfg.url,
          status: res.status,
          ...(duration !== undefined && { durationMs: duration }),
        });
        return res;
      },
      async (error) => {
        const cfg = error.config as TimedRequestConfig | undefined;
        const duration = cfg?._startMs !== undefined ? Date.now() - cfg._startMs : undefined;
        const status: number | undefined = error.response?.status;
        const msg: string = error.response?.data?.error?.message || error.message;

        if (status === 401 && !cfg?._retried) {
          logger.warn('graph 401 — retrying with fresh token', {
            method: cfg?.method?.toUpperCase(),
            url: cfg?.url,
            ...(duration !== undefined && { durationMs: duration }),
          });
          const token = await this.tokenManager.getAccessToken();
          error.config.headers.Authorization = `Bearer ${token}`;
          error.config._retried = true;
          return this.http.request(error.config);
        }

        logger.error('graph error', {
          method: cfg?.method?.toUpperCase(),
          url: cfg?.url,
          status,
          message: msg,
          ...(duration !== undefined && { durationMs: duration }),
        });
        throw new Error(`Graph API error ${status}: ${msg}`);
      }
    );
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
  private tokenManager: TokenManager;

  constructor(tokenManager: TokenManager) {
    this.tokenManager = tokenManager;
    this.http = axios.create({ baseURL: GRAPH_BETA });

    this.http.interceptors.request.use(async (config: TimedRequestConfig) => {
      const token = await this.tokenManager.getAccessToken();
      config.headers.Authorization = `Bearer ${token}`;
      config.headers['Content-Type'] = 'application/json';
      config._startMs = Date.now();
      return config;
    });

    this.http.interceptors.response.use(
      (res: AxiosResponse) => {
        const cfg = res.config as TimedRequestConfig;
        const duration = cfg._startMs !== undefined ? Date.now() - cfg._startMs : undefined;
        logger.info('graph(beta) ok', {
          method: cfg.method?.toUpperCase(),
          url: cfg.url,
          status: res.status,
          ...(duration !== undefined && { durationMs: duration }),
        });
        return res;
      },
      async (error) => {
        const cfg = error.config as TimedRequestConfig | undefined;
        const duration = cfg?._startMs !== undefined ? Date.now() - cfg._startMs : undefined;
        const status: number | undefined = error.response?.status;
        const msg: string = error.response?.data?.error?.message || error.message;

        if (status === 401 && !cfg?._retried) {
          logger.warn('graph(beta) 401 — retrying with fresh token', {
            method: cfg?.method?.toUpperCase(),
            url: cfg?.url,
            ...(duration !== undefined && { durationMs: duration }),
          });
          const token = await this.tokenManager.getAccessToken();
          error.config.headers.Authorization = `Bearer ${token}`;
          error.config._retried = true;
          return this.http.request(error.config);
        }

        logger.error('graph(beta) error', {
          method: cfg?.method?.toUpperCase(),
          url: cfg?.url,
          status,
          message: msg,
          ...(duration !== undefined && { durationMs: duration }),
        });
        throw new Error(`Graph API error ${status}: ${msg}`);
      }
    );
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
