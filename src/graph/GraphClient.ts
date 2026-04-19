import axios, { AxiosInstance, AxiosRequestConfig, AxiosResponse } from 'axios';
import { TokenManager } from '../auth/TokenManager';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export class GraphClient {
  private http: AxiosInstance;
  private tokenManager: TokenManager;

  constructor(tokenManager: TokenManager) {
    this.tokenManager = tokenManager;
    this.http = axios.create({ baseURL: GRAPH_BASE });

    this.http.interceptors.request.use(async (config) => {
      const token = await this.tokenManager.getAccessToken();
      config.headers.Authorization = `Bearer ${token}`;
      config.headers['Content-Type'] = 'application/json';
      return config;
    });

    this.http.interceptors.response.use(
      (res) => res,
      async (error) => {
        if (error.response?.status === 401) {
          // Force token refresh by retrying once
          const token = await this.tokenManager.getAccessToken();
          error.config.headers.Authorization = `Bearer ${token}`;
          return this.http.request(error.config);
        }
        const msg = error.response?.data?.error?.message || error.message;
        throw new Error(`Graph API error ${error.response?.status}: ${msg}`);
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
      queryParams = undefined; // nextLink already contains params
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
