// ABOUTME: Calls the elevated Azure Functions API /api/domain endpoint (origin-restricted).
// ABOUTME: Uses the anonymous HttpClient; the browser sends the Origin the API allow-lists.

import { HttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { ApiSpServiceBase } from './ApiSpServiceBase';

export class ApiDomainSpService extends ApiSpServiceBase {
  constructor(
    private readonly httpClient: HttpClient,
    apiBaseUrl: string,
    siteUrl: string
  ) {
    super(apiBaseUrl, siteUrl, 'domain');
  }

  protected send(url: string, options: IHttpClientOptions): Promise<HttpClientResponse> {
    return this.httpClient.fetch(url, HttpClient.configurations.v1, options);
  }
}
