// ABOUTME: Calls the elevated Azure Functions API /api/entra endpoint (Entra-token protected).
// ABOUTME: Uses an AadHttpClient that acquires a bearer token for the API's app registration.

import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { ApiSpServiceBase } from './ApiSpServiceBase';

export class ApiEntraSpService extends ApiSpServiceBase {
  constructor(
    private readonly aadClient: AadHttpClient,
    apiBaseUrl: string,
    siteUrl: string
  ) {
    super(apiBaseUrl, siteUrl, 'entra');
  }

  protected send(url: string, options: IHttpClientOptions): Promise<HttpClientResponse> {
    return this.aadClient.fetch(url, AadHttpClient.configurations.v1, options);
  }
}
