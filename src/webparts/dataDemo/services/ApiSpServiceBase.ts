// ABOUTME: Shared CRUD logic for the elevated Azure Functions API (apiDemo) endpoints.
// ABOUTME: Subclasses supply the HTTP client — anonymous HttpClient (domain) vs. AadHttpClient (entra).

import { HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';

// 'domain' -> /api/domain (origin-restricted), 'entra' -> /api/entra (bearer token).
export type ApiMode = 'domain' | 'entra';

// The same field set the SharePoint/PnP services request, so the table renders
// identically regardless of which endpoint produced the rows.
const READ_SELECT = [
  'Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink',
  'Speaker/Id', 'Speaker/Title', 'Speaker/EMail'
].join(',');

const JSON_HEADERS = { 'content-type': 'application/json' };

export abstract class ApiSpServiceBase implements ISpService {
  constructor(
    protected readonly apiBaseUrl: string,
    protected readonly siteUrl: string,
    protected readonly mode: ApiMode
  ) {}

  // Subclasses issue the request with their SPFx HTTP client (HttpClient / AadHttpClient).
  protected abstract send(url: string, options: IHttpClientOptions): Promise<HttpClientResponse>;

  private endpoint(itemId?: number): string {
    const base = `${this.apiBaseUrl.replace(/\/$/, '')}/api/${this.mode}`;
    return itemId === undefined ? base : `${base}/${itemId}`;
  }

  // siteUrl/listTitle ride in the query string for GET/DELETE (no body) and in
  // the JSON body for POST/PATCH. The API's resolveTarget accepts either.
  private targetQuery(list: IListIdentifier): string {
    return `siteUrl=${encodeURIComponent(this.siteUrl)}&listTitle=${encodeURIComponent(list.title)}`;
  }

  // The API wraps failures as { error: string }; surface that when present.
  private async errorMessage(response: HttpClientResponse, fallback: string): Promise<string> {
    try {
      const body = await response.json();
      return body?.error ?? fallback;
    } catch {
      return fallback;
    }
  }

  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    const filter = encodeURIComponent(`SessionDate ge datetime'${startOfTodayIso()}'`);
    const url = `${this.endpoint()}?${this.targetQuery(list)}`
      + `&$select=${encodeURIComponent(READ_SELECT)}`
      + `&$expand=Speaker`
      + `&$filter=${filter}`
      + `&$orderby=${encodeURIComponent('SessionDate asc')}`;

    const response = await this.send(url, { method: 'GET' });
    if (!response.ok) {
      throw new Error(await this.errorMessage(response, `Failed to get items: ${response.statusText}`));
    }

    // Read responses are shaped { mode, operation: 'read', item: <array> }.
    const body = await response.json();
    return (body.item ?? []) as IEventItem[];
  }

  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const response = await this.send(this.endpoint(), {
      method: 'POST',
      headers: JSON_HEADERS,
      body: JSON.stringify({ siteUrl: this.siteUrl, listTitle: list.title, fields: toSpWritePayload(item) })
    });
    if (!response.ok) {
      throw new Error(await this.errorMessage(response, `Failed to create item: ${response.statusText}`));
    }

    const body = await response.json();
    return body.item as IEventItem;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    const response = await this.send(this.endpoint(itemId), {
      method: 'PATCH',
      headers: JSON_HEADERS,
      body: JSON.stringify({ siteUrl: this.siteUrl, listTitle: list.title, fields: toSpWritePayload(item) })
    });
    if (!response.ok) {
      throw new Error(await this.errorMessage(response, `Failed to update item ${itemId}: ${response.statusText}`));
    }

    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    const url = `${this.endpoint(itemId)}?${this.targetQuery(list)}`;
    const response = await this.send(url, { method: 'DELETE' });
    if (!response.ok) {
      throw new Error(await this.errorMessage(response, `Failed to delete item ${itemId}: ${response.statusText}`));
    }
  }
}
