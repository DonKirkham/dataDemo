// ABOUTME: Calls the elevated Azure Functions API (apiDemo) via PnPjs's low-level Queryable pipeline.
// ABOUTME: Shows PnPjs driving an arbitrary REST endpoint — BearerToken/Origin auth + DefaultParse, so no manual fetch/response.ok/.json().

import {
  Queryable,
  get,
  post,
  patch,
  del,
  op,
  body,
  headers,
  BearerToken,
  DefaultParse,
  BrowserFetchWithRetry,
  ResolveOnData,
  RejectOnError
} from '@pnp/queryable';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { ApiMode } from './ApiSpServiceBase';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';

// Same field set the SharePoint/PnP services request, so the table renders
// identically regardless of which endpoint produced the rows.
const READ_SELECT = [
  'Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink',
  'Speaker/Id', 'Speaker/Title', 'Speaker/EMail'
].join(',');

// Read responses are shaped { mode, operation: 'read', item: <array> }; writes
// echo the affected row as { ..., item: <object> }.
interface IApiReadResponse { item?: IEventItem[]; }
interface IApiWriteResponse { item: IEventItem; }

export class PnPjsApiSpService implements ISpService {
  // One configured Queryable. Per-request children are cloned from it via the
  // [parent, path] tuple and inherit its behaviors (auth, parse, fetch), so each
  // method only sets the path + query string.
  private readonly api: Queryable;

  constructor(
    apiBaseUrl: string,
    private readonly siteUrl: string,
    private readonly mode: ApiMode,
    accessToken?: string
  ) {
    let q = new Queryable(apiBaseUrl.replace(/\/$/, ''))
      .using(BrowserFetchWithRetry())
      // DefaultParse throws an HttpRequestError on non-2xx and yields parsed JSON
      // on success — the methods below never touch response.ok or .json().
      .using(DefaultParse())
      // A bare Queryable doesn't settle its own promise: ResolveOnData resolves it
      // when data is emitted, RejectOnError rejects it (and registers the error
      // observer) on failure. The spfi/graphfi SPFx bundles add these for us; here
      // we wire them by hand. Without them every call hangs forever.
      .using(ResolveOnData(), RejectOnError());

    // Entra App: attach the bearer token PnPjs sends on every request. Simple
    // Auth (domain) is origin-restricted, so the browser's Origin header — added
    // automatically by fetch — is all the API needs; no token.
    if (accessToken) {
      q = q.using(BearerToken(accessToken));
    }

    this.api = q;
  }

  // Builds a child Queryable for /api/<mode>[/id], inheriting the base's behaviors.
  // The (parent, path) form combines the base URL with the path; the [parent, path]
  // tuple form would instead REPLACE the base URL, so we must not use it here.
  private endpoint(itemId?: number): Queryable {
    const path = itemId === undefined ? `api/${this.mode}` : `api/${this.mode}/${itemId}`;
    return new Queryable(this.api, path);
  }

  // siteUrl/listTitle ride in the query string for GET/DELETE (no body); the
  // API's resolveTarget accepts them there or in the JSON body.
  private withTarget(q: Queryable, list: IListIdentifier): Queryable {
    q.query.set('siteUrl', this.siteUrl);
    q.query.set('listTitle', list.title);
    return q;
  }

  // READ: PnPjs runs the GET, error-checks, and parses — we just read .item.
  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    const q = this.withTarget(this.endpoint(), list);
    q.query.set('$select', READ_SELECT);
    q.query.set('$expand', 'Speaker');
    q.query.set('$filter', `SessionDate ge datetime'${startOfTodayIso()}'`);
    q.query.set('$orderby', 'SessionDate asc');

    const result = await op<IApiReadResponse>(q, get);
    return result.item ?? [];
  }

  // CREATE: body() JSON-stringifies the payload; headers() sets content-type.
  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const result = await op<IApiWriteResponse>(
      this.endpoint(),
      post,
      body(
        { siteUrl: this.siteUrl, listTitle: list.title, fields: toSpWritePayload(item) },
        headers({ 'content-type': 'application/json' })
      )
    );
    return result.item;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    await op(
      this.endpoint(itemId),
      patch,
      body(
        { siteUrl: this.siteUrl, listTitle: list.title, fields: toSpWritePayload(item) },
        headers({ 'content-type': 'application/json' })
      )
    );
    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await op(this.withTarget(this.endpoint(itemId), list), del);
  }
}
