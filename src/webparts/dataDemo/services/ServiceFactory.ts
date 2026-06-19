// ABOUTME: Factory that creates the appropriate service implementation based on the selected approach.
// ABOUTME: Handles PnP JS initialization (spfi/graphfi) and Graph client setup for each service type.

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx as spSPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/security';
import { PermissionKind } from '@pnp/sp/security';
import { graphfi, SPFx as graphSPFx } from '@pnp/graph';
import { Caching } from '@pnp/queryable';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { IJokeService } from '../models/IJokeService';
import { IGraphQueryService } from '../models/IGraphQueryService';
import { SpfxSpService } from './SpfxSpService';
import { PnPjsSpService } from './PnPjsSpService';
import { SpfxGraphSpService } from './SpfxGraphSpService';
import { PnPjsGraphSpService } from './PnPjsGraphSpService';
import { SpfxAnonymousService } from './SpfxAnonymousService';
import { PnPjsAnonymousService } from './PnPjsAnonymousService';
import { SpfxGraphQueryService } from './SpfxGraphQueryService';
import { ApiDomainSpService } from './ApiDomainSpService';
import { ApiEntraSpService } from './ApiEntraSpService';
import { PnPjsApiSpService } from './PnPjsApiSpService';

export type Transport = 'SPFx' | 'PnPjs';
export type Endpoint = 'SharePoint' | 'MS Graph (Explorer)' | 'MS Graph (SP)' | 'Anonymous' | 'Simple Auth' | 'Entra App';

export interface ISiteInfo {
  url: string;
  id: string;
}

// PnPjs-only options. `useCache` adds the Caching behavior to the spfi/graphfi
// instance. (Batching is exercised on demand via the services' runBatchDemo.)
export interface IPnPjsOptions {
  useCache: boolean;
}

// Configuration for the elevated Azure Functions API (apiDemo). Public and
// mutable so the web part can refresh it from the property pane between renders.
export interface IApiConfig {
  baseUrl: string;
  resourceUri: string;
}

// The signed-in user's effective permissions on the target list. Gates the
// read/write UI on user-identity endpoints (elevated endpoints ignore these).
export interface IListPermissions {
  canRead: boolean;
  canAdd: boolean;
  canEdit: boolean;
  canDelete: boolean;
}

export class ServiceFactory {
  constructor(
    private _context: WebPartContext,
    public api: IApiConfig
  ) {}

  public get context(): WebPartContext {
    return this._context;
  }

  public async createSpService(
    transport: Transport,
    endpoint: Endpoint,
    site: ISiteInfo,
    options?: IPnPjsOptions
  ): Promise<ISpService> {
    if (transport === 'SPFx' && endpoint === 'SharePoint') {
      return new SpfxSpService(this.context.spHttpClient, site.url);
    }

    if (transport === 'SPFx' && endpoint === 'MS Graph (SP)') {
      const graphClient = await this.context.msGraphClientFactory.getClient('3');
      return new SpfxGraphSpService(graphClient, site.id);
    }

    if (transport === 'PnPjs' && endpoint === 'SharePoint') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const sp = spfi(site.url).using(spSPFx(this.context as any));
      if (options?.useCache) {
        // Caching co-exists with the normal read path — same fluent query, now
        // served from sessionStorage on repeat calls.
        sp.using(Caching({ store: 'session' }));
      }
      return new PnPjsSpService(sp);
    }

    if (transport === 'PnPjs' && endpoint === 'MS Graph (SP)') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const graph = graphfi().using(graphSPFx(this.context as any));
      if (options?.useCache) {
        graph.using(Caching({ store: 'session' }));
      }
      return new PnPjsGraphSpService(graph, site.id);
    }

    // The elevated-API endpoints call the apiDemo Azure Function (write app-only).
    // Both transports hit the same endpoint; only who makes the HTTP call differs:
    // SPFx uses the native HttpClient/AadHttpClient, PnPjs uses its Queryable pipeline.
    if (endpoint === 'Simple Auth') {
      if (transport === 'PnPjs') {
        // Origin-restricted: no token — fetch's Origin header is the credential.
        return new PnPjsApiSpService(this.api.baseUrl, site.url, 'domain');
      }
      return new ApiDomainSpService(this.context.httpClient, this.api.baseUrl, site.url);
    }

    if (endpoint === 'Entra App') {
      if (transport === 'PnPjs') {
        // Acquire a bearer token for the API's app registration, then let PnPjs
        // send it via the BearerToken behavior (no AadHttpClient involved).
        const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
        const token = await tokenProvider.getToken(this.api.resourceUri);
        return new PnPjsApiSpService(this.api.baseUrl, site.url, 'entra', token);
      }
      const aadClient = await this.context.aadHttpClientFactory.getClient(this.api.resourceUri);
      return new ApiEntraSpService(aadClient, this.api.baseUrl, site.url);
    }

    throw new Error(`Unsupported SharePoint combination: ${transport} + ${endpoint}`);
  }

  // Reads the signed-in user's effective list permissions, once per site/list.
  // Reflects what the *user* can do directly — what the user-identity tabs gate on.
  public async getListPermissions(site: ISiteInfo, list: IListIdentifier): Promise<IListPermissions> {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sp = spfi(site.url).using(spSPFx(this.context as any));
    const splist = sp.web.lists.getByTitle(list.title);
    const perms = await splist.getCurrentUserEffectivePermissions();
    return {
      canRead: splist.hasPermissions(perms, PermissionKind.ViewListItems),
      canAdd: splist.hasPermissions(perms, PermissionKind.AddListItems),
      canEdit: splist.hasPermissions(perms, PermissionKind.EditListItems),
      canDelete: splist.hasPermissions(perms, PermissionKind.DeleteListItems)
    };
  }

  public createJokeService(transport: Transport): IJokeService {
    if (transport === 'SPFx') {
      return new SpfxAnonymousService(this.context.httpClient);
    }
    if (transport === 'PnPjs') {
      return new PnPjsAnonymousService();
    }
    throw new Error(`Unsupported transport for joke service: ${transport}`);
  }

  public async createGraphQueryService(): Promise<IGraphQueryService> {
    const graphClient = await this.context.msGraphClientFactory.getClient('3');
    return new SpfxGraphQueryService(graphClient);
  }
}
