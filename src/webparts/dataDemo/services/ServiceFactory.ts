// ABOUTME: Factory that creates the appropriate service implementation based on the selected approach.
// ABOUTME: Handles PnP JS initialization (spfi/graphfi) and Graph client setup for each service type.

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx as spSPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/security';
import { PermissionKind } from '@pnp/sp/security';
import { graphfi, SPFx as graphSPFx } from '@pnp/graph';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
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

export type Transport = 'SPFx' | 'PnPjs';
export type Endpoint = 'SharePoint' | 'MS Graph (Explorer)' | 'MS Graph (SP)' | 'Anonymous' | 'Simple Auth' | 'Entra App';

export interface ISiteInfo {
  url: string;
  id: string;
}

// Configuration for the elevated Azure Functions API (apiDemo). Public and
// mutable so the web part can refresh it from the property pane between renders.
export interface IApiConfig {
  baseUrl: string;
  resourceUri: string;
}

// The signed-in user's effective permissions on the target list. Used to gate
// the read/write UI on the user-identity endpoints (the elevated API endpoints
// read/write app-only, so they ignore these).
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

  public async createSpService(transport: Transport, endpoint: Endpoint, site: ISiteInfo): Promise<ISpService> {
    if (transport === 'SPFx' && endpoint === 'SharePoint') {
      return new SpfxSpService(this.context.spHttpClient, site.url);
    }

    if (transport === 'SPFx' && endpoint === 'MS Graph (SP)') {
      const graphClient = await this.context.msGraphClientFactory.getClient('3');
      return new SpfxGraphSpService(graphClient, site.id);
    }

    if (transport === 'PnPjs' && endpoint === 'SharePoint') {
      const sp = spfi(site.url)
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        .using(spSPFx(this.context as any))
        //.using(Caching({ store: 'session' })); // comment out to disable caching
      return new PnPjsSpService(sp);
    }

    if (transport === 'PnPjs' && endpoint === 'MS Graph (SP)') {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const graph = graphfi().using(graphSPFx(this.context as any));
      return new PnPjsGraphSpService(graph, site.id);
    }

    // The elevated-API endpoints call the apiDemo Azure Function, which performs
    // the SharePoint write app-only. They behave identically under either
    // transport, so the transport pivot does not branch them.
    if (endpoint === 'Simple Auth') {
      return new ApiDomainSpService(this.context.httpClient, this.api.baseUrl, site.url);
    }

    if (endpoint === 'Entra App') {
      const aadClient = await this.context.aadHttpClientFactory.getClient(this.api.resourceUri);
      return new ApiEntraSpService(aadClient, this.api.baseUrl, site.url);
    }

    throw new Error(`Unsupported SharePoint combination: ${transport} + ${endpoint}`);
  }

  // Reads the signed-in user's effective permissions on the target list, once
  // per site/list. Independent of the selected endpoint — it reflects what the
  // *user* can do directly, which is what the user-identity tabs are gated on.
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
