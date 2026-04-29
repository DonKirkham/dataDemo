// ABOUTME: Factory that creates the appropriate service implementation based on the selected approach.
// ABOUTME: Handles PnP JS initialization (spfi/graphfi) and Graph client setup for each service type.

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { spfi, SPFx as spSPFx } from '@pnp/sp';
import { graphfi, SPFx as graphSPFx } from '@pnp/graph';
import { Caching } from '@pnp/queryable';
import { ISpService } from '../models/ISpService';
import { IJokeService } from '../models/IJokeService';
import { IGraphQueryService } from '../models/IGraphQueryService';
import { SpfxSpService } from './SpfxSpService';
import { PnPjsSpService } from './PnPjsSpService';
import { SpfxGraphSpService } from './SpfxGraphSpService';
import { PnPjsGraphSpService } from './PnPjsGraphSpService';
import { SpfxAnonymousService } from './SpfxAnonymousService';
import { PnPjsAnonymousService } from './PnPjsAnonymousService';
import { SpfxGraphQueryService } from './SpfxGraphQueryService';

export type Transport = 'SPFx' | 'PnPjs';
export type Endpoint = 'SharePoint' | 'MS Graph' | 'MS Graph (SP)' | 'Anonymous' | 'Simple Auth' | 'Entra App';

export interface ISiteInfo {
  url: string;
  id: string;
}

export class ServiceFactory {
  constructor(private _context: WebPartContext) {}

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

    throw new Error(`Unsupported SharePoint combination: ${transport} + ${endpoint}`);
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
