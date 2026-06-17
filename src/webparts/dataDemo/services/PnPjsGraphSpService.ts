// ABOUTME: SharePoint CRUD operations using @pnp/graph (PnP JS v4 Graph).
// ABOUTME: Uses graphfi() with SPFx behavior for authenticated Graph access to list items.

import { GraphFI } from '@pnp/graph';
import '@pnp/graph/sites';
import '@pnp/graph/lists';
import '@pnp/graph/list-item';
import '@pnp/graph/batching';
import { InjectHeaders } from '@pnp/queryable';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier, IPnPjsBatchOptions } from '../models/ISpService';
import { startOfTodayIsoNoMs } from '../utilities/dateUtils';
import {
  GRAPH_FIELD_SELECT,
  IGraphListItem,
  fromGraphItem,
  toGraphWriteFields
} from './graphMappers';

interface IGraphAddResult {
  id: string;
}

const PREFER_NONINDEXED = 'HonorNonIndexedQueriesWarningMayFailRandomly';

const NO_BATCHING: IPnPjsBatchOptions = { useBatching: false, batchSize: 5 };

export class PnPjsGraphSpService implements ISpService {
  // Caching, when enabled, is configured on the GraphFI instance in
  // ServiceFactory, so the read methods below benefit from it transparently.
  constructor(
    private graph: GraphFI,
    private siteId: string,
    private batch: IPnPjsBatchOptions = NO_BATCHING
  ) {}

  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    return this.batch.useBatching
      ? this.getItemsBatched(list, this.batch.batchSize)
      : this.getItemsNoBatch(list);
  }

  private async getItemsNoBatch(list: IListIdentifier): Promise<IEventItem[]> {
    const itemsQuery = this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .using(InjectHeaders({ Prefer: PREFER_NONINDEXED }));

    itemsQuery.query.set('expand', `fields(select=${GRAPH_FIELD_SELECT})`);
    itemsQuery.query.set('$filter', `fields/SessionDate ge '${startOfTodayIsoNoMs()}'`);
    itemsQuery.query.set('$orderby', 'fields/SessionDate asc');

    const items = await itemsQuery() as IGraphListItem[];
    return items.map(fromGraphItem);
  }

  // Batched read: fetch each matching id in $batch calls of `batchSize`. Graph
  // caps a $batch at 20 requests; PnPjs chunks beyond that automatically.
  private async getItemsBatched(list: IListIdentifier, batchSize: number): Promise<IEventItem[]> {
    const idQuery = this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .using(InjectHeaders({ Prefer: PREFER_NONINDEXED }));

    idQuery.query.set('select', 'id');
    idQuery.query.set('$filter', `fields/SessionDate ge '${startOfTodayIsoNoMs()}'`);
    idQuery.query.set('$orderby', 'fields/SessionDate asc');

    const idRows = await idQuery() as { id: string }[];

    const [batchedGraph, executeBatch] = this.graph.batched({ maxRequests: batchSize });
    const promises = idRows.map((row) => {
      const itemQuery = batchedGraph.sites
        .getById(this.siteId)
        .lists
        .getById(list.id)
        .items
        .getById(row.id);
      itemQuery.query.set('expand', `fields(select=${GRAPH_FIELD_SELECT})`);
      return itemQuery() as Promise<IGraphListItem>;
    });
    await executeBatch();
    const items = await Promise.all(promises);
    return items.map(fromGraphItem);
  }

  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const result = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .add({ fields: toGraphWriteFields(item) }) as IGraphAddResult;
    return { ...item, Id: parseInt(result.id, 10) };
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      .update({ fields: toGraphWriteFields(item) });
    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      .delete();
  }
}
