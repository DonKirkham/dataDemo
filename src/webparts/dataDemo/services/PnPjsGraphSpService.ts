// ABOUTME: SharePoint CRUD operations using @pnp/graph (PnP JS v4 Graph).
// ABOUTME: Uses graphfi() with SPFx behavior for authenticated Graph access to list items.

import { GraphFI } from '@pnp/graph';
import '@pnp/graph/sites';
import '@pnp/graph/lists';
import '@pnp/graph/list-item';
import { InjectHeaders } from '@pnp/queryable';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
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

export class PnPjsGraphSpService implements ISpService {
  constructor(
    private graph: GraphFI,
    private siteId: string
  ) {}

  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    const itemsQuery = this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .using(InjectHeaders({ Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' }));

    itemsQuery.query.set('expand', `fields(select=${GRAPH_FIELD_SELECT})`);
    itemsQuery.query.set('$filter', `fields/SessionDate ge '${startOfTodayIsoNoMs()}'`);
    itemsQuery.query.set('$orderby', 'fields/SessionDate asc');

    const items = await itemsQuery() as IGraphListItem[];
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
