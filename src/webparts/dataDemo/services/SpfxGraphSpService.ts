// ABOUTME: SharePoint CRUD operations using the MS Graph API via SPFx MSGraphClientV3.
// ABOUTME: No extra packages — uses the built-in Graph client. Mirrors PnPjsGraphSpService.

import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { startOfTodayIsoNoMs } from '../utilities/dateUtils';
import {
  GRAPH_FIELD_SELECT,
  IGraphListItem,
  fromGraphItem,
  toGraphWriteFields
} from './graphMappers';
import { MSGraphClientV3 } from '@microsoft/sp-http';

export class SpfxGraphSpService implements ISpService {
  constructor(
    private graphClient: MSGraphClientV3,
    private siteId: string
  ) {}

  // READ: upcoming sessions with the Speaker lookup expanded, sorted by date.
  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    const qs = [
      `expand=fields(select=${GRAPH_FIELD_SELECT})`,
      `$filter=fields/SessionDate ge '${startOfTodayIsoNoMs()}'`,
      `$orderby=fields/SessionDate asc`
    ].join('&');
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .version('v1.0')
      .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
      .query(qs)
      .get();

    return (response.value as IGraphListItem[]).map(fromGraphItem);
  }

  // CREATE: add a new event item.
  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .version('v1.0')
      .post({ fields: toGraphWriteFields(item) });

    return fromGraphItem(response as IGraphListItem);
  }

  // UPDATE: save changes to an existing item (PATCH).
  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}/fields`)
      .version('v1.0')
      .patch(toGraphWriteFields(item));

    return { ...item, Id: itemId };
  }

  // DELETE: remove an item by id.
  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }
}
