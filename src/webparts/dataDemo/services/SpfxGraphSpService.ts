// ABOUTME: SharePoint CRUD operations using the MS Graph API via SPFx MSGraphClientV3.
// ABOUTME: No additional packages required — uses the built-in Graph client from SPFx context.

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { startOfTodayIsoNoMs } from '../utilities/dateUtils';
import {
  GRAPH_FIELD_SELECT,
  IGraphListItem,
  fromGraphItem,
  toGraphWriteFields
} from './graphMappers';

export class SpfxGraphSpService implements ISpService {
  constructor(
    private graphClient: MSGraphClientV3,
    private siteId: string
  ) {}

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

  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const response = await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items`)
      .version('v1.0')
      .post({ fields: toGraphWriteFields(item) });

    return fromGraphItem(response as IGraphListItem);
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}/fields`)
      .version('v1.0')
      .patch(toGraphWriteFields(item));

    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.graphClient
      .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}`)
      .version('v1.0')
      .delete();
  }
}
