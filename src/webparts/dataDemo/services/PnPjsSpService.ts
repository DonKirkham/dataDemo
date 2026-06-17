// ABOUTME: SharePoint CRUD operations using @pnp/sp (PnP JS v4).
// ABOUTME: Mirrors the SPFx REST service; adds an optional $batch read path for the batching demo.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier, IPnPjsBatchOptions } from '../models/ISpService';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';

// Columns every read selects — shared by the single-call and batched paths so
// both return identically-shaped items.
const SELECT_FIELDS: string[] = [
  'Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink',
  'Speaker/Id', 'Speaker/Title', 'Speaker/EMail'
];

const NO_BATCHING: IPnPjsBatchOptions = { useBatching: false, batchSize: 5 };

export class PnPjsSpService implements ISpService {
  // Caching, when enabled, is configured on the SPFI instance in ServiceFactory,
  // so the read methods below benefit from it transparently — no code change here.
  constructor(private sp: SPFI, private batch: IPnPjsBatchOptions = NO_BATCHING) {}

  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    return this.batch.useBatching
      ? this.getItemsBatched(list, this.batch.batchSize)
      : this.getItemsNoBatch(list);
  }

  // Single fluent query — the PnPjs equivalent of the SPFx REST getItems().
  private async getItemsNoBatch(list: IListIdentifier): Promise<IEventItem[]> {
    const filter = `SessionDate ge datetime'${startOfTodayIso()}'`;
    return await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select(...SELECT_FIELDS)
      .expand('Speaker')
      .filter(filter)
      .orderBy('SessionDate', true)() as IEventItem[];
  }

  // Batched read: find the matching ids, then fetch each by id in $batch calls of
  // `batchSize` (maxRequests). Lower it to split the work across more round-trips.
  private async getItemsBatched(list: IListIdentifier, batchSize: number): Promise<IEventItem[]> {
    const filter = `SessionDate ge datetime'${startOfTodayIso()}'`;
    const idRows = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select('Id')
      .filter(filter)
      .orderBy('SessionDate', true)() as { Id: number }[];

    const [batchedSP, executeBatch] = this.sp.batched({ maxRequests: batchSize });
    const promises = idRows.map((row) =>
      batchedSP.web.lists
        .getByTitle(list.title)
        .items
        .getById(row.Id)
        .select(...SELECT_FIELDS)
        .expand('Speaker')() as Promise<IEventItem>
    );
    await executeBatch();
    return Promise.all(promises);
  }

  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .add(toSpWritePayload(item));
    return result.data as IEventItem;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .update(toSpWritePayload(item));
    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .delete();
  }
}
