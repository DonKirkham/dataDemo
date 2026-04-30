// ABOUTME: SharePoint CRUD operations using @pnp/sp (PnP JS v4).
// ABOUTME: Uses spfi() with SPFx behavior for authenticated list item access.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';
// eslint-disable-next-line @typescript-eslint/no-unused-vars
import { CacheNever } from '@pnp/queryable';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';

export class PnPjsSpService implements ISpService {
  constructor(private sp: SPFI) {}

  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    const filter = `SessionDate ge datetime'${startOfTodayIso()}'`;

    // * Single-call version (comment out to use the batched block):
    return await this.sp.web.lists
      .getByTitle(list.title)
      .items
      //.using(CacheNever()) // comment out to allow caching for this call
      .select('Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink', 'Speaker/Id', 'Speaker/Title', 'Speaker/EMail')
      .expand('Speaker')
      .filter(filter)
      .orderBy('SessionDate', true)() as IEventItem[];
    // */ //--- END SINGLE CALL ---

    // --- BATCHED: pull items 5 at a time across pages in one $batch request ---
    // Comment out this whole block to fall back to the single-call query below.
    /*
    const [batchedSP, executeBatch] = this.sp.batched({ maxRequests: 100 });
    const pageSize = 5;
    const pageCount = 100; // up to 250 items in one batch — bump if your list is larger
    const pagePromises: Promise<IEventItem[]>[] = [];
    for (let i = 0; i < pageCount; i++) {
      pagePromises.push(
          batchedSP.web.lists
          .getByTitle(list.title)
          .items
          .select('Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink', 'Speaker/Id', 'Speaker/Title', 'Speaker/EMail')
          .expand('Speaker')
          //.filter(filter)
          .orderBy('SessionDate', true)
          .top(pageSize)
          .skip(i * pageSize)() as Promise<IEventItem[]>
      );
    }
    await executeBatch();
    const pages = await Promise.all(pagePromises);
    return pages.reduce<IEventItem[]>((acc, page) => acc.concat(page), []);
    // */ //--- END BATCHED ---


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
