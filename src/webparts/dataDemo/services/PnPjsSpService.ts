// ABOUTME: SharePoint CRUD operations using @pnp/sp (PnP JS v4).
// ABOUTME: Uses spfi() with SPFx behavior for authenticated list item access.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';

export class PnPjsSpService implements ISpService {
  constructor(private sp: SPFI) {}

  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    return await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select('Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink', 'Speaker/Id', 'Speaker/Title', 'Speaker/EMail')
      .expand('Speaker')
      .filter(`SessionDate ge datetime'${startOfTodayIso()}'`)
      .orderBy('SessionDate', true)() as IEventItem[];
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
