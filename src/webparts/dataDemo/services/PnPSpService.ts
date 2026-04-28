// ABOUTME: SharePoint CRUD operations using @pnp/sp (PnP JS v4).
// ABOUTME: Uses spfi() with SPFx behavior for authenticated list item access.

import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from './logDebug';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';
import { startOfTodayIso } from './dateUtils';

export class PnPSpService implements ISpService {
  constructor(private sp: SPFI) {}

  private static readonly SELECT_FIELDS: string[] = [
    'Id',
    'Title',
    'Session',
    'SessionDate',
    'SessionType',
    'EventSite',
    'SessionLink',
    'Speaker/Id',
    'Speaker/Title',
    'Speaker/EMail'
  ];

  private static readonly EXPAND_FIELDS: string[] = ['Speaker'];

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    Logger.write(`[DataDemo] PnPSpService.getItems: list=${list.title}`, LogLevel.Info);
    const todayIso = startOfTodayIso();
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select(...PnPSpService.SELECT_FIELDS)
      .expand(...PnPSpService.EXPAND_FIELDS)
      .filter(`SessionDate ge datetime'${todayIso}'`)
      .orderBy('SessionDate', true)() as IListItem[];
    logDebug('PnPSpService.getItems result:', result);
    return result;
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPSpService.getItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .select(...PnPSpService.SELECT_FIELDS)
      .expand(...PnPSpService.EXPAND_FIELDS)() as IListItem;
    logDebug('PnPSpService.getItem result:', result);
    return result;
  }

  private toWritePayload(item: IListItem): Record<string, unknown> {
    const payload: Record<string, unknown> = { Title: item.Title };
    if (item.Session !== undefined) payload.Session = item.Session;
    if (item.SessionDate !== undefined) payload.SessionDate = item.SessionDate;
    if (item.SessionType !== undefined) payload.SessionType = item.SessionType;
    if (item.EventSite !== undefined) {
      payload.EventSite = item.EventSite?.Url
        ? { Url: item.EventSite.Url, Description: item.EventSite.Description ?? item.EventSite.Url }
        : null;
    }
    if (item.Speaker !== undefined) {
      payload.SpeakerId = item.Speaker.map((s) => s.Id).filter((id) => id > 0);
    }
    return payload;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPSpService.createItem: list=${list.title}`, LogLevel.Info);
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .add(this.toWritePayload(item));

    logDebug('PnPSpService.createItem result:', result.data);
    return result.data as IListItem;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPSpService.updateItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .update(this.toWritePayload(item));

    const result = { ...item, Id: itemId };
    logDebug('PnPSpService.updateItem result:', result);
    return result;
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    Logger.write(`[DataDemo] PnPSpService.deleteItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .delete();
    logDebug('PnPSpService.deleteItem deleted id:', itemId);
  }
}
