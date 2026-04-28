// ABOUTME: SharePoint CRUD operations using the built-in SPHttpClient REST API.
// ABOUTME: No additional packages required — uses SPFx context directly against /_api/web/lists.

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from './logDebug';
import { IListItem } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';
import { startOfTodayIso } from './dateUtils';

export class RestSpService implements ISpService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) {}

  private static readonly SELECT = [
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
  ].join(',');

  private static readonly EXPAND = 'Speaker';

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    Logger.write(`[DataDemo] RestSpService.getItems: list=${list.title}`, LogLevel.Info);
    const filter = encodeURIComponent(`SessionDate ge datetime'${startOfTodayIso()}'`);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items?$select=${RestSpService.SELECT}&$expand=${RestSpService.EXPAND}&$filter=${filter}&$orderby=SessionDate asc`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get items: ${response.statusText}`);
    }

    const data = await response.json();
    logDebug('RestSpService.getItems result:', data.value);
    return data.value as IListItem[];
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    Logger.write(`[DataDemo] RestSpService.getItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})?$select=${RestSpService.SELECT}&$expand=${RestSpService.EXPAND}`;
    const response: SPHttpClientResponse = await this.spHttpClient.get(
      url,
      SPHttpClient.configurations.v1
    );

    if (!response.ok) {
      throw new Error(`Failed to get item ${itemId}: ${response.statusText}`);
    }

    const result = await response.json() as IListItem;
    logDebug('RestSpService.getItem result:', result);
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
    Logger.write(`[DataDemo] RestSpService.createItem: list=${list.title}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`;
    const options: ISPHttpClientOptions = {
      body: JSON.stringify(this.toWritePayload(item))
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to create item: ${response.statusText}`);
    }

    const result = await response.json() as IListItem;
    logDebug('RestSpService.createItem result:', result);
    return result;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] RestSpService.updateItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(this.toWritePayload(item))
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to update item ${itemId}: ${response.statusText}`);
    }

    const result = { ...item, Id: itemId };
    logDebug('RestSpService.updateItem result:', result);
    return result;
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    Logger.write(`[DataDemo] RestSpService.deleteItem: list=${list.title}, id=${itemId}`, LogLevel.Info);
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to delete item ${itemId}: ${response.statusText}`);
    }

    logDebug('RestSpService.deleteItem deleted id:', itemId);
  }
}
