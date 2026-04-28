// ABOUTME: SharePoint CRUD operations using the built-in SPHttpClient REST API.
// ABOUTME: No additional packages required — uses SPFx context directly against /_api/web/lists.

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';

export class SpfxSpService implements ISpService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) {}

  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`
      + `?$select=Id,Title,Session,SessionDate,SessionType,EventSite,SessionLink,Speaker/Id,Speaker/Title,Speaker/EMail`
      + `&$expand=Speaker`
      + `&$filter=${encodeURIComponent(`SessionDate ge datetime'${startOfTodayIso()}'`)}`
      + `&$orderby=SessionDate asc`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(`Failed to get items: ${response.statusText}`);
    }

    const data = await response.json();
    return data.value as IEventItem[];
  }

  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`;
    const options: ISPHttpClientOptions = {
      body: JSON.stringify(toSpWritePayload(item))
    };

    const response: SPHttpClientResponse = await this.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      options
    );

    if (!response.ok) {
      throw new Error(`Failed to create item: ${response.statusText}`);
    }

    return await response.json() as IEventItem;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE'
      },
      body: JSON.stringify(toSpWritePayload(item))
    };

    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);

    if (!response.ok) {
      throw new Error(`Failed to update item ${itemId}: ${response.statusText}`);
    }

    return { ...item, Id: itemId };
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
    const options: ISPHttpClientOptions = {
      headers: {
        'IF-MATCH': '*',
        'X-HTTP-Method': 'DELETE'
      }
    };

    const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);

    if (!response.ok) {
      throw new Error(`Failed to delete item ${itemId}: ${response.statusText}`);
    }
  }
}
