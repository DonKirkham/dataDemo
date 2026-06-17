// ABOUTME: SharePoint CRUD operations using the built-in SPHttpClient REST API.
// ABOUTME: No extra packages — SPFx context calls /_api/web/lists directly. Mirrors PnPjsSpService.

import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

// Columns each read selects — kept identical to PnPjsSpService for side-by-side.
const SELECT_FIELDS: string[] = [
  'Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink',
  'Speaker/Id', 'Speaker/Title', 'Speaker/EMail'
];

export class SpfxSpService implements ISpService {
  constructor(
    private spHttpClient: SPHttpClient,
    private siteUrl: string
  ) {}

  // READ: passes through to getItemsNoBatch, mirroring the PnPjs getItems/
  // getItemsNoBatch split for side-by-side reading.
  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    return this.getItemsNoBatch(list);
  }

  // hand-build the URL, GET, check response.ok, then parse + unwrap .value.
  private async getItemsNoBatch(list: IListIdentifier): Promise<IEventItem[]> {
    const filter = `SessionDate ge datetime'${startOfTodayIso()}'`;
    const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`
      + `?$select=${SELECT_FIELDS.join(',')}`
      //+ `&$top=5` // Uncomment this line to show how paging works
      + `&$expand=Speaker`
      + `&$filter=${encodeURIComponent(filter)}`
      + `&$orderby=SessionDate asc`;
    const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);

    if (!response.ok) {
      throw new Error(`Failed to get items: ${response.statusText}`);
    }

    const data = await response.json();
    return data.value as IEventItem[];
  }

  // CREATE: hand-build the URL, JSON.stringify the body, POST, check, parse.
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

  // UPDATE: POST with IF-MATCH + X-HTTP-Method: MERGE headers set by hand.
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

  // DELETE: POST with IF-MATCH + X-HTTP-Method: DELETE headers set by hand.
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
