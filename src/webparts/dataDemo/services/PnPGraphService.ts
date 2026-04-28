// ABOUTME: SharePoint CRUD operations using @pnp/graph (PnP JS v4 Graph).
// ABOUTME: Uses graphfi() with SPFx behavior for authenticated Graph access to list items.

import { GraphFI } from '@pnp/graph';
import '@pnp/graph/sites';
import '@pnp/graph/lists';
import '@pnp/graph/list-item';
import { InjectHeaders } from '@pnp/queryable';
import { Logger, LogLevel } from '@pnp/logging';
import { logDebug } from './logDebug';
import { IListItem, ISpeaker, IUrlField, SessionType } from '../models/IListItem';
import { ISpService, IListIdentifier } from './ISpService';
import { startOfTodayIsoNoMs } from './dateUtils';

interface IGraphSpeakerLookup {
  LookupId: number;
  LookupValue?: string;
  Email?: string;
}

interface IGraphUrlField {
  Description?: string;
  Url?: string;
}

interface IGraphFields {
  Title?: string;
  Session?: string;
  SessionDate?: string;
  SessionType?: string;
  EventSite?: string | IGraphUrlField;
  Speaker?: IGraphSpeakerLookup[];
}

const FIELD_SELECT = 'Title,Session,SessionDate,SessionType,EventSite,Speaker';

function toListItem(rawId: string, fields: IGraphFields | undefined): IListItem {
  const f = fields ?? {};
  const speakers: ISpeaker[] | undefined = f.Speaker?.map((s) => ({
    Id: s.LookupId,
    Title: s.LookupValue ?? '',
    EMail: s.Email
  }));
  const eventSite: IUrlField | undefined = typeof f.EventSite === 'string'
    ? { Url: f.EventSite }
    : f.EventSite?.Url
      ? { Url: f.EventSite.Url, Description: f.EventSite.Description }
      : undefined;

  return {
    Id: parseInt(rawId, 10),
    Title: f.Title ?? '',
    Session: f.Session,
    SessionDate: f.SessionDate,
    SessionType: f.SessionType as SessionType | undefined,
    EventSite: eventSite,
    Speaker: speakers
  };
}

export class PnPGraphService implements ISpService {
  constructor(
    private graph: GraphFI,
    private siteId: string
  ) {}

  public async getItems(list: IListIdentifier): Promise<IListItem[]> {
    Logger.write(`[DataDemo] PnPGraphService.getItems: list=${list.id}`, LogLevel.Info);
    const itemsQuery = this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .using(InjectHeaders({ Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' }));

    itemsQuery.query.set('expand', `fields(select=${FIELD_SELECT})`);
    itemsQuery.query.set('$filter', `fields/SessionDate ge '${startOfTodayIsoNoMs()}'`);
    itemsQuery.query.set('$orderby', 'fields/SessionDate asc');

    const items = await itemsQuery();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = items.map((item: any) => toListItem(item.id, item.fields as IGraphFields));
    logDebug('PnPGraphService.getItems result:', result);
    return result;
  }

  public async getItem(list: IListIdentifier, itemId: number): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPGraphService.getItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    const itemQuery = this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString());

    itemQuery.query.set('expand', `fields(select=${FIELD_SELECT})`);

    const item = await itemQuery();

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const result = toListItem((item as any).id ?? itemId.toString(), (item as any).fields as IGraphFields);
    logDebug('PnPGraphService.getItem result:', result);
    return result;
  }

  private toWriteFields(item: IListItem): Record<string, unknown> {
    const fields: Record<string, unknown> = { Title: item.Title };
    if (item.Session !== undefined) fields.Session = item.Session;
    if (item.SessionDate !== undefined) fields.SessionDate = item.SessionDate;
    if (item.SessionType !== undefined) fields.SessionType = item.SessionType;
    if (item.EventSite !== undefined) {
      fields.EventSite = item.EventSite?.Url
        ? { Url: item.EventSite.Url, Description: item.EventSite.Description ?? item.EventSite.Url }
        : null;
    }
    // Speaker is intentionally not written via Graph — Person field writes require lookup-id wrappers.
    return fields;
  }

  public async createItem(list: IListIdentifier, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPGraphService.createItem: list=${list.id}`, LogLevel.Info);
    // FieldValueSet is typed as empty in MS Graph types, so we cast.
    const result = await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .add({ fields: this.toWriteFields(item) } as any);

    const created = {
      ...item,
      Id: parseInt(result.id, 10)
    };
    logDebug('PnPGraphService.createItem result:', created);
    return created;
  }

  public async updateItem(list: IListIdentifier, itemId: number, item: IListItem): Promise<IListItem> {
    Logger.write(`[DataDemo] PnPGraphService.updateItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      .update({ fields: this.toWriteFields(item) } as any);

    const result = { ...item, Id: itemId };
    logDebug('PnPGraphService.updateItem result:', result);
    return result;
  }

  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    Logger.write(`[DataDemo] PnPGraphService.deleteItem: list=${list.id}, id=${itemId}`, LogLevel.Info);
    await this.graph.sites
      .getById(this.siteId)
      .lists
      .getById(list.id)
      .items
      .getById(itemId.toString())
      .delete();
    logDebug('PnPGraphService.deleteItem deleted id:', itemId);
  }
}
