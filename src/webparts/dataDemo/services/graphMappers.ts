// ABOUTME: Shared Graph item shaping for the MS Graph and PnP Graph services.
// ABOUTME: Maps between Graph fields payloads and the IEventItem model used by the UI.

import { IEventItem, ISpeaker, IUrlField, SessionType } from '../models/IEventItem';

export const GRAPH_FIELD_SELECT = 'Title,Session,SessionDate,SessionType,EventSite,Speaker';

interface IGraphSpeakerLookup {
  LookupId: number;
  LookupValue?: string;
  Email?: string;
}

interface IGraphUrlField {
  Description?: string;
  Url?: string;
}

export interface IGraphFields {
  Title?: string;
  Session?: string;
  SessionDate?: string;
  SessionType?: string;
  EventSite?: string | IGraphUrlField;
  Speaker?: IGraphSpeakerLookup[];
}

export interface IGraphListItem {
  id: string;
  fields?: IGraphFields;
}

export function fromGraphItem(graphItem: IGraphListItem): IEventItem {
  const f = graphItem.fields ?? {};
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
    Id: parseInt(graphItem.id, 10),
    Title: f.Title ?? '',
    Session: f.Session,
    SessionDate: f.SessionDate,
    SessionType: f.SessionType as SessionType | undefined,
    EventSite: eventSite,
    Speaker: speakers
  };
}

export function toGraphWriteFields(item: IEventItem): Record<string, unknown> {
  const fields: Record<string, unknown> = { Title: item.Title };
  if (item.Session !== undefined) fields.Session = item.Session;
  if (item.SessionDate !== undefined) fields.SessionDate = item.SessionDate;
  if (item.SessionType !== undefined) fields.SessionType = item.SessionType;
  if (item.EventSite !== undefined) {
    fields.EventSite = item.EventSite?.Url
      ? { Url: item.EventSite.Url, Description: item.EventSite.Description ?? item.EventSite.Url }
      : null;
  }
  return fields;
}
