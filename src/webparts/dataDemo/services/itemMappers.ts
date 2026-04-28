// ABOUTME: Shared write-payload shaping for SharePoint REST and PnPjs services.
// ABOUTME: Keeps the SP-flavored field encoding out of the transport-specific service files.

import { IEventItem } from '../models/IEventItem';

export function toSpWritePayload(item: IEventItem): Record<string, unknown> {
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
