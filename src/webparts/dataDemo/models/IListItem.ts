// ABOUTME: Data model for the Speaking Events SharePoint list.
// ABOUTME: Title holds the Event name; remaining fields describe the session.

export type SessionType =
  | '30 minute session'
  | '45 minute session'
  | '50 minute session'
  | '60 minute session'
  | '70 minute session'
  | '75 minute session'
  | 'Half day workshop'
  | 'Full day workshop';

export interface ISpeaker {
  Id: number;
  Title: string;
  EMail?: string;
}

export interface IUrlField {
  Url: string;
  Description?: string;
}

export interface IThumbnailField {
  serverRelativeUrl?: string;
  serverUrl?: string;
  fileName?: string;
}

export interface IListItem {
  Id?: number;
  Title: string;
  Speaker?: ISpeaker[];
  Session?: string;
  SessionDate?: string;
  SessionType?: SessionType;
  EventLogo?: IThumbnailField;
  EventSite?: IUrlField;
  SessionLink?: IUrlField;
}
