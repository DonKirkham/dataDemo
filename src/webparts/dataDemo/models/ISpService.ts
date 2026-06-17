// ABOUTME: Service contract for SharePoint CRUD operations.
// ABOUTME: All service implementations (REST, PnP SP, Graph, PnP Graph) conform to this interface.

import { IEventItem } from '../models/IEventItem';

export interface IListIdentifier {
  title: string;
  id: string;
}

export interface ISpService {
  getItems(list: IListIdentifier): Promise<IEventItem[]>;
  createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem>;
  updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem>;
  deleteItem(list: IListIdentifier, itemId: number): Promise<void>;
}

// Batching options for the PnPjs services. `batchSize` maps to PnPjs's
// `maxRequests` — the number of sub-requests grouped into each $batch HTTP call.
export interface IPnPjsBatchOptions {
  useBatching: boolean;
  batchSize: number;
}
