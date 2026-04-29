// ABOUTME: Props interface for the DataDemo component.
// ABOUTME: Receives the service factory, site info, and list identifier from the web part.

import { IListIdentifier } from '../models/ISpService';
import { ServiceFactory, ISiteInfo } from '../services/ServiceFactory';

export interface IDataDemoProps {
  factory: ServiceFactory | undefined;
  site: ISiteInfo;
  list: IListIdentifier;
}
