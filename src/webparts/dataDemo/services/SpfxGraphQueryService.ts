// ABOUTME: Free-form MS Graph queries using the SPFx MSGraphClientV3.
// ABOUTME: Sends GET requests to any v1.0 path and returns the raw JSON response.

import { MSGraphClientV3 } from '@microsoft/sp-http';
import { IGraphQueryService } from '../models/IGraphQueryService';

export class SpfxGraphQueryService implements IGraphQueryService {
  constructor(private graphClient: MSGraphClientV3) {}

  public async runQuery(path: string): Promise<unknown> {
    return await this.graphClient.api(path).version('v1.0').get();
  }
}
