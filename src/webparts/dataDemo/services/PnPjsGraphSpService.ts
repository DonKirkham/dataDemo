// ABOUTME: SharePoint CRUD operations using @pnp/graph (PnP JS v4 Graph).
// ABOUTME: Mirrors SpfxGraphSpService; adds runBatchDemo to showcase $batch on bulk writes.

import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { IBatchDemoService, IBatchDemoPhase, IBatchDemoResult } from '../models/IBatchDemoService';
import { startOfTodayIsoNoMs } from '../utilities/dateUtils';
import {
  GRAPH_FIELD_SELECT,
  IGraphListItem,
  fromGraphItem,
  toGraphWriteFields
} from './graphMappers';
import { GraphFI } from '@pnp/graph';
import '@pnp/graph/sites';
import '@pnp/graph/lists';
import '@pnp/graph/list-item';
import type { IListItems } from '@pnp/graph/list-item';
import { GraphQueryable, IGraphQueryable, graphPatch } from '@pnp/graph/graphqueryable';
import '@pnp/graph/batching';
import { InjectHeaders, body } from '@pnp/queryable';

interface IGraphAddResult {
  id: string;
}

const PREFER_NONINDEXED = 'HonorNonIndexedQueriesWarningMayFailRandomly';

// Graph caps a single JSON $batch at 20 sub-requests; PnPjs chunks beyond that.
const GRAPH_BATCH_LIMIT = 20;

// Marker title for the items the batch demo creates, so a later run can find and
// clean them up regardless of how the previous run ended.
const SAMPLE_PREFIX = 'SAMPLE:';
const SAMPLE_TITLE = 'SAMPLE: New event. ID: ';
const DEMO_COUNT = 10;

export class PnPjsGraphSpService implements ISpService, IBatchDemoService {
  // Caching, when enabled, is configured on the GraphFI instance in
  // ServiceFactory, so the read methods below benefit from it transparently.
  constructor(
    private graph: GraphFI,
    private siteId: string
  ) {}

  // The list's items collection, optionally from a batched root.
  private items(list: IListIdentifier, root: GraphFI = this.graph): IListItems {
    return root.sites.getById(this.siteId).lists.getById(list.id).items;
  }

  // The /fields endpoint of one item — Graph writes column values here, not on
  // the item itself. Optionally from a batched root so it joins the batch.
  private fields(list: IListIdentifier, id: string, root: GraphFI = this.graph): IGraphQueryable {
    return GraphQueryable(this.items(list, root).getById(id), 'fields');
  }

  // READ: upcoming sessions with the Speaker lookup expanded, sorted by date.
  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    return this.getItemsNoBatch(list);
  }

  private async getItemsNoBatch(list: IListIdentifier): Promise<IEventItem[]> {
    const itemsQuery = this.items(list).using(InjectHeaders({ Prefer: PREFER_NONINDEXED }));

    itemsQuery.query.set('expand', `fields(select=${GRAPH_FIELD_SELECT})`);
    itemsQuery.query.set('$filter', `fields/SessionDate ge '${startOfTodayIsoNoMs()}'`);
    itemsQuery.query.set('$orderby', 'fields/SessionDate asc');

    const items = await itemsQuery() as IGraphListItem[];
    return items.map(fromGraphItem);
  }

  // CREATE: add a new event item.
  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const result = await this.items(list).add({ fields: toGraphWriteFields(item) }) as IGraphAddResult;
    return { ...item, Id: parseInt(result.id, 10) };
  }

  // UPDATE: PATCH the item's /fields endpoint with the column values.
  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    await graphPatch(this.fields(list, itemId.toString()), body(toGraphWriteFields(item)));
    return { ...item, Id: itemId };
  }

  // DELETE: remove an item by id.
  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.items(list).getById(itemId.toString()).delete();
  }

  // PnPjs-only: a create/update/delete lifecycle, each phase sent as one $batch.
  // Graph caps a $batch at 20 sub-requests, so the request count is split there too.
  public async runBatchDemo(
    list: IListIdentifier,
    batchSize: number,
    onPhase?: (phase: IBatchDemoPhase) => Promise<void> | void
  ): Promise<IBatchDemoResult> {
    const phases: IBatchDemoPhase[] = [];
    const effectiveSize = Math.min(batchSize, GRAPH_BATCH_LIMIT);
    const requestsFor = (ops: number): number => (ops === 0 ? 0 : Math.ceil(ops / effectiveSize));
    const report = async (name: string, label: string, operations: number): Promise<void> => {
      const phase = { name, label, operations, requests: requestsFor(operations) };
      phases.push(phase);
      if (onPhase) await onPhase(phase);
    };

    // PHASE 1 — cleanup: find leftover SAMPLE items (filtered client-side, since
    // Graph column filters are limited), then delete them in one batch.
    const scan = this.items(list).using(InjectHeaders({ Prefer: PREFER_NONINDEXED }));
    scan.query.set('expand', 'fields(select=Title)');
    scan.query.set('$top', '500');
    const rows = await scan() as { id: string; fields?: { Title?: string } }[];
    const staleIds = rows.filter((r) => (r.fields?.Title ?? '').indexOf(SAMPLE_PREFIX) === 0).map((r) => r.id);
    if (staleIds.length) {
      const [batch, execute] = this.graph.batched({ maxRequests: batchSize });
      const ops = staleIds.map((id) => this.items(list, batch).getById(id).delete());
      await execute();
      await Promise.all(ops);
    }
    await report('cleanup', `Deleted ${staleIds.length} leftover ${SAMPLE_PREFIX} item(s)`, staleIds.length);

    // PHASE 2 — create DEMO_COUNT items in one batch (dated today so they show).
    const nowIso = new Date().toISOString();
    const createFields: Record<string, unknown> = { Title: SAMPLE_TITLE, SessionDate: nowIso };
    const [createBatch, execCreate] = this.graph.batched({ maxRequests: batchSize });
    const createOps = Array.from({ length: DEMO_COUNT }, () =>
      this.items(list, createBatch).add({ fields: createFields })
    );
    await execCreate();
    const created = await Promise.all(createOps) as IGraphAddResult[];
    const ids = created.map((r) => r.id);
    await report('create', `Created ${ids.length} items`, ids.length);

    // PHASE 3 — update each item, appending its own id to the title.
    const [updateBatch, execUpdate] = this.graph.batched({ maxRequests: batchSize });
    const updateOps = ids.map((id) =>
      graphPatch(this.fields(list, id, updateBatch), body({ Title: `${SAMPLE_TITLE}${id}` }))
    );
    await execUpdate();
    await Promise.all(updateOps);
    await report('update', `Appended the id to ${ids.length} titles`, ids.length);

    // PHASE 4 — delete the odd-id items in one batch (even ids survive).
    const oddIds = ids.filter((id) => parseInt(id, 10) % 2 === 1);
    if (oddIds.length) {
      const [deleteBatch, execDelete] = this.graph.batched({ maxRequests: batchSize });
      const deleteOps = oddIds.map((id) => this.items(list, deleteBatch).getById(id).delete());
      await execDelete();
      await Promise.all(deleteOps);
    }
    await report('delete-odd', `Deleted ${oddIds.length} odd-id items`, oddIds.length);

    return {
      phases,
      totalOperations: phases.reduce((sum, p) => sum + p.operations, 0),
      totalRequests: phases.reduce((sum, p) => sum + p.requests, 0),
      batchSize
    };
  }
}
