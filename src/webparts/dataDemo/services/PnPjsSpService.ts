// ABOUTME: SharePoint CRUD operations using @pnp/sp (PnP JS v4).
// ABOUTME: Mirrors SpfxSpService; adds runBatchDemo to showcase $batch on bulk writes.

import { IEventItem } from '../models/IEventItem';
import { ISpService, IListIdentifier } from '../models/ISpService';
import { IBatchDemoService, IBatchDemoPhase, IBatchDemoResult } from '../models/IBatchDemoService';
import { startOfTodayIso } from '../utilities/dateUtils';
import { toSpWritePayload } from './itemMappers';
import { SPFI } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';

// Columns each read selects — kept identical to SpfxSpService for side-by-side.
const SELECT_FIELDS: string[] = [
  'Id', 'Title', 'Session', 'SessionDate', 'SessionType', 'EventSite', 'SessionLink',
  'Speaker/Id', 'Speaker/Title', 'Speaker/EMail'
];

// Marker title for the items the batch demo creates, so a later run can find and
// clean them up regardless of how the previous run ended.
const SAMPLE_PREFIX = 'SAMPLE:';
const SAMPLE_TITLE = 'SAMPLE: New event. ID: ';
const DEMO_COUNT = 10;

export class PnPjsSpService implements ISpService, IBatchDemoService {
  // Caching, when enabled, is configured on the SPFI instance in ServiceFactory,
  // so the read methods below benefit from it transparently — no code change here.
  constructor(private sp: SPFI) {}

  // READ: one fluent chain returns typed items — no URL, parsing, or status check.
  public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
    return this.getItemsNoBatch(list);
  }

  private async getItemsNoBatch(list: IListIdentifier): Promise<IEventItem[]> {
    const filter = `SessionDate ge datetime'${startOfTodayIso()}'`;
    return await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select(...SELECT_FIELDS)
      //.top(5) // Uncomment this line to show how paging works
      .expand('Speaker')
      .filter(filter)
      .orderBy('SessionDate', true)() as IEventItem[];
  }

  // CREATE: .add(payload) resolves to the new item (v4 has no .data wrapper).
  public async createItem(list: IListIdentifier, item: IEventItem): Promise<IEventItem> {
    const result = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .add(toSpWritePayload(item));
    return result as IEventItem;
  }

  // UPDATE: .getById(id).update(payload) — PnPjs sends the MERGE for you.
  public async updateItem(list: IListIdentifier, itemId: number, item: IEventItem): Promise<IEventItem> {
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .update(toSpWritePayload(item));
    return { ...item, Id: itemId };
  }

  // DELETE: .getById(id).delete().
  public async deleteItem(list: IListIdentifier, itemId: number): Promise<void> {
    await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .getById(itemId)
      .delete();
  }

  // PnPjs-only: a create/update/delete lifecycle, each phase sent as one $batch.
  // Lower the batch size to split a phase across more $batch round-trips.
  public async runBatchDemo(
    list: IListIdentifier,
    batchSize: number,
    onPhase?: (phase: IBatchDemoPhase) => Promise<void> | void
  ): Promise<IBatchDemoResult> {
    const phases: IBatchDemoPhase[] = [];
    const requestsFor = (ops: number): number => (ops === 0 ? 0 : Math.ceil(ops / batchSize));
    const report = async (name: string, label: string, operations: number): Promise<void> => {
      const phase = { name, label, operations, requests: requestsFor(operations) };
      phases.push(phase);
      if (onPhase) await onPhase(phase);
    };

    // PHASE 1 — cleanup: delete every leftover SAMPLE item in one batch.
    const stale = await this.sp.web.lists
      .getByTitle(list.title)
      .items
      .select('Id')
      .filter(`startswith(Title,'${SAMPLE_PREFIX}')`)
      .top(500)() as { Id: number }[];
    if (stale.length) {
      const [batch, execute] = this.sp.batched({ maxRequests: batchSize });
      const ops = stale.map((r) => batch.web.lists.getByTitle(list.title).items.getById(r.Id).delete());
      await execute();
      await Promise.all(ops);
    }
    await report('cleanup', `Deleted ${stale.length} leftover ${SAMPLE_PREFIX} item(s)`, stale.length);

    // PHASE 2 — create DEMO_COUNT items in one batch (dated today so they show).
    const nowIso = new Date().toISOString();
    const [createBatch, execCreate] = this.sp.batched({ maxRequests: batchSize });
    const createOps = Array.from({ length: DEMO_COUNT }, () =>
      createBatch.web.lists.getByTitle(list.title).items.add({ Title: SAMPLE_TITLE, SessionDate: nowIso })
    );
    await execCreate();
    // In PnP JS v4 .add() resolves to the created item directly (no .data wrapper).
    const created = await Promise.all(createOps);
    const ids = created.map((r) => (r as IEventItem).Id as number);
    await report('create', `Created ${ids.length} items`, ids.length);

    // PHASE 3 — update each item, appending its own id to the title.
    const [updateBatch, execUpdate] = this.sp.batched({ maxRequests: batchSize });
    const updateOps = ids.map((id) =>
      updateBatch.web.lists.getByTitle(list.title).items.getById(id).update({ Title: `${SAMPLE_TITLE}${id}` })
    );
    await execUpdate();
    await Promise.all(updateOps);
    await report('update', `Appended the id to ${ids.length} titles`, ids.length);

    // PHASE 4 — delete the odd-id items in one batch (even ids survive).
    const oddIds = ids.filter((id) => id % 2 === 1);
    if (oddIds.length) {
      const [deleteBatch, execDelete] = this.sp.batched({ maxRequests: batchSize });
      const deleteOps = oddIds.map((id) => deleteBatch.web.lists.getByTitle(list.title).items.getById(id).delete());
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
