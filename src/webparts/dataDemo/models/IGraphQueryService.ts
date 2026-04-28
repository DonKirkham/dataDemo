// ABOUTME: Service contract for arbitrary MS Graph endpoint queries.
// ABOUTME: Distinct from ISpService — calls any Graph path, no SharePoint shape assumed.

export interface IGraphQueryService {
  runQuery(path: string): Promise<unknown>;
}
