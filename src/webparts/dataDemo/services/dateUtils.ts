// ABOUTME: Shared date helpers for SharePoint OData queries.
// ABOUTME: SharePoint compares dates in UTC; we anchor "today" at local midnight, then convert.

export function startOfTodayIso(): string {
  const now = new Date();
  const localMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  return localMidnight.toISOString();
}

export function startOfTodayIsoNoMs(): string {
  return startOfTodayIso().replace(/\.\d{3}Z$/, 'Z');
}
