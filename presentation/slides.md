# SharePoint Data Access in SPFx

**A 60-minute talk in two passes: SPFx-native first, then PnPjs**

Presenter: Don Kirkham
Runtime target: 60 min (≈22 slides / ≈32 demo / ≈6 Q&A)

---

## Slide 1 — Title

**Reading and Writing SharePoint Data from SPFx**
*The native way, then the better way*

Don Kirkham · Solution Foundry

Speaker notes:
- Set the room: "By the end you'll know exactly which API to reach for, and why."
- Promise the structure: "I'll teach you the SPFx-native way first. Then we'll go back over the same ground with PnPjs and you'll see why most of us don't write SPFx-native code anymore."

---

## Slide 2 — Why this talk

You have **four** ways to get data into an SPFx web part:

1. SP REST via `SPHttpClient`
2. MS Graph via `MSGraphClientV3`
3. Anonymous external APIs via `HttpClient`
4. PnPjs (`@pnp/sp`, `@pnp/graph`, `@pnp/queryable`) — wraps the first three

Most tutorials pick one and move on. Today we compare them on the same problem.

Speaker notes:
- The "same problem" is a Speaking Events list — Title, Speaker (people picker), Session, SessionDate, SessionType, EventLogo, EventSite, SessionLink.
- Same CRUD operations against the same list. Six implementations.

---

## Slide 3 — The roadmap

**Pass 1 — SPFx native (≈24 min)**
- SP REST against the list
- MS Graph against the same list (and the Graph Explorer sandbox)
- Anonymous calls to a public API

**Pivot (≈2 min)** — *what's missing?*

**Pass 2 — PnPjs (≈23 min)**
- Same SP operations, fluent
- Same Graph operations, fluent
- The free upgrades: logging, caching, batching

**Wrap (≈10 min)** — when to still use SPFx-native, decision cheat sheet, Q&A

---

# PASS 1 — The SPFx Way

---

## Slide 4 — `SPHttpClient`: the mechanics

What it is: a fetch wrapper from `@microsoft/sp-http`. Already authenticated to your SharePoint site.

What you get for free:
- Auth headers (cookies + bearer)
- Form digest on writes
- Sensible defaults (`Accept: application/json;odata=nometadata`)

What you don't get:
- Logging
- Caching
- Batching
- Retry
- Type-safe results

Speaker notes:
- "If you've ever written `fetch` against `/_api/web`, you know the headers dance. SPFx hides that. That's already a win over raw fetch."

---

## Slide 5 — SPFx + SP REST: read

```ts
public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
  const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`
    + `?$select=Id,Title,Session,SessionDate,SessionType,EventSite,SessionLink,`
    +   `Speaker/Id,Speaker/Title,Speaker/EMail`
    + `&$expand=Speaker`
    + `&$filter=${encodeURIComponent(`SessionDate ge datetime'${startOfTodayIso()}'`)}`
    + `&$orderby=SessionDate asc`;
  const response = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) throw new Error(`Failed to get items: ${response.statusText}`);
  const data = await response.json();
  return data.value as IEventItem[];
}
```

This is fine. It works. It's also a string-concatenated URL, a manual cast, and zero observability.

Speaker notes:
- Point out the URL — `$select`, `$expand`, `$filter`, `$orderby`. Memorize this URL shape; it'll matter in 30 minutes.
- Highlight the `as IEventItem[]` cast — TypeScript trusts you. SharePoint doesn't validate.

---

## Slide 6 — SPFx + SP REST: write

```ts
public async updateItem(list, itemId, item): Promise<IEventItem> {
  const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
  const options = {
    headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
    body: JSON.stringify(toSpWritePayload(item))
  };
  const response = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  if (!response.ok) throw new Error(`Failed to update item ${itemId}: ${response.statusText}`);
  return { ...item, Id: itemId };
}
```

Note the ceremony:
- `IF-MATCH: '*'` — opt out of ETag concurrency
- `X-HTTP-Method: MERGE` — because it's a POST pretending to be a PATCH
- A separate `toSpWritePayload(item)` mapper to translate field shapes (Speaker becomes `SpeakerId` array, URL fields become `{ Url, Description }`, …)

[Demo cue → Demo 1]

---

## Slide 7 — When SP REST stops being enough

Reach for **MS Graph** when:
- You need data outside SharePoint (Outlook, Teams, OneDrive, users, groups)
- You need to span sites or tenants without site-collection-scoped tokens
- You want one API instead of three

Caveat: **Graph for SharePoint list items is awkward.** You'll see why in a minute.

Speaker notes:
- This is where the "use Graph for everything" advice falls down. Graph is great for non-SP workloads. For SP lists, REST is often cleaner.

---

## Slide 8 — SPFx + Graph: read list items

```ts
const qs = [
  `expand=fields(select=${GRAPH_FIELD_SELECT})`,            // <-- no $ on expand
  `$filter=fields/SessionDate ge '${startOfTodayIsoNoMs()}'`,// <-- single-quoted ISO
  `$orderby=fields/SessionDate asc`
].join('&');

const response = await this.graphClient
  .api(`/sites/${this.siteId}/lists/${list.id}/items`)
  .version('v1.0')
  .header('Prefer', 'HonorNonIndexedQueriesWarningMayFailRandomly')
  .query(qs)
  .get();
```

Three Graph-specific gotchas already on this slide:
1. `expand=fields(select=...)` — **no `$`**. With `$expand` you get an OData parser error.
2. Datetime literal must be a single-quoted ISO 8601 string. `datetime'...'` returns 400.
3. `$filter` / `$orderby` against `fields/<column>` requires the column to be **indexed** in the list.

Speaker notes:
- The `Prefer` header is a fallback for non-indexed columns — **unreliable**. Indexing is the real fix.
- "If you've never seen these errors, that's because you've been using PnPjs and you didn't know it was hiding them."

---

## Slide 9 — The Graph Explorer detour

Same SPFx `MSGraphClientV3`, no list at all — just `path → JSON`.

```ts
public async runQuery(path: string): Promise<unknown> {
  return await this.graphClient.api(path).version('v1.0').get();
}
```

[Demo cue → Demo 2]

Use this in real life when you're **prototyping** or **exploring** a Graph endpoint before you commit to a service implementation. Cheaper than Postman because the auth is already there.

---

## Slide 10 — SPFx + anonymous APIs

The third HTTP client in `@microsoft/sp-http`: **`HttpClient`**. No auth. Plain fetch, basically.

```ts
const response = await this.httpClient.get(
  'https://official-joke-api.appspot.com/random_joke',
  HttpClient.configurations.v1
);
const joke: IJokeResponse = await response.json();
return { id: joke.id, setup: joke.setup, punchline: joke.punchline };
```

Use this for any external API that doesn't need your tenant's auth.

[Demo cue → Demo 3]

---

## Slide 11 — Pivot: notice anything?

Look at what we wrote across three services:

| | Logging | Caching | Batching | Retry | Types |
|---|---|---|---|---|---|
| `SPHttpClient` | ❌ | ❌ | ❌ | ❌ | cast |
| `MSGraphClientV3` | ❌ | ❌ | ❌ | ❌ | cast |
| `HttpClient` | ❌ | ❌ | ❌ | ❌ | cast |

> Every empty cell is a problem you'd solve yourself.
> PnPjs already solved them.

Speaker notes:
- This is the turn. Pause here. Don't rush.
- "We just wrote three services. Every one of them needs the same five things added before it's production-ready. We're going to write the same three services again with PnPjs, and you're going to watch all five columns light up."

---

# PASS 2 — The PnPjs Way

---

## Slide 12 — What PnPjs is

A **fluent, composable HTTP pipeline** for SharePoint, Graph, and arbitrary endpoints.

Three packages today:
- `@pnp/sp` — SharePoint REST, fluent
- `@pnp/graph` — MS Graph, fluent
- `@pnp/queryable` — the underlying pipeline (also works against any URL)

Init pattern (once per service):

```ts
const sp    = spfi(siteUrl).using(spSPFx(this.context));
const graph = graphfi().using(graphSPFx(this.context));
```

Same SPFx context, same auth. **PnPjs is not a separate auth story.** It rides on top of the SPFx clients you already have.

---

## Slide 13 — The URL reveal

```ts
// SPFx-native
spHttpClient.get(
  `${siteUrl}/_api/web/lists/getbytitle('Speaking Events')/items`
  + `?$select=Id,Title,SessionDate&$orderby=SessionDate desc&$top=50`,
  SPHttpClient.configurations.v1
);

// PnPjs
sp.web.lists.getByTitle('Speaking Events').items
  .select('Id', 'Title', 'SessionDate')
  .orderBy('SessionDate', false)
  .top(50)();
```

Both produce **the same network request**.

> PnPjs is not magic.
> It's the URL you'd write anyway, just typed for you.

[Demo cue → Demo 4]

Speaker notes:
- This is the most important slide in the deck. Stay on it.
- The audience needs to leave this slide trusting that PnPjs isn't doing anything weird underneath. That trust is what makes the next slides land.

---

## Slide 14 — PnPjs + SP REST: read

```ts
public async getItems(list): Promise<IEventItem[]> {
  return await this.sp.web.lists
    .getByTitle(list.title)
    .items
    .select('Id', 'Title', 'Session', 'SessionDate', 'SessionType',
            'EventSite', 'SessionLink',
            'Speaker/Id', 'Speaker/Title', 'Speaker/EMail')
    .expand('Speaker')
    .filter(`SessionDate ge datetime'${startOfTodayIso()}'`)
    .orderBy('SessionDate', true)() as IEventItem[];
}
```

Same fields. Same filter. Same order. **No URL string. No `response.ok` check. No `response.json()`.**

[Demo cue → Demo 5]

---

## Slide 15 — PnPjs + SP REST: write

```ts
public async updateItem(list, itemId, item): Promise<IEventItem> {
  await this.sp.web.lists
    .getByTitle(list.title)
    .items
    .getById(itemId)
    .update(toSpWritePayload(item));
  return { ...item, Id: itemId };
}
```

Compare to SPFx:
- No `IF-MATCH` header.
- No `X-HTTP-Method: MERGE`.
- No URL.

Same `toSpWritePayload` mapper — that's a SharePoint shape problem, not an HTTP-client problem.

---

## Slide 16 — PnPjs + Graph: same operations

```ts
const itemsQuery = this.graph.sites
  .getById(this.siteId)
  .lists
  .getById(list.id)
  .items
  .using(InjectHeaders({ Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly' }));

itemsQuery.query.set('expand', `fields(select=${GRAPH_FIELD_SELECT})`);
itemsQuery.query.set('$filter', `fields/SessionDate ge '${startOfTodayIsoNoMs()}'`);
itemsQuery.query.set('$orderby', 'fields/SessionDate asc');

const items = await itemsQuery() as IGraphListItem[];
```

Honesty check: **the Graph SP-list awkwardness is a Graph problem, not a PnPjs problem.** Both clients have to deal with it. PnPjs at least gives you `InjectHeaders` and a query builder.

[Demo cue → Demo 6]

---

## Slide 17 — PnPjs + anonymous: yes, that works too

```ts
const q = new Queryable('https://official-joke-api.appspot.com/random_joke');
q.using(BrowserFetch(), RejectOnError(), ResolveOnData(), JSONParse());
const joke: IJokeResponse = await q();
```

The same pipeline that talks to SharePoint can talk to **anything**. You just compose the behaviors you need.

(Real-world note: strip `X-PnPjs-RequestId` on external calls to avoid CORS preflight — see the source.)

---

## Slide 18 — The free upgrades

Three toggles. Three real problems gone.

```ts
// 1. Logging — @pnp/logging wrapped once; level driven by a property-pane toggle.
Logger.attachConsoleListener();
Logger.setLevel(this.properties.enhancedLogging ? LogLevel.Verbose : LogLevel.Warning);

// 2. Caching — one line on the SPFI, toggled by un/commenting in ServiceFactory.
const sp = spfi(siteUrl)
  .using(spSPFx(this.context))
  .using(Caching({ store: 'session' })); // <-- the only line that flips

// 3. Batching — pack many reads (or writes) into one $batch envelope.
const [batchedSp, execute] = sp.batched({ maxRequests: 100 });
const pagePromises: Promise<IEventItem[]>[] = [];
for (let i = 0; i < pageCount; i++) {
  pagePromises.push(
    batchedSp.web.lists.getByTitle('Speaking Events').items
      .top(pageSize).skip(i * pageSize)() as Promise<IEventItem[]>
  );
}
await execute(); // one HTTP request to /_api/$batch
```

[Demo cue → Demo 7]

Speaker notes:
- The Caching demo is the visceral one. Click → wait → click → instant. Network tab tells the whole story. Session-store means it even survives a hard refresh.
- The batching demo is the *architectural* one. 100 paginated reads collapse into one HTTP request.

---

## Slide 19 — What logging gets you

```
[DataDemo] onInit: starting (enhancedLogging=on)
[DataDemo] loadItems: requesting items from list Speaking Events
[DataDemo] loadItems: received 12 item(s)        ▶ (12) [{…}, {…}, …]
```

A thin wrapper over `@pnp/logging` (info / debug / warn / error) emits via `console.*` so object payloads render as collapsible trees. Listeners are pluggable — `ConsoleListener` today, App Insights tomorrow. Same call sites.

Speaker notes:
- Tease: "If you're shipping into production without observability on your data calls, this is the easiest win you'll ever take."
- The wrapper isn't a hard requirement — `Logger.subscribe(ConsoleListener())` works out of the box. The wrapper just lets us prefix `[DataDemo]` and route through `console.info`/`console.debug` for the collapsible-tree affordance.

---

## Slide 20 — What caching gets you

Two scopes. Pick one.

```ts
// Global: every read on this SPFI is cached.
const sp = spfi(siteUrl)
  .using(spSPFx(this.context))
  .using(Caching({ store: 'session' })); // 'session' | 'local'

// Per-call: opt one query *out* of caching when global is on.
.items.using(CacheNever())()

// Per-call: opt one query *in* with a custom expiry when global is off.
.items.using(Caching({ store: 'session', expireFunc: () => addMinutes(new Date(), 5) }))()
```

Defaults: 60-second expiry, keyed by URL. `'session'` survives a hard refresh; `'local'` survives a tab close. (Default is in-memory and dies with the page.)

When to use it:
- Reference data (lookups, choice columns, taxonomies)
- Repeated reads in the same session
- Anything you'd hate to re-fetch on tab switch

When **not** to use it:
- Frequently-changing transactional data
- Anything where stale = wrong (use `CacheNever()` on those queries)

---

## Slide 21 — What batching gets you

Reads, in our demo. Writes work the same way.

```ts
const [batchedSp, execute] = this.sp.batched({ maxRequests: 100 });
const pagePromises: Promise<IEventItem[]>[] = [];
for (let i = 0; i < pageCount; i++) {
  pagePromises.push(
    batchedSp.web.lists.getByTitle(list.title).items
      .select('Id', 'Title', 'SessionDate', /* ... */)
      .orderBy('SessionDate', true)
      .top(pageSize).skip(i * pageSize)() as Promise<IEventItem[]>
  );
}
await execute();
const items = (await Promise.all(pagePromises)).flat();
```

100 paginated reads → **one** HTTP request to `/_api/$batch`. The same wiring works for `add`/`update`/`delete`.

Why it matters:
- **Latency**: one round trip beats N. Especially over slow links.
- **Throttling**: SharePoint counts requests, not inner operations. Batching is your throttle escape hatch.
- **Atomicity-ish**: not a real transaction, but closer than N independent calls.

[Demo cue → Demo 7 batching segment]

Speaker notes:
- Be honest: `$batch` is not transactional. If page 3 fails, pages 1 and 2 still came back. PnPjs surfaces per-operation results so you can detect partial failures.
- The demo uses paginated *reads* because the Speaking Events list isn't big enough for a bulk-write story to land. The pattern is identical for writes — same `batched()` call, different inner verbs.

---

# WRAP

---

## Slide 22 — When to still reach for SPFx-native

PnPjs is the default. But it's not free. Reach for **SPFx-native** when:

- **Bundle size matters.** PnPjs adds ~50–80 KB depending on what you import. For a single-purpose tile that does one Graph call, that's a lot.
- **One-off calls** with no list interaction (a single `/me` ping, a webhook trigger).
- **Debugging a network issue** — fewer abstractions between you and the wire.
- **Your team doesn't know PnPjs yet** and the call is trivial.

Default: **PnPjs**. Exceptions: above.

---

## Slide 23 — Decision cheat sheet

| Scenario | Use |
|---|---|
| List CRUD on the current site | PnPjs SP |
| List CRUD across sites (one tenant) | PnPjs Graph |
| Cross-tenant or app-only | PnPjs Graph (with custom auth) |
| One-off anonymous fetch | SPFx `HttpClient` |
| Many anonymous fetches, want consistency | PnPjs `Queryable` |
| Prototyping a Graph endpoint | SPFx `MSGraphClientV3` (or Graph Explorer in browser) |
| Single-purpose perf-critical tile | SPFx native |
| You need logging/caching/batching | PnPjs (every time) |

---

## Slide 24 — Resources

- PnPjs docs: https://pnp.github.io/pnpjs/
- SPFx HTTP clients: https://learn.microsoft.com/sharepoint/dev/spfx/web-parts/basics/connect-to-sharepoint
- Microsoft Graph: https://learn.microsoft.com/graph/
- This demo project: [point at repo URL]

**Three takeaways:**
1. PnPjs makes the same HTTP request, just typed for you.
2. Logging, caching, and batching are one line each.
3. Reach for SPFx-native deliberately, not by default.

---

## Slide 25 — Q&A

Questions?

Speaker notes:
- Likely questions to be ready for:
  - "Does PnPjs work with app-only auth?" → Yes, but not via SPFx behavior — you provide your own `MSAL` token via `using()`.
  - "What about Selectively Disabled Telemetry?" → `X-PnPjs-RequestId` header; strip via `q.on.pre` for external APIs (you saw this in the joke service).
  - "Bundle impact?" → Tree-shakes well. Import the selectives (`@pnp/sp/webs`, `@pnp/sp/lists`, `@pnp/sp/items`) not the barrel.
  - "PnPjs v3 vs v4?" → v4 is current. v3 → v4 was the batching API rewrite (`sp.createBatch()` → `sp.batched()`).
  - "Why not just use Graph for everything?" → You saw the `expand=fields` mess. SP REST is cleaner for SP data.
