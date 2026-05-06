# Enhancing SPFx Projects: Harnessing Live Data Integration

**A 60-minute talk in two passes: SPFx-native first, then PnPjs**

Presenter: Don Kirkham · Solution Foundry
Runtime target: 60 min (≈26 content slides + intro/outro / ≈26 min demos / ≈6 min Q&A)

> **Source of truth:** the live deck is [data.pptx](data.pptx). This file is documentation — it mirrors the deck's slide order and content so a fresh reader can understand the talk without opening PowerPoint. `build-deck.mjs` builds a fresh themed deck (`data-access-spfx.pptx`) from this file; that's a starting point, not the canonical artifact.

---

## Slide 1 — Title

**Enhancing SPFx Projects: Harnessing Live Data Integration**

Don Kirkham, Microsoft MVP

Speaker notes:
- Set the room: "By the end you'll know exactly which API to reach for, and why."
- Promise the structure: "I'll teach you the SPFx-native way first. Then we'll go back over the same ground with PnPjs and you'll see why most of us don't write SPFx-native code anymore."

---

## Slide 2 — *(visual / placeholder)*

---

## Slide 3 — About me

Don Kirkham · Senior Technology Advisor · Arlington, Texas, USA
Microsoft MVP, M365 Development & Microsoft Graph

- https://donkirkham.com
- @DonKirkham
- /in/DonKirkham

---

## Slide 4 — Solution Foundry

**M365 Solutions and Beyond. Built for your needs. By our experts.**

Secure Microsoft 365 modernization, automation, and custom development — delivered by senior consultants, developers, and Microsoft MVPs.

Core services:
- SPFx Development
- SharePoint Modernization
- Power Platform (Power BI, Power Automate, Power Apps)
- AI & Copilot Integration

Federal & regulated program support; enterprise portals, dashboards, compliance workflows, knowledge platforms.

www.solutionfoundry.com

---

## Slide 5 — Picture time

*(audience photo / interactive moment)*

---

## Slide 6 — Why are we here?

**DATA IS KING!**

Speaker notes:
- The "same problem" we'll work on all hour is a Speaking Events list — Title, Speaker (people picker), Session, SessionDate, SessionType, EventLogo, EventSite, SessionLink.
- Same CRUD operations against the same list. Six implementations.

---

## Slide 7 — Why are we here? (the menu)

You have several ways to get data into an SPFx web solution:

- **External APIs** via `HttpClient` (Anonymous and Authenticated)
- **SP REST** via `SPHttpClient`
- **MS Graph** via `MSGraphClientV3`
- **Entra ID Authenticated APIs** via `AadHttpClient`
- **PnPjs** — wraps the calls above
  - `@pnp/queryable`
  - `@pnp/sp`
  - `@pnp/graph`
  - `@pnp/adalclient`

Speaker notes:
- Same Speaking Events list, six implementations. We're going to compare them on the same problem.

---

## Slide 8 — Why are we here? (the promise)

Same menu, plus the framing:

> Most tutorials pick one and move on. **Today we compare them on the same problem.**

[Demo cue → Demo 1 — web part orientation tour]

Speaker notes:
- This is where Demo 1 fires: spin up the web part, show the two-tier Pivot, walk the audience through the chrome they'll see for the rest of the hour. Top tier = transport (SPFx vs PnPjs); second tier = endpoint (Anonymous, SharePoint, MS Graph (SP), Graph Explorer). One screen shape, six implementations underneath.

---

## Slide 9 — The roadmap

**Pass 1 — SPFx native**
- Anonymous calls to a public API
- SP REST against the list
- MS Graph against the same list (and Graph Explorer)

**Pivot — what's missing? What are the pain points?**

**Pass 2 — PnPjs**
- Same anonymous calls to a public API
- Same SP operations, fluent
- Same Graph operations, fluent
- The free upgrades

**Wrap** — when to still use SPFx-native, decision cheat sheet, Q&A

---

## Slide 10 — Intro: Anatomy of a REST call

*(section divider)*

---

## Slide 11 — Client-side data interaction: REST APIs

**RE**presentational **S**tate **T**ransfer — the protocol SharePoint, Graph, and most web APIs speak.

```
https://mydomain.sharepoint.com/_api/web/lists?$select=Title&$filter=Title eq 'Events'
↑ Host                          ↑ Endpoint   ↑ OData query parameters
```

**HTTP methods:** GET (read) · POST (create) · PATCH/MERGE (update) · PUT (replace) · DELETE (remove) · OPTIONS/HEAD (metadata)

**Headers:** Authentication (cookies, bearer token) · Content-Type · Accept · Language

**Body** (POST / PATCH):
```json
{ "Title": "New List", "Description": "This is my new list for stuff" }
```

---

## Slide 12 — OData query operators

```
https://mydomain.sharepoint.com/sites/MySite/_api/web/lists/getbytitle('MyList')/items
  ?$select=Id,Title,Description,EventDate
  &$filter=EventDate gt DateTime'2024-01-01T00:00:00'
  &$orderby=EventDate
  &$top=10
```

- **`$select`** — only return specific fields
- **`$filter`** — only return matching items
- **`$orderby`** — sort by a field, ascending or descending
- **`$top`** — limit results to N items

Speaker notes:
- This is the URL shape we'll see again in every Pass 1 demo. PnPjs builds the same shape — that's the point of the URL reveal in Pass 2.

---

# PASS 1 — The SPFx Way

## Slide 13 — Pass 1 divider

*(section divider: "PASS 1 — The SPFx Way")*

---

## Slide 14 — SPFx + anonymous APIs

The basic HTTP client in `@microsoft/sp-http`: **`HttpClient`**. No auth. Plain fetch.

```ts
const response = await this.context.httpClient.get(
  'https://official-joke-api.appspot.com/random_joke',
  HttpClient.configurations.v1
);
const joke: IJokeResponse = await response.json();
return { id: joke.id, setup: joke.setup, punchline: joke.punchline };
```

Use this for any external API that doesn't need your tenant's auth.

[Demo cue → Demo 2 — SPFx + Anonymous]

---

## Slide 15 — `SPHttpClient`: the mechanics

A fetch wrapper from `@microsoft/sp-http`, already authenticated to your site.

**What you get for free:**
- Auth headers (cookies + bearer)
- OData v4
- Form digest on writes
- Header defaults (`Accept: application/json; odata=nometadata`)
- Methods for POST operations (merge/update)
- Version-matching commands

**What you don't:**
- Logging
- Caching
- Batching
- Retry
- Type-safe results

Speaker notes:
- "If you've ever written `fetch` against `/_api/web`, you know the headers dance. SPFx hides that. That's already a win over raw fetch."

---

## Slide 16 — SPFx + SP REST: read

```ts
public async getItems(list: IListIdentifier): Promise<IEventItem[]> {
  const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`
    + `?$select=Id,Title,Session,SessionDate,SessionType,EventSite,SessionLink,`
    +   `Speaker/Id,Speaker/Title,Speaker/EMail`
    + `&$expand=Speaker`
    + `&$filter=${encodeURIComponent(`SessionDate ge datetime'${startOfTodayIso()}'`)}`
    + `&$orderby=SessionDate asc`;
  const response = await this.context.spHttpClient.get(url, SPHttpClient.configurations.v1);
  if (!response.ok) throw new Error(`Failed to get items: ${response.statusText}`);
  const data = await response.json();
  return data.value as IEventItem[];
}
```

It works. It's also a string-concatenated URL, a manual cast, and zero observability.

Speaker notes:
- Point at the URL shape — `$select`, `$expand`, `$filter`, `$orderby`. Memorize it; we come back to it after the pivot.
- Highlight the `as IEventItem[]` cast — TypeScript trusts you. SharePoint doesn't validate.

---

## Slide 17 — SPFx + SP REST: update

```ts
public async updateItem(list, itemId, item): Promise<IEventItem> {
  const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
  const options = {
    headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'MERGE' },
    body: JSON.stringify(toSpWritePayload(item))
  };
  const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  if (!response.ok) throw new Error(`Failed to update item ${itemId}: ${response.statusText}`);
  return { ...item, Id: itemId };
}
```

Note the ceremony:
- `X-HTTP-Method: MERGE` — because it's a POST pretending to be a PATCH
- `IF-MATCH: '*'` — opt out of ETag concurrency
- A separate `toSpWritePayload(item)` mapper for SharePoint field shapes

---

## Slide 18 — SPFx + SP REST: add (bonus)

**Bonus — the cleanest verb in the SP REST set.**

```ts
public async createItem(list, item): Promise<IEventItem> {
  const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items`;
  const options = { body: JSON.stringify(toSpWritePayload(item)) };
  const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  if (!response.ok) throw new Error(`Failed to create item: ${response.statusText}`);
  return await response.json() as IEventItem;
}
```

A real POST to the list endpoint. No `IF-MATCH`. No `X-HTTP-Method` override. The body is just your fields.

> And no, you don't need `ListItemEntityTypeFullName`. That was an `odata=verbose` thing. SPFx defaults to `nometadata`.

Speaker notes:
- Pre-empt the audience question: "Don't I need to send `__metadata: { type: 'SP.Data.…ListItem' }` on create?" Only under `odata=verbose`. SPFx ships `nometadata` by default; PnPjs `.items.add()` doesn't include it either. Multi-content-type lists are the edge case — set `ContentTypeId` in the payload, not `__metadata`.

---

## Slide 19 — SPFx + SP REST: delete (bonus)

**Bonus — another POST in disguise.**

```ts
public async deleteItem(list, itemId): Promise<void> {
  const url = `${this.siteUrl}/_api/web/lists/getbytitle('${list.title}')/items(${itemId})`;
  const options = { headers: { 'IF-MATCH': '*', 'X-HTTP-Method': 'DELETE' } };
  const response = await this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, options);
  if (!response.ok) throw new Error(`Failed to delete item ${itemId}: ${response.statusText}`);
}
```

Note the ceremony:
- It's a POST. The only thing that turns it into a delete is `X-HTTP-Method: DELETE` in the headers.
- Same `IF-MATCH: '*'` opt-out as update.

> Why isn't this `spHttpClient.delete()`? Because the SP REST endpoint expects POST with the override header. SPFx mirrors what the wire wants.

[Demo cue → Demo 3 — SPFx + SP REST CRUD]

Speaker notes:
- POST + header override is the SP REST way of expressing every non-GET, non-create verb. Three of the four CRUD verbs go through POST. That's the punchline that primes Pass 2's "no header gymnastics."

---

## Slide 20 — When SP REST stops being enough

Reach for **MS Graph** when:
- You need data outside SharePoint (Outlook, Teams, OneDrive, users, groups)
- You need to span sites or tenants without site-collection-scoped tokens
- You want one API instead of three

Caveat: **Graph for SharePoint list items is awkward.** You'll see why in a minute.

Speaker notes:
- The "use Graph for everything" advice falls down here. Graph is great for non-SP workloads. For SP lists, REST is often cleaner.

---

## Slide 21 — The Graph Explorer detour

Same SPFx `MSGraphClientV3`, no list at all — just `path → JSON`.

```ts
public async runQuery(path: string): Promise<unknown> {
  return await this.graphClient.api(path).version('v1.0').get();
}
```

Use this in real life when **prototyping** or **exploring** a Graph endpoint before you commit to a service implementation. Cheaper than Postman because the auth is already there.

---

## Slide 22 — SPFx + Graph: read

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

## Slide 23 — SPFx + Graph: add (bonus)

**Bonus — a real POST, plus the fields envelope.**

```ts
public async createItem(list, item): Promise<IEventItem> {
  const response = await this.graphClient
    .api(`/sites/${this.siteId}/lists/${list.id}/items`)
    .version('v1.0')
    .post({ fields: toGraphWriteFields(item) });
  return fromGraphItem(response as IGraphListItem);
}
```

A real `.post()` — no method-override gymnastics, no `IF-MATCH`.

But the `{ fields: … }` envelope is mandatory:
- Graph won't accept a flat field bag — it wants the wrapper.
- `toGraphWriteFields` does the same translations as the SP REST mapper (Speaker → `SpeakerLookupId`, URL → `{ Url, Description }`).
- Mapping ceremony is a SharePoint shape problem, not an HTTP-client problem.

Speaker notes:
- The honesty beat: Graph's verbs are nicer than SP REST's, but the field-shape mapping is unchanged.

---

## Slide 24 — SPFx + Graph: delete (bonus)

**Bonus — the cleanest delete in the deck.**

```ts
public async deleteItem(list, itemId): Promise<void> {
  await this.graphClient
    .api(`/sites/${this.siteId}/lists/${list.id}/items/${itemId}`)
    .version('v1.0')
    .delete();
}
```

A real `.delete()`. No POST. No `X-HTTP-Method`. No `IF-MATCH`.

> The SP REST workaround is REST-the-protocol's fault, not Graph's.

[Demo cue → Demo 4 — SPFx + Graph CRUD + Explorer]

Speaker notes:
- If asked about soft-delete vs hard-delete: Graph deletes are hard. If you want recycle-bin behavior, SP REST has `recycle()` — Graph does not.

---

## Slide 25 — Pivot: notice anything?

Look at what we wrote across three services:

| | Logging | Caching | Batching | Retry | Types |
|---|---|---|---|---|---|
| `HttpClient` | ❌ | ❌ | ❌ | ❌ | cast |
| `SPHttpClient` | ❌ | ❌ | ❌ | ❌ | cast |
| `MSGraphClientV3` | ❌ | ❌ | ❌ | ❌ | cast |

> Every empty cell is a problem you'd solve yourself.
> PnPjs already solved them.

Speaker notes:
- This is the turn. Pause here. Don't rush.
- "We just wrote three services. Every one of them needs the same five things added before it's production-ready. We're going to write the same three services again with PnPjs, and you're going to watch all five columns light up."
- **Disambiguation if asked / if visible:** the `[DataDemo]` lines in the console aren't the data-layer logging — those are component-level `Logger.info` calls announcing user actions (transport changed, loadItems requested, etc.). The empty-cell claim is about *automatic request/response logging at the data layer*, which we'll wire up for free in the PnPjs pass. Keep DevTools focused on Network during Pass 1.

---

# PASS 2 — The PnPjs Way

## Slide 26 — Pass 2 divider

*(section divider: "PASS 2 — The PnPjs Way")*

---

## Slide 27 — What PnPjs is

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

Same SPFx context, same auth. **PnPjs is not a separate auth story** — it rides on top of the SPFx clients you already have.

---

## Slide 28 — The URL reveal

```ts
// SPFx-native
spHttpClient.get(
  `${siteUrl}/_api/web/lists/`
  + `getbytitle('Speaking Events')/items`
  + `?$select=Id,Title,SessionDate&`
  + `$orderby=SessionDate desc&$top=50`,
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

Speaker notes:
- This is the most important slide in the deck. Stay on it.
- The audience needs to leave this slide trusting that PnPjs isn't doing anything weird underneath. That trust is what makes the next slides land.
- The URL reveal is demoed live as the prelude to Demo 5 (slide 29's cue) — same Network-tab comparison the original talk used.

---

## Slide 29 — PnPjs + anonymous

```ts
const q = new Queryable('https://official-joke-api.appspot.com/random_joke');
q.using(BrowserFetch(), RejectOnError(), ResolveOnData(), JSONParse());
const joke: IJokeResponse = await q();
```

The same pipeline that talks to SharePoint can talk to **anything**. You just compose the behaviors you need.

> Real-world note: strip `X-PnPjs-RequestId` on external calls to avoid CORS preflight — see the source.

[Demo cue → Demo 5 — URL reveal + PnPjs Anonymous]

---

## Slide 30 — PnPjs + SP REST: read

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

---

## Slide 31 — PnPjs + Graph: read

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

**Honesty check:** the Graph SP-list awkwardness is a Graph problem, not a PnPjs problem. Both clients have to deal with it. PnPjs at least gives you `InjectHeaders` and a query builder.

Speaker notes:
- Graph read is a code-walkthrough beat in Pass 2 (no separate live demo) — by here the audience already trusts the URL reveal. The same gotchas from slide 22 still apply.

---

## Slide 32 — PnPjs + SP REST: update

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

## Slide 33 — PnPjs: add (bonus)

**Bonus — same `.add()` across SP and Graph.**

```ts
// PnPjs SP
const result = await this.sp.web.lists
  .getByTitle(list.title)
  .items
  .add(toSpWritePayload(item));

// PnPjs Graph
await this.graph.sites
  .getById(this.siteId)
  .lists.getById(list.id)
  .items
  .add({ fields: toGraphWriteFields(item) });
```

Same `.add()` verb across **both** APIs. The only difference is the path-builder chain — and that's the API surface, not the HTTP layer.

> Note: the Graph version still needs the `{ fields: … }` envelope — PnPjs hides transport ceremony, not API shape.

Speaker notes:
- Pairs with slide 18 (SPFx SP add) and slide 23 (SPFx Graph add) for the "compare add across all three" beat.

---

## Slide 34 — PnPjs: delete (bonus)

**Bonus — same `.delete()` across SP and Graph.**

```ts
// PnPjs SP
await this.sp.web.lists
  .getByTitle(list.title)
  .items
  .getById(itemId)
  .delete();

// PnPjs Graph
await this.graph.sites
  .getById(this.siteId)
  .lists.getById(list.id)
  .items
  .getById(itemId.toString())
  .delete();
```

No URLs. No `IF-MATCH`. No `X-HTTP-Method`.

> PnPjs normalizes the verb across SP REST (which fakes DELETE via POST) and Graph (which uses real DELETE). Same call site, different wire format.

[Demo cue → Demo 6 — PnPjs + SP REST CRUD]

Speaker notes:
- Tease back to slides 19 and 24: "Remember the SP REST `X-HTTP-Method: DELETE` trick? PnPjs handles that for you. The Graph version drops to a real DELETE underneath — same call site, different wire format."

---

## Slide 35 — Free Upgrades: LOGGING

A thin wrapper over `@pnp/logging` (info / debug / warn / error) emits via `console.*` so object payloads render as collapsible trees. Listeners are pluggable — `ConsoleListener` today, App Insights tomorrow. Same call sites.

```ts
Logger.attachConsoleListener();
Logger.setLevel(this.properties.enhancedLogging ? LogLevel.Verbose : LogLevel.Warning);
```

```
[DataDemo] onInit: starting (enhancedLogging=on)
[DataDemo] loadItems: requesting items from list Speaking Events
[DataDemo] loadItems: received 12 item(s)        ▶ (12) [{…}, {…}, …]
```

> You are hopefully doing it anyway. Let PnPjs supercharge it.

Speaker notes:
- Tease: "If you're shipping into production without observability on your data calls, this is the easiest win you'll ever take."
- The wrapper isn't a hard requirement — `Logger.subscribe(ConsoleListener())` works out of the box. The wrapper just lets us prefix `[DataDemo]` and route through `console.info`/`console.debug` for the collapsible-tree affordance.

---

## Slide 36 — Free Upgrades: CACHING

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

**Use it for:** reference data · repeated reads · anything you'd hate to re-fetch on tab switch
**Avoid for:** frequently-changing transactional data · anything where stale = wrong (use `CacheNever()` there)

---

## Slide 37 — Free Upgrades: BATCHING

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
- **Throttling**: SharePoint counts requests, not inner operations. Your throttle escape hatch.
- **Atomicity-ish**: not a real transaction, but closer than N independent calls.

[Demo cue → Demo 7 — Free upgrades (logging, caching, batching)]

Speaker notes:
- Be honest: `$batch` is not transactional. If page 3 fails, pages 1 and 2 still came back. PnPjs surfaces per-operation results so you can detect partial failures.
- The demo uses paginated *reads* because the Speaking Events list isn't big enough for a bulk-write story to land. The pattern is identical for writes.

---

# WRAP

## Slide 38 — Wrap divider

*(section divider: "WRAP — When and Why")*

---

## Slide 39 — When to still reach for SPFx-native

PnPjs is the default. But it's not free. Reach for **SPFx-native** when:

- **Bundle size matters.** PnPjs adds ~50–80 KB depending on imports. For a single-purpose tile that does one Graph call, that's a lot.
- **One-off calls** with no list interaction (a single `/me` ping, a webhook trigger).
- **Debugging a network issue** — fewer abstractions between you and the wire.
- **Your team doesn't know PnPjs yet** and the call is trivial.

> Default: **PnPjs**. Exceptions: above.

---

## Slide 40 — Decision cheat sheet

| Scenario | Use |
|---|---|
| List CRUD on the current site | PnPjs SP |
| List CRUD across sites (one tenant) | PnPjs Graph |
| Cross-tenant or app-only | PnPjs Graph (with custom auth) |
| One-off anonymous fetch | SPFx `HttpClient` |
| Many anonymous fetches, want consistency | PnPjs `Queryable` |
| Prototyping a Graph endpoint | SPFx `MSGraphClientV3` (or Graph Explorer) |
| Single-purpose perf-critical tile | SPFx native |
| You need logging/caching/batching | PnPjs (every time) |

---

## Slide 41 — Three takeaways + Resources

**Three takeaways:**
1. PnPjs makes the same HTTP request, just typed for you.
2. Logging, caching, and batching are one toggle each.
3. Reach for SPFx-native deliberately, not by default.

**Resources:**
- PnPjs docs — https://pnp.github.io/pnpjs/
- SPFx HTTP clients — https://learn.microsoft.com/sharepoint/dev/spfx/web-parts/basics/connect-to-sharepoint
- Microsoft Graph — https://learn.microsoft.com/graph/
- This demo project — https://github.com/DonKirkham/DataDemo

---

## Slide 42 — Q&A

**Questions?**

Speaker notes:
- Likely questions to be ready for:
  - "Does PnPjs work with app-only auth?" → Yes, but not via SPFx behavior — you provide your own MSAL token via `using()`.
  - "Bundle impact?" → Tree-shakes well. Import the selectives (`@pnp/sp/webs`, `@pnp/sp/lists`, `@pnp/sp/items`) not the barrel.
  - "PnPjs v3 vs v4?" → v4 is current. v3 → v4 was the batching API rewrite (`sp.createBatch()` → `sp.batched()`).
  - "Why not just use Graph for everything?" → You saw the `expand=fields` mess. SP REST is cleaner for SP data.

---

## Slide 43 — Thank you

**THANK YOU, YOU ARE AWESOME** ❤️

Please rate this session in the mobile app.

- https://donkirkham.com
- @DonKirkham
- /in/DonKirkham

---

## Demo cue summary

Seven live-demo cues are placed in the deck. The detailed steps live in [demo-scripts.md](demo-scripts.md):

| Cue | Slide | Demo |
|---|---|---|
| Demo 1 | 8 | Web part orientation tour |
| Demo 2 | 14 | SPFx + Anonymous (Joke panel) |
| Demo 3 | 19 | SPFx + SP REST CRUD |
| Demo 4 | 24 | SPFx + Graph CRUD + Graph Explorer |
| Demo 5 | 29 | URL reveal + PnPjs Anonymous |
| Demo 6 | 34 | PnPjs + SP REST CRUD |
| Demo 7 | 37 | Free upgrades (logging, caching, batching) |
