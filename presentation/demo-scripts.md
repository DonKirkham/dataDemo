# Demo Scripts

**Audience-facing notes for each live demo. Read these before going on stage.**

The web part has a two-tier Pivot:
- **Top tier** = transport (SPFx vs PnPjs)
- **Second tier** = endpoint (SharePoint, MS Graph, Anonymous, plus Graph Explorer)

Every demo is "click the right Pivot tab, then do the operation." The code differences are in the underlying service classes; the UI is identical so the audience sees the *behavior*, not new chrome.

---

## Pre-talk setup checklist (do this 30 min before)

- [ ] `heft start` running, web part loaded in workbench or hosted page
- [ ] Network tab open in DevTools, filter set to `XHR/Fetch`
- [ ] Console tab cleared
- [ ] Speaking Events list seeded with ~10 items (mix of dates, some in past, some in future)
- [ ] At least one test item you're willing to delete during the demo
- [ ] VS Code open, side-by-side with browser, with the six service files in tabs:
  - `SpfxSpService.ts`
  - `PnPjsSpService.ts`
  - `SpfxGraphSpService.ts`
  - `PnPjsGraphSpService.ts`
  - `SpfxAnonymousService.ts`
  - `PnPjsAnonymousService.ts`
- [ ] Backup screen recording of each demo on local disk in case the tenant flakes
- [ ] Joke API smoke-tested (`https://official-joke-api.appspot.com/random_joke` returns 200) — it's a free public API and goes down sometimes

---

## Demo 1 — SPFx + SP REST: read and write

**Goal:** show the audience the URL the SPFx code builds, and the headers a write requires.

**Slide cue:** Slide 6 (write code).

**Time budget:** 5 min.

**Steps:**

1. **Set the Pivot tabs.** Click *SPFx* → *SharePoint*.
2. **Open VS Code to `SpfxSpService.ts`.** Highlight lines 17–22 (the URL builder).
   - **Say:** "This is the URL we're sending. Memorize the shape — `$select`, `$expand`, `$filter`, `$orderby`. We'll come back to it."
3. **Switch to browser, Network tab.** Click the **Refresh** / **Load** button on the web part.
4. **Click the request in Network.** Show the URL. Show that it matches the slide.
5. **Show the response.** Point at `value: [...]` and the field shape. Note `Speaker` is an array of `{ Id, Title, EMail }` because we expanded it.
6. **Click Add Event.** Fill in: Title = "Demo Event", SessionDate = a future date, SessionType = "60 minute session", Speaker = yourself.
7. **Before you click Save:** "Watch the network tab. This is going to be a POST with two headers I want you to see."
8. **Click Save.** Find the POST in Network.
9. **Show the request headers.** Point at `IF-MATCH: *` and `X-HTTP-Method: MERGE`. Wait — actually that's the *update*. For the create, just show the POST body.
10. **Click Edit on the new item.** Change the title. Save.
11. **Show the update request.** Now point at `IF-MATCH: *` and `X-HTTP-Method: MERGE`.
    - **Say:** "It's a POST that says 'I'm actually a MERGE.' That's how SharePoint REST does PATCH. The SPFx client doesn't hide that from you."

**Recovery moves:**
- If the list is empty: have a backup screenshot of the network tab.
- If a write fails (digest issue, list locked): pivot to "this is exactly the kind of thing PnPjs error messages help with" and skip ahead.

---

## Demo 2 — SPFx + Graph: read list items, then Graph Explorer

**Goal:** show the Graph URL is *different* from the SP REST URL, and demonstrate the gotchas.

**Slide cue:** Slide 8 (Graph read code), then Slide 9 (Graph Explorer).

**Time budget:** 6 min.

**Steps:**

1. **Click Pivot to *SPFx* → *MS Graph (SP)*.**
2. **Open VS Code to `SpfxGraphSpService.ts`.** Highlight lines 22–32.
   - **Point at line 23:** "`expand=fields(select=...)`. **No dollar sign on expand.** If you put `$expand` here, you get an OData parser error. I learned this the hard way."
   - **Point at line 24:** "Datetime literal. Single-quoted ISO string. The legacy `datetime'...'` wrapper from SP REST? Returns a 400 here."
   - **Point at line 30:** "This `Prefer` header is a fallback for non-indexed columns. **Do not rely on it.** Index your filter and sort columns. The header just delays the failure."
3. **Browser, Network tab, click Load.**
4. **Click the request.** URL is now `graph.microsoft.com/v1.0/sites/{id}/lists/{id}/items`. Different host, different shape.
5. **Show the response.** Point at the nested `fields` object — every column is inside `fields`, not on the root. Hyperlink fields are `{ Description, Url }` objects.
   - **Say:** "Every Graph response for a list item is wrapped in `fields`. Your mapper has to unwrap it. That's why we have a separate `graphMappers.ts`."
6. **Click Pivot to *SPFx* → *MS Graph*.** (This is the Graph Explorer view.)
7. **Type `/me` in the path box.** Click Run.
8. **Show the response.** "This is the same `MSGraphClientV3` we just used for list items. No list, no site, just a path. Use this when you're prototyping."
9. **Try `/me/messages?$top=3`.** Show the result.
10. **Try `/sites/root`.** Show the site object. Point at the `id` — "this is what we feed into the list-items call."

**Recovery moves:**
- If `/me/messages` fails (permissions): use `/me` and `/sites/root` only.
- If Graph throws on the list call: the list might not be indexed. Talk through the error, switch to PnPjs section early to recover the demo arc.

---

## Demo 3 — SPFx + anonymous: the Joke panel

**Goal:** quick palate cleanser. Show that `HttpClient` is just plain fetch with a friendlier API.

**Slide cue:** Slide 10.

**Time budget:** 3 min.

**Steps:**

1. **Click Pivot to *SPFx* → *Anonymous*.**
2. **Open VS Code to `SpfxAnonymousService.ts`.** Show all 30 lines on screen.
   - **Say:** "That's the whole service. Three lines of HTTP, two lines of mapping. No auth. No headers."
3. **Browser, click Get Joke.**
4. **Show the network request.** Point at the URL — `official-joke-api.appspot.com`. Not SharePoint.
5. **Wait for the punchline reveal.** (Use it as a beat.)
6. **Click Get Joke again.** Show that it hits the network every time. (Foreshadowing for caching.)

**Why this matters:**
- **Say:** "Most SPFx tutorials forget this client exists. If you ever need to call a public weather API, a sports scores API, anything not in your tenant — `HttpClient` is the answer. Same package, no auth, you're done."

---

## Demo 4 — The URL reveal (Pass 2 opener)

**Goal:** prove that PnPjs and SPFx-native produce the same network request.

**Slide cue:** Slide 13.

**Time budget:** 4 min.

**Steps:**

1. **Set up split screen:** VS Code on the left with `SpfxSpService.ts` and `PnPjsSpService.ts` open side-by-side. Browser on the right with Network tab.
2. **Click Pivot to *SPFx* → *SharePoint*.** Refresh.
3. **Click the request in Network. Copy the URL to a sticky-note view** (or just leave the request highlighted).
4. **Click Pivot to *PnPjs* → *SharePoint*.** Refresh.
5. **Click the new request in Network.** Show the URL.
6. **Compare side-by-side.** They should be identical (same `$select`, `$expand`, `$filter`, `$orderby`).
   - **Say:** "Same URL. Same method. Same response. The only thing that changed was the code we wrote to *build* the URL. PnPjs is not magic — it's the URL you'd write anyway, just typed for you."
7. **Switch to VS Code.** Show the `getItems` method in `PnPjsSpService.ts` (lines 16–24) next to the SPFx version (lines 16–30).
   - **Count lines aloud:** "13 lines vs 8 lines. No URL string. No `response.ok` check. No `response.json()`. No cast — well, one cast, because we're being paranoid."

**Why this is the most important demo:**
- The audience needs to leave this demo *trusting* that PnPjs isn't doing anything weird. That trust is what makes logging/caching/batching land. If they think it's magic, they think it's risky. It isn't — it's a URL builder.

---

## Demo 5 — PnPjs + SP REST: same operations, less code

**Goal:** reinforce the URL reveal with the write side.

**Slide cue:** Slide 14, 15.

**Time budget:** 4 min.

**Steps:**

1. **Pivot stays at *PnPjs* → *SharePoint*.**
2. **Open VS Code to `PnPjsSpService.ts`.**
3. **Click Add Event.** Same payload as Demo 1 (Title = "PnPjs Demo Event", future date, etc.).
4. **Watch Network tab.** Show the POST.
5. **Show the request body.** It's the same `toSpWritePayload(item)` shape as before — same mapper, both services use it. Point at the import on line 11.
   - **Say:** "The payload is the same because the *list* is the same. The mapper isn't an SPFx-native artifact or a PnPjs artifact — it's a SharePoint shape problem. Both services share it. No duplication."
6. **Click Edit on the new item.** Change something. Save.
7. **Show the network request.** Note: PnPjs may use `PATCH` directly here instead of `POST`+`X-HTTP-Method: MERGE`. Either way, point at how clean it is.
8. **Click Delete on a test item.**
9. **Show the request.** Note: `DELETE` method, no `X-HTTP-Method` gymnastics.

**Compare vs Demo 1:**
- **Say:** "Same network behavior. Half the code. Zero header ceremony. And we haven't even turned on the good stuff yet."

---

## Demo 6 — PnPjs + Graph: same Graph operations

**Goal:** show that `@pnp/graph` smooths the *code* over the awkward Graph URL — but the Graph awkwardness itself doesn't go away.

**Slide cue:** Slide 16.

**Time budget:** 3 min.

**Steps:**

1. **Click Pivot to *PnPjs* → *MS Graph (SP)*.**
2. **Open VS Code to `PnPjsGraphSpService.ts`.** Lines 30–43.
3. **Walk the code:**
   - **Line 35 (`InjectHeaders`):** "This is how PnPjs adds the `Prefer` header. Composable behavior, no manual `.header()` chain."
   - **Lines 37–39 (`itemsQuery.query.set`):** "Honesty time. The PnPjs `.expand()` method has a bug — it splits the comma-separated `fields(select=A,B,C)` argument and produces a malformed URL. So we set the raw query string directly. This is in the codebase as a comment somewhere."
4. **Browser, click Load.**
5. **Show the network request.** Same URL shape as Demo 2. Same response shape with the nested `fields`.
6. **Be honest.**
   - **Say:** "Graph for SP-list items is awkward in *both* clients. PnPjs doesn't fix Graph — Graph is what it is. What PnPjs gives you is `InjectHeaders` and a query builder so the *code* is cleaner even when the URL isn't."

---

## Demo 7 — The free upgrades

**Goal:** make logging, caching, and batching tangible.

**Slide cue:** Slide 18, 19, 20, 21.

**Time budget:** 6 min total. Budget the segments: 1.5 min logging, 2 min caching, 2.5 min batching.

This demo requires three small code changes you'll make live OR pre-stage as a separate "Upgraded" Pivot tab. **Recommendation:** pre-stage the changes in a feature branch and check out the branch before this demo. Live editing fluent chains under time pressure is asking for trouble.

### 7a — Logging (1.5 min)

**Steps:**
1. **Open browser Console tab** (alongside Network).
2. **Show the code change** in VS Code — adding these lines at the top of `PnPjsSpService.ts` constructor (or in `DataDemoWebPart.ts` `onInit`):
   ```ts
   import { Logger, LogLevel, ConsoleListener } from '@pnp/logging';
   Logger.subscribe(ConsoleListener());
   Logger.activeLogLevel = LogLevel.Info;
   ```
3. **Save, let HMR refresh.**
4. **Click Refresh on the web part.**
5. **Show the Console tab.** PnPjs is now logging every request, every response timing, every cache hit/miss.
   - **Say:** "Three lines. Now you have observability. Swap `ConsoleListener` for an App Insights listener and it's production telemetry."

### 7b — Caching (2 min)

**Steps:**
1. **Show the code change:** add `Caching()` to the SPFI:
   ```ts
   import { Caching } from '@pnp/queryable';
   const sp = spfi(siteUrl).using(spSPFx(this.context), Caching());
   ```
2. **Save, let HMR refresh.**
3. **Clear the Network tab.**
4. **Click Refresh on the web part.** Show one request go out.
5. **Click Refresh again immediately.** Show **zero new requests** — and the data is back instantly.
   - **Say (slow it down):** "Watch the Network tab. First click — request goes out. Second click — *nothing*. The data renders instantly because it came from in-memory cache."
6. **Wait 60 seconds (default expiry)** OR open a new tab. Click Refresh. Network request again.
   - **Say:** "60-second default. You can tune it. You can swap to session storage. You can scope it to one query instead of globally."

### 7c — Batching (2.5 min)

This is the one we simulated. Pre-stage the code:

```ts
public async createMany(list: IListIdentifier, items: IEventItem[]): Promise<void> {
  const [batchedSp, execute] = this.sp.batched();
  for (const item of items) {
    batchedSp.web.lists.getByTitle(list.title).items.add(toSpWritePayload(item));
  }
  await execute();
}
```

Wire a "Bulk Add 5 Demo Events" button to call this with 5 throwaway items.

**Steps:**
1. **First, the SPFx-native baseline.** Click Pivot to *SPFx* → *SharePoint*.
2. **Have a "Bulk Add 5" button** that calls `createItem` five times in a Promise.all. (Pre-stage this too.)
3. **Clear Network tab. Click Bulk Add 5.**
4. **Count the requests.** Five separate `POST /_api/web/lists/.../items` requests.
   - **Say:** "Five items, five requests. Five round trips to SharePoint. Each one counts toward your throttling quota."
5. **Click Pivot to *PnPjs* → *SharePoint*.**
6. **Clear Network tab. Click Bulk Add 5.**
7. **Count the requests.** **One** request to `/_api/$batch`.
   - **Say:** "Same five items. One request. One round trip. One throttle hit."
8. **Click the `$batch` request, show the body.** It's a multipart body with five inner POSTs. Show the audience the shape.
   - **Say:** "SharePoint always supported `$batch`. Most code never uses it because building this multipart body by hand is awful. PnPjs builds it for you."

**Honesty beat at the end:**
- **Say:** "One caveat. `$batch` is not a database transaction. If item 3 fails, items 1 and 2 still happened. PnPjs surfaces per-operation results so you can detect partial failures. Plan your retry logic accordingly."

---

## Recovery: what to do if a demo dies

| Failure | Recovery |
|---|---|
| Tenant unreachable | Switch to backup screen recording. Apologize once, move on. |
| List doesn't exist / wrong list | Use the property pane (or refresh the page) to re-select. Have the list ID memorized. |
| Joke API down | Skip Demo 3, mention it in passing during slide 10. |
| Graph permission denied (`/me/messages`) | Stick to `/me` and `/sites/root` in Demo 2. |
| HMR breaks during live code edits in Demo 7 | Don't fight it — fall back to "this is what the code looks like" in VS Code, point at the slide. |
| Caching demo doesn't show instant second click | Check that `Caching()` was added to the SPFI in `ServiceFactory`, not just the local query. Most likely cause: edited the wrong file. |
| Audience asks an aggressive "PnPjs is bloat" question mid-demo | "Let's hold that for the wrap slide — slide 22 has the honest answer." Don't get pulled off the demo arc. |

---

## Timing reminders

If running long at the pivot (Slide 11), cut from the back:
- **First cut:** the anonymous PnPjs demo (Demo 7 doesn't include it; Slide 17 covers it conceptually).
- **Second cut:** Demo 6 (Graph + PnPjs read). The audience already trusts PnPjs by then; the Graph code is in the slides.
- **Last cut:** Q&A buffer down from 6 to 3 min.

**Never cut:**
- Demo 4 (URL reveal). It's the keystone.
- Demo 7 caching segment. It's the most memorable visual in the talk.
- Slide 11 (the pivot). It's the argument.
