# Demo Scripts

**Audience-facing notes for each live demo. Read these before going on stage.**

The web part has a two-tier Pivot:
- **Top tier** = transport (SPFx vs PnPjs)
- **Second tier** = endpoint (Anonymous, SharePoint, MS Graph (SP), Graph Explorer) — Anonymous leads to match the deck flow

Every demo is "click the right Pivot tab, then do the operation." The code differences are in the underlying service classes; the UI is identical so the audience sees the *behavior*, not new chrome.

The deck (`data.pptx`) leads with Anonymous in both passes — Demo 1 is the Joke panel, not SP REST. That ordering is intentional: the simplest HTTP client first, so the audience sees the basic shape before SP REST adds ceremony.

---

## Pre-talk setup checklist (do this 30 min before)

- [ ] Laptop setup
  - Desktop 1
    - [ ] Slide deck 
  - Desktop 2
    - [ ] VS Code with the codebase loaded, ready to go 
  - Desktop 3 (if you use a third for notes/timer)
    - [ ] Browser 1 with Dev site loaded for background
    - [ ] Debug browser with DevTools loaded in lower pane
- [ ] `heft start` running, web part loaded in workbench or hosted page
- [ ] Network tab open in DevTools, filter set to `XHR/Fetch`
- [ ] Console tab cleared
- [ ] **Endpoint Pivot defaults to Anonymous** (matches deck flow — verify, don't assume)
- [ ] Speaking Events list seeded with ~10 items (mix of dates, some past, some future)
- [ ] At least one test item you're willing to delete during the demo
- [ ] VS Code open, side-by-side with browser, with the six service files in tabs:
  - `SpfxAnonymousService.ts`
  - `SpfxSpService.ts`
  - `SpfxGraphSpService.ts`
  - `PnPjsAnonymousService.ts`
  - `PnPjsSpService.ts`
  - `PnPjsGraphSpService.ts`
- [ ] Backup screen recording of each demo on local disk in case the tenant flakes
- [ ] **Joke API smoke-tested** (`https://official-joke-api.appspot.com/random_joke` returns 200) — it's a free public API and goes down sometimes. **This is now the opener** — if it's flaky, swap order: open with SP REST and pull Anonymous to last. Don't discover this on stage.
- [ ] **Demo 8 baseline state confirmed:**
  - [ ] In `ServiceFactory.ts`, the `.using(Caching({ store: 'session' }))` line is **commented out** (caching off until 8b)
  - [ ] In `PnPjsSpService.getItems`, the **single-call** return (above the BATCHED block) is **active**; the BATCHED block below is **commented out**
  - [ ] Browser sessionStorage cleared (DevTools → Application → Storage → Clear site data) so 8b shows a real cache miss → hit progression

---

## Demo 1 — SPFx + anonymous: the Joke panel

**Goal:** open the talk with the simplest HTTP client. Show that `HttpClient` is just plain fetch with a friendlier API. Set the rhythm of "click a tab, watch the network."

**Slide cue:** Slide 14.

**Time budget:** 3 min.

**Steps:**

1. **Set the Pivot tabs.** Click *SPFx* → *Anonymous*.
2. **Open VS Code to `SpfxAnonymousService.ts`.** Show all ~30 lines on screen.
   - **Say:** "That's the whole service. Three lines of HTTP, two lines of mapping. No auth. No headers. This is the simplest of the three SPFx clients, so it's where we'll start."
3. **Browser, click Get Joke.**
4. **Show the network request.** Point at the URL — `official-joke-api.appspot.com`. Not SharePoint. Not Graph. Just plain HTTPS.
5. **Wait for the punchline reveal.** (Use it as a beat — laughter buys you setup time.)
6. **Click Get Joke again.** Show that it hits the network every time. (Foreshadowing for caching in Demo 8.)

**Why this matters:**
- **Say:** "Most SPFx tutorials forget this client exists. If you ever need to call a public weather API, a sports scores API, anything not in your tenant — `HttpClient` is the answer. Same `@microsoft/sp-http` package, no auth, you're done. Now let's add some auth and see how it gets messier."

**Recovery moves:**
- If the Joke API is down: this is the opener — pivot fast. "The API's down, which is why we don't write production code against random public services without a fallback." Skip to Demo 2 (SP REST). Re-thread the Anonymous beat verbally on the way to slide 15.

---

## Demo 2 — SPFx + SP REST: read, update, add, delete

**Goal:** show the URL the SPFx code builds, the headers a write requires, and the ceremony around update/delete.

**Slide cue:** Slide 16 (read), Slide 17 (update). Slides 18 and 19 (bonus add/delete) ride along if you have time.

**Time budget:** 6 min.

**Steps:**

1. **Set the Pivot tabs.** Click *SPFx* → *SharePoint*.
2. **Open VS Code to `SpfxSpService.ts`.** Highlight lines 17–22 (the URL builder).
   - **Say:** "This is the URL we're sending. Memorize the shape — `$select`, `$expand`, `$filter`, `$orderby`. We'll come back to it after the pivot."
3. **Switch to browser, Network tab.** Click the **Refresh** / **Load** button on the web part.
4. **Click the request in Network.** Show the URL. Show that it matches the slide.
5. **Show the response.** Point at `value: [...]` and the field shape. Note `Speaker` is an array of `{ Id, Title, EMail }` because we expanded it.
6. **Click Add Event.** Fill in: Title = "Demo Event", SessionDate = a future date, SessionType = "60 minute session", Speaker = yourself.
7. **Click Save.** Find the POST in Network. Show the body — just your fields. (Slide 18 is the matching code if you want to flip to it.)
   - **Say:** "Add is the cleanest verb in SP REST. A real POST. No header gymnastics yet."
8. **Click Edit on the new item.** Change the title. Save.
9. **Show the update request.** Point at `IF-MATCH: *` and `X-HTTP-Method: MERGE` in request headers.
   - **Say:** "It's a POST that says 'I'm actually a MERGE.' That's how SharePoint REST does PATCH. The SPFx client doesn't hide that from you — slide 17 has the code."
10. **(Bonus, if time)** Click Delete on a test item. Show the network request — another POST, this time with `X-HTTP-Method: DELETE`. Slide 19 has the code.
    - **Say:** "Delete is *another* POST. The SP REST endpoint expects POST + override header for everything that isn't GET or create. Three of the four CRUD verbs go through POST."

**Recovery moves:**
- If the list is empty: have a backup screenshot of the network tab.
- If a write fails (digest issue, list locked): pivot to "this is exactly the kind of thing PnPjs error messages help with" and skip ahead to the Graph section.

---

## Demo 3 — SPFx + Graph: read, then Graph Explorer (and bonus add/delete)

**Goal:** show the Graph URL is *different* from the SP REST URL, demonstrate the gotchas, then use the Graph Explorer detour as a tool you'd actually reach for in real life.

**Slide cue:** Slide 21 (Graph Explorer detour), Slide 22 (Graph read code). Slides 23–24 if doing bonus add/delete.

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
6. **(Bonus, if time)** Click Add Event, save. Show the POST. The body has the `{ fields: {...} }` envelope (slide 23). Click Delete on a test item — a real `DELETE` verb (slide 24).
   - **Say:** "Graph's verbs are nicer than SP REST's. Real POST, real DELETE. But the field-shape mapping is unchanged — that's a SharePoint problem, not an HTTP-client problem."
7. **Click Pivot to *SPFx* → *Graph Explorer*.**
8. **Type `/me` in the path box.** Click Run.
9. **Show the response.** "This is the same `MSGraphClientV3` we just used for list items. No list, no site, just a path. Use this when you're prototyping."
10. **Try `/me/messages?$top=3`.** Show the result.
11. **Try `/sites/root`.** Show the site object. Point at the `id` — "this is what we feed into the list-items call we just saw."

**Recovery moves:**
- If `/me/messages` fails (permissions): use `/me` and `/sites/root` only.
- If Graph throws on the list call: the list might not be indexed. Talk through the error, switch to PnPjs section early to recover the demo arc.

---

## Demo 4 — The URL reveal (Pass 2 opener)

**Goal:** prove that PnPjs and SPFx-native produce the same network request.

**Slide cue:** Slide 28.

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

## Demo 5 — PnPjs + anonymous: same Queryable, different endpoint

**Goal:** show that the PnPjs pipeline isn't SharePoint-only. Same `Queryable` composes against any URL — including the Joke API the audience just saw via SPFx `HttpClient`.

**Slide cue:** Slide 29.

**Time budget:** 2 min.

**Steps:**

1. **Click Pivot to *PnPjs* → *Anonymous*.**
2. **Open VS Code to `PnPjsAnonymousService.ts`.** Show the `Queryable` construction.
   - **Point at the `using()` chain:** "`BrowserFetch`, `RejectOnError`, `ResolveOnData`, `JSONParse`. Each one is a behavior. You compose only what you need. No SharePoint context. No auth."
3. **Browser, click Get Joke.** Show the network request — same URL, same `official-joke-api.appspot.com` host as Demo 1.
   - **Say:** "Same external API. Same network request. Different code path. The PnPjs pipeline is just an HTTP pipeline — `@pnp/sp` and `@pnp/graph` are pre-composed bundles of behaviors aimed at SharePoint and Graph. `Queryable` is the raw thing they're built on, and it'll talk to anything."
4. **Mention the CORS gotcha (briefly):** "One real-world note — `@pnp/sp` adds an `X-PnPjs-RequestId` header for telemetry. On external APIs that triggers a CORS preflight. The service strips it on the way out. See the source if you're curious."

**Why this matters:**
- Sets up the framing for slide 35 onward: "logging/caching/batching are pipeline behaviors. Once you know the pipeline composes, you stop seeing them as features and start seeing them as middleware you opt into."

**Recovery moves:**
- If the Joke API is down here too (it died for Demo 1 already): cut this demo entirely, point at slide 29's code. "PnPjs talks to the same endpoint via `Queryable` — same network request as Demo 1, different pipeline composition."

---

## Demo 6 — PnPjs + SP REST: read, update, add, delete

**Goal:** reinforce the URL reveal with the write side. Same operations as Demo 2, half the code, no header ceremony.

**Slide cue:** Slide 30 (read), Slide 32 (update). Slides 33 and 34 (bonus add/delete) ride along.

**Time budget:** 4 min.

**Steps:**

1. **Pivot stays at *PnPjs* → *SharePoint*** (carried over from Demo 4).
2. **Open VS Code to `PnPjsSpService.ts`.** Show the `getItems` chain (lines 16–24).
   - **Say:** "Same `$select`, `$expand`, `$filter`, `$orderby` from the URL we just built by hand. Now it's typed."
3. **Click Add Event.** Same payload pattern as Demo 2 (Title = "PnPjs Demo Event", future date, etc.).
4. **Watch Network tab.** Show the POST.
5. **Show the request body.** It's the same `toSpWritePayload(item)` shape as before — same mapper, both services use it. Point at the import.
   - **Say:** "The payload is the same because the *list* is the same. The mapper isn't an SPFx-native artifact or a PnPjs artifact — it's a SharePoint shape problem. Both services share it. No duplication."
6. **Click Edit on the new item.** Change something. Save.
7. **Show the network request.** Note: PnPjs may use `PATCH` directly here instead of `POST`+`X-HTTP-Method: MERGE`. Either way, point at how clean the call site is — slide 32.
8. **Click Delete on a test item.**
9. **Show the request.** Note: `DELETE` method, no `X-HTTP-Method` gymnastics. Slide 34's code is the matching call.
   - **Say:** "Remember Demo 2's delete? Three of the four SP REST verbs went through POST. Here, delete is just `.delete()` and PnPjs picks the right wire format. Slide 34 also shows the Graph version of the same call — exact same call site, even though the wire formats differ."

**Compare vs Demo 2:**
- **Say:** "Same network behavior. Half the code. Zero header ceremony. And we haven't even turned on the good stuff yet."

---

## Demo 7 — PnPjs + Graph: same Graph operations

**Goal:** show that `@pnp/graph` smooths the *code* over the awkward Graph URL — but the Graph awkwardness itself doesn't go away.

**Slide cue:** Slide 31.

**Time budget:** 3 min.

**Steps:**

1. **Click Pivot to *PnPjs* → *MS Graph (SP)*.**
2. **Open VS Code to `PnPjsGraphSpService.ts`.** Lines 30–43.
3. **Walk the code:**
   - **Line 35 (`InjectHeaders`):** "This is how PnPjs adds the `Prefer` header. Composable behavior, no manual `.header()` chain — same `using()` shape you saw in Demo 5's Anonymous service."
   - **Lines 37–39 (`itemsQuery.query.set`):** "Honesty time. The PnPjs `.expand()` method has a bug — it splits the comma-separated `fields(select=A,B,C)` argument and produces a malformed URL. So we set the raw query string directly. This is in the codebase as a comment somewhere."
4. **Browser, click Load.**
5. **Show the network request.** Same URL shape as Demo 3. Same response shape with the nested `fields`.
6. **Be honest.**
   - **Say:** "Graph for SP-list items is awkward in *both* clients. PnPjs doesn't fix Graph — Graph is what it is. What PnPjs gives you is `InjectHeaders` and a query builder so the *code* is cleaner even when the URL isn't."

---

## Demo 8 — The free upgrades

**Goal:** make logging, caching, and batching tangible.

**Slide cue:** Slide 35 (logging), Slide 36 (caching), Slide 37 (batching).

**Time budget:** 6 min total. Budget the segments: 1.5 min logging, 2 min caching, 2.5 min batching.

This demo requires three small code changes you'll make live OR pre-stage as a separate "Upgraded" Pivot tab. **Recommendation:** pre-stage the changes in a feature branch and check out the branch before this demo. Live editing fluent chains under time pressure is asking for trouble.

### 8a — Logging (1.5 min)

Already wired. `@pnp/logging` is wrapped in [utilities/logger.ts](src/webparts/dataDemo/utilities/logger.ts), attached in `onInit`, and the level is driven by an **Enhanced logging** property-pane checkbox. No live editing.

**Steps:**
1. **Open browser Console tab** (alongside Network).
2. **Open the property pane.** Toggle **Enhanced logging** off if it's on. Click Refresh on the web part — Console is mostly quiet (Warning level only).
3. **Toggle Enhanced logging on.** Click Refresh again.
4. **Show the Console.** Each call emits `[DataDemo] ...` lines via `console.info`/`console.debug`, with object payloads rendered as collapsible trees (because we emit through `console.*`, not a flat string).
5. **Click the source link on the right side of any log line.** DevTools jumps to the actual call site (e.g. `JokePanel.tsx:42`), not to `logger.ts`.
   - **Say:** "One toggle. `@pnp/logging` underneath, behind a small wrapper. Swap `ConsoleListener` for an App Insights listener and the same `Logger.info(...)` calls become production telemetry. And because the wrapper pre-binds `console.*`, the source link still points where you actually wrote the call — the wrapper isn't a tax, it's an asset."

### 8b — Caching (2 min)

The wiring is already in `ServiceFactory.ts` — the `.using(Caching({ store: 'session' }))` line is there, just commented out. You uncomment one line. No live editing of import statements, no HMR roulette.

**Steps:**
1. **Open VS Code to [ServiceFactory.ts:46-50](src/webparts/dataDemo/services/ServiceFactory.ts#L46-L50).** Show the current state — the `.using(Caching(...))` line is commented:
   ```ts
   const sp = spfi(site.url)
     .using(spSPFx(this.context as any))
     //.using(Caching({ store: 'session' })); // comment out to disable caching
   ```
2. **Uncomment the line.** Save. Let HMR refresh.
3. **Clear the Network tab.**
4. **Click Refresh on the web part.** Show one request go out.
5. **Click Refresh again immediately.** Show **zero new requests** — data is back instantly.
   - **Say (slow it down):** "Watch the Network tab. First click — request goes out. Second click — *nothing*. The data renders instantly because it came from sessionStorage."
6. **Hard-refresh the browser (Ctrl+F5).** Click Refresh on the web part once more.
   - **Say:** "Notice — still no network request. That's because we used `store: 'session'`. The default cache is in-memory and dies with the page. Session storage survives a refresh, local storage survives a tab close. Pick the scope that matches your data's freshness budget."
7. **Show [PnPjsSpService.ts:23-30](src/webparts/dataDemo/services/PnPjsSpService.ts#L23-L30).** Point at the commented `.using(CacheNever())` line on the items query.
   - **Say:** "When global caching is on but one specific query has to be live — pricing, inventory, anything where stale equals wrong — `CacheNever()` opts that single query out. Composable, per-call, no global state to wrangle."
8. **(Optional reset)** Re-comment the `.using(Caching(...))` line before moving on, so the next segment isn't reading stale data.

### 8c — Batching (2.5 min)

We pivoted away from "Bulk Add 5 writes" — the Speaking Events list isn't big enough for a write-heavy story to land. Instead we batch **paginated reads**: 100 calls of `top(5)/skip(N)` packaged into one `$batch` envelope. Same `sp.batched()` call, same payoff.

The wiring is in `PnPjsSpService.getItems`, gated by a comment block. You toggle which version runs by flipping which block is commented out. Pre-flight before the demo: make sure the *single-call* version is active so Demos 4, 6, and 7 looked normal.

**Steps:**
1. **Pivot stays at *PnPjs* → *SharePoint*.**
2. **Open VS Code to [PnPjsSpService.ts:19-59](src/webparts/dataDemo/services/PnPjsSpService.ts#L19-L59).** Show the audience both blocks side-by-side: the single-call return (active) above, the batched block (commented out) below.
   - **Say:** "These are the same query. The top one runs as one HTTP GET. The bottom one fires 100 paginated GETs of 5 items each, packed into one $batch POST. Watch the network tab as I switch them."
3. **Comment out the single-call return** (lines 22-31) and **uncomment the batched block** (lines 33-56). Save. Let HMR refresh.
4. **Clear the Network tab.**
5. **Click Refresh on the web part.**
6. **Show the network tab.** Exactly **one** request: `POST /_api/$batch`. No 100 separate gets.
   - **Say:** "100 page reads. One HTTP request. SharePoint counts requests, not items, when it throttles you — this *is* the throttle escape hatch."
7. **Click the `$batch` request → Payload tab.** Show the multipart body. Each part is a `GET ...?$top=5&$skip=N` line.
   - **Say:** "PnPjs built this multipart body for you. Doing it by hand is the reason most SPFx code never uses `$batch` even though SharePoint's supported it forever."
8. **Click the Response tab.** Show the matching multipart response — one inner result per inner request.
   - **Say:** "Per-page results come back in the same envelope. The pages we fired off as a `Promise.all` resolve when `execute()` returns."
9. **Compare against SPFx mentally.** Switch to [SpfxSpService.ts:19](src/webparts/dataDemo/services/SpfxSpService.ts#L19) — the commented `&$top=5` line.
   - **Say:** "On the SPFx-native side there's no batching primitive at all. To page through 500 items in chunks of 5 you'd fire 100 separate XHRs. `Promise.all` parallelizes them client-side, but SharePoint still sees 100 requests. Batching turns that into one."
10. **(Reset)** Re-comment the batched block and uncomment the single-call return before continuing — otherwise Q&A demos will look weird.

**Honesty beat at the end:**
- **Say:** "One caveat. `$batch` is not a database transaction. If page 3 fails, pages 1 and 2 still came back. PnPjs surfaces per-operation results so you can detect partial failures. Plan your retry logic accordingly."
- **Say (if asked about writes):** "Same `sp.batched()` call. Swap `top/skip` for `add/update/delete`. Five POSTs in one envelope works exactly the same way — we just don't have a list big enough for the read story to be the dramatic one in this room."

---

## Recovery: what to do if a demo dies

| Failure | Recovery |
|---|---|
| Tenant unreachable | Switch to backup screen recording. Apologize once, move on. |
| List doesn't exist / wrong list | Use the property pane (or refresh the page) to re-select. Have the list ID memorized. |
| Joke API down (Demo 1) | This is the OPENER — don't fumble. "The API's down, which is why we don't write production code against random public services without a fallback." Skip to Demo 2 immediately. Mention the Anonymous beat verbally on the way to slide 15. |
| Joke API down (Demo 5) | Cut it. Point at slide 29's code, talk through the `Queryable` composition for 30 seconds, move to Demo 6. |
| Graph permission denied (`/me/messages`) | Stick to `/me` and `/sites/root` in Demo 3. |
| HMR breaks during live code edits in Demo 8 | Don't fight it — fall back to "this is what the code looks like" in VS Code, point at the slide. |
| Caching demo doesn't show instant second click | The `.using(Caching({ store: 'session' }))` line in `ServiceFactory.ts` is on the SPFI, not the per-call query. Confirm you uncommented *that* line, not something else. Also: stale sessionStorage from a previous run will mask the "first click → real request" beat — clear site data and retry. |
| Batching demo still shows N requests | You uncommented the batched block but forgot to comment out the single-call `return` above it — both ran. Re-check that exactly one of the two paths is active in `PnPjsSpService.getItems`. |
| Audience asks an aggressive "PnPjs is bloat" question mid-demo | "Let's hold that for the wrap slide — slide 39 has the honest answer." Don't get pulled off the demo arc. |

---

## Timing reminders

If running long at the pivot (Slide 25), cut from the back:
- **First cut:** Demo 7 (Graph + PnPjs read). The audience already trusts PnPjs by then; the Graph code is on slide 31 and the awkwardness story is the same one you told in Demo 3.
- **Second cut:** Demo 5 (PnPjs anonymous). Slide 29 carries the framing on its own. (You already opened with SPFx Anonymous in Demo 1, so the audience has a concrete reference.)
- **Third cut:** Demo 3 bonus add/delete (slides 23–24). Stick to read + Graph Explorer.
- **Last cut:** Q&A buffer down from 6 to 3 min.

**Never cut:**
- Demo 1 (SPFx Anonymous). It's the opener — without it, Demo 2 lands cold.
- Demo 4 (URL reveal). It's the keystone.
- Demo 8 caching segment. It's the most memorable visual in the talk.
- Slide 25 (the pivot). It's the argument.
