# Presentation Handoff Note

This file exists so a fresh Claude Code session (or a future you, on a different machine) can pick up the presentation work without losing the *why* behind the structure.

If you're a new Claude reading this: read [slides.md](slides.md) and [demo-scripts.md](demo-scripts.md) first, then this file. The deliverables explain *what*; this file explains *what we decided and what's still open*. [slot-50min.md](slot-50min.md) describes the cut plan if the talk lands in a 50-minute slot rather than the canonical 60.

---

## What this presentation is

A 60-minute professional talk Don is preparing to give. Target split: roughly 60/40 slides/demos (the current draft is closer to 45/55 — we deliberately leaned demo-heavy because the comparison **is** the lesson).

**Audience:** SPFx developers. Mix of people already shipping and people evaluating PnPjs.

**Topic:** Data access from SPFx. The web part in this repo (`src/webparts/dataDemo`) is a comparison harness — the same CRUD operations against a "Speaking Events" SharePoint list, implemented six ways:

| Transport | Endpoint | File |
|---|---|---|
| SPFx-native | SP REST | `SpfxSpService.ts` |
| SPFx-native | MS Graph (SP list) | `SpfxGraphSpService.ts` |
| SPFx-native | MS Graph (free-form) | `SpfxGraphQueryService.ts` |
| SPFx-native | Anonymous public API | `SpfxAnonymousService.ts` |
| PnPjs | SP REST | `PnPjsSpService.ts` |
| PnPjs | MS Graph (SP list) | `PnPjsGraphSpService.ts` |
| PnPjs | Anonymous public API | `PnPjsAnonymousService.ts` |

The two-tier Pivot UI in `DataDemo.tsx` is the demo surface — top tier picks transport, second tier picks endpoint.

---

## Why the two-pass structure

This was Don's call and it's the right call. We considered three framings:

- **A — Six Ways to Skin a List** (feature tour, pivot-driven)
- **B — From Raw HTTP to PnPjs: A Refactoring Story** (evolution narrative)
- **C — Field Guide to SharePoint Data Access** (decision-focused)

Don picked a fourth: **two passes over the same problem space.** Pass 1 teaches SPFx-native (REST → Graph → anonymous). Pivot slide says "notice what's missing." Pass 2 re-does the same operations with PnPjs and lights up the missing columns (logging, caching, batching).

**Why this beats Option A:** A side-by-side feature tour invites "PnPjs is better, here's why" framing, which sounds like a sales pitch. The two-pass structure earns the conclusion — the audience does the comparing themselves. Pass 1 has to be honest (no strawmanning the SPFx APIs) for Pass 2 to land.

**The keystone slide is #11 (the pivot)** and **the keystone demo is #4 (the URL reveal).** If anything in the talk has to survive a time crunch, it's those two. Slide 11 is a five-column missing-features table. Demo 4 shows that PnPjs and SPFx-native produce the *exact same network request* — the audience needs to leave that demo trusting that PnPjs isn't magic, just a URL builder.

---

## Decisions already made

1. **Graph Explorer demo lives in Pass 1**, not Pass 2 (Don's call). It's the SPFx `MSGraphClientV3` exposed as a sandbox, so it belongs with the SPFx-native pass. It's also useful as emotional fuel — the audience sees raw Graph JSON and starts feeling the "I'm parsing this by hand" pain before the pivot.
2. **Batching demo is a paginated read, not bulk writes.** The list isn't big enough for a write-heavy story. Instead, `PnPjsSpService.getItems` carries two implementations gated by comment markers — a single-call return and a batched block that fires 100 `top(5)/skip(N)` reads through `sp.batched({ maxRequests: 100 })`. Network tab shows 100 page reads collapse into one `POST /_api/$batch`. The "same `sp.batched()` call works for writes" note covers the bulk-write angle verbally. See `demo-scripts.md` Demo 7c.
3. **Honest framing throughout**, not PnPjs evangelism. Slide 16 says explicitly that Graph awkwardness is a Graph problem (not a PnPjs problem). Slide 22 lists when *not* to use PnPjs (bundle size, one-off calls, debugging at the wire).
4. **Demo time > slide time.** We accepted 45/55 instead of 60/40 because the demos *are* the argument. If timing pressure forces cuts, see the "Timing reminders" section at the end of `demo-scripts.md`.

---

## Open items (work still to do)

**Code/demo-harness work is frozen as of 2026-04-30.** Slides and demo scripts may still be edited; service code, logger, and toggle wiring should not change unless a real bug surfaces during rehearsal.

### Don's TODOs before the talk

1. **Verify the bundle-size claim** on slide 22 ("PnPjs adds ~50–80 KB"). That number is approximate, not measured. Run `heft package-solution --production` against a build with and without PnPjs and check actual deltas. If the number is way off, update the slide.
2. **Decide the format for slide 11 (the pivot).** Current draft is a five-column missing-features table. Alternative discussed: a screenshot of SPFx Graph code with five colored highlights, each labeled with what's missing. The screenshot version has more visual punch but takes longer to produce. Either works.
3. **Seed the Speaking Events list** with ~10 items (mix of past/future dates) and at least one disposable test item. Pre-talk checklist at the top of `demo-scripts.md` has the full list.
4. **Backup screen recordings** for each demo. Tenant network is unpredictable.
5. **Rehearse the Demo 7 toggles** until the comment-block swap and the `Caching()` uncomment are muscle memory. The pre-talk checklist asserts the baseline state (caching commented off, single-call read active, sessionStorage cleared); confirm it 30 minutes before going on stage.

### Done (kept here for traceability)

- ~~**Build the "Bulk Add 5" button.**~~ Superseded. The batching demo became a paginated *read* in `PnPjsSpService.getItems`: 100 calls of `top(5)/skip(N)` packed into one `sp.batched({ maxRequests: 100 })` envelope. Single-call vs batched paths both live in the file, gated by comment markers; the demo flips which block is active. Slide 21 and Demo 7c reflect this.
- ~~**Pre-stage logging/caching/batching code** for Demo 7.~~ Caching is a one-line uncomment in `ServiceFactory.ts` (`.using(Caching({ store: 'session' }))`), with a `CacheNever()` per-call opt-out commented in `PnPjsSpService.getItems`. Batching is the comment-block swap described above. Logging is wired through `utilities/logger.ts` and toggled via the **Enhanced logging** property-pane checkbox. No live editing of fluent chains during the talk.

---

## Things I considered and rejected

- **Recording each demo and embedding video instead of going live.** Lower risk, but kills the room energy. Live caching demo (instant second click in the network tab) is the visceral moment of the talk. Don't pre-record it.
- **Adding a section on PnPjs v3 → v4 migration.** Out of scope for this audience and would balloon runtime. Mention in passing during Q&A if asked (the batching API rewrite is the headline).
- **Adding an MGT or Fluent UI angle.** Same — out of scope. This is a *data access* talk, not a UI talk.
- **A demo of `@pnp/queryable` against a non-SP API** (PnPjs version of the joke service). Slide 17 covers it conceptually. A demo would be redundant after Demo 3 unless we want to show the X-PnPjs-RequestId / CORS preflight gotcha — and that's too rabbit-holey for the runtime.

---

## Don's preferences relevant to this work

(Cross-referenced with `~/.claude/CLAUDE.md` — for the new session's benefit.)

- **No sycophancy.** Direct technical opinions, push back when warranted.
- **Address Don as "Don".**
- **Honest pacing.** If Pass 1 isn't honest (e.g., strawmanning SPFx APIs), the talk fails. Don explicitly framed the structure as "teaching the SPFx way first, then going back."
- **Commit hygiene** — atomic commits, no AI references in messages. The presentation work is in `b11d64f` (docs) and `674487f` (refactor), both pushed to `origin/main`.

---

## State of the repo as of this handoff

- Branch: `main`
- Last commit: `674487f refactor: rename SpServiceFactory to ServiceFactory`
- Previous commit: `b11d64f docs: add 60-minute presentation outline and demo scripts`
- Remote: `https://github.com/DonKirkham/DataDemo`
- All work pushed. No uncommitted changes pending the handoff write itself.

The web part code is the *artifact* of the talk. The presentation files in this folder are the *plan* for the talk. Both are needed.

---

## How to resume on a new machine

1. `git clone https://github.com/DonKirkham/DataDemo`
2. `pnpm install`
3. `heft start` — confirm web part loads in workbench against your tenant
4. Re-read [slides.md](slides.md) and [demo-scripts.md](demo-scripts.md)
5. Pick up from "Open items" above
6. If continuing with a fresh Claude session: paste this file in or let it find the file on disk. The slides and demo scripts are self-contained enough that no transcript replay is needed.
