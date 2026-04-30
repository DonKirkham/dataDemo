# 50-Minute Slot Variant

Use this when the talk is in a 50-minute slot instead of the canonical 60. The full deck in [slides.md](slides.md) and [demo-scripts.md](demo-scripts.md) remains the source of truth — this file lists what to **skip** and the revised timing, so the master files don't fork.

---

## Runtime target

≈19 slides / ≈26 min demos / ≈4 min Q&A. Total ~52 min with a small flex buffer.

---

## What to skip (in order)

### 1. Demo 3 — SPFx + anonymous (Joke API). Save ~3 min.

Stay on slide 10 conceptually. Don't switch to the Anonymous pivot tab.

> **Say:** "If you ever need to call a public API — weather, sports scores, anything not in your tenant — `HttpClient` is the answer. Same package, no auth, you're done. We're not going to demo it live because the comparison we care about today is on SharePoint."

### 2. Slide 17 — PnPjs + anonymous. Save ~1 min.

Skip past it. Slide 10's framing already implied PnPjs handles the same case via `@pnp/queryable`. If the audience asks how, point them to slide 17 in the posted deck.

### 3. Demo 6 — PnPjs + Graph live demo. Save ~2 min.

Stay on slide 16. Walk the code, do not switch to the browser. By this point the audience trusts PnPjs from Demo 4's URL reveal — a third "yep, same network request" beat is overkill.

### 4. Q&A: 6 min → 4 min. Save 2 min.

Plenty for a 50-min slot; most conferences bake Q&A into a longer Q&A panel anyway.

**Total saved: ~8 min.**

---

## Protect at all costs

- **Slide 11** (the pivot)
- **Demo 4** (URL reveal — the keystone)
- **Demo 7 caching** segment (most memorable visual in the talk)
- **Demo 7 batching** segment (the architectural beat)
- **Slides 18 / 19 / 20 / 21** (the free upgrades)
- **Slide 22** (when SPFx-native is the right answer — earns honesty)

---

## Do NOT cut

- **Slides 4–6** (SP REST mechanics). They earn the pivot. Cutting them makes slide 11 land softer because the audience hasn't felt the URL-string pain yet.
- **Demos 1 and 2.** They set up the "this is the work you're doing yourself" feeling that makes slide 11 inevitable rather than evangelical.

---

## Day-of timing breakdown

| Section | 60-min | 50-min | Delta |
|---|---|---|---|
| Title + roadmap (1–3) | ≈3 min | ≈3 min | — |
| Pass 1 SPFx-native (4–10, demos 1–2) | ≈24 min | ≈21 min | Demo 3 cut |
| Pivot (11) | ≈2 min | ≈2 min | Never cut |
| Pass 2 PnPjs (12–21, demos 4–7) | ≈23 min | ≈20 min | Slide 17 skipped, Demo 6 → code walkthrough only |
| Wrap (22–24) | ≈4 min | ≈4 min | — |
| Q&A | ≈6 min | ≈4 min | Trimmed |

---

## If you're still running long mid-talk

Fallback order, on stage, in real time:

1. **First emergency cut:** skip Demo 5 (PnPjs + SP write). Slide 15's code carries the point. Save ~3 min.
2. **Second emergency cut:** shorten Demo 7 batching — skip the multipart-body Payload-tab click-through, just show the one `$batch` request in Network and move on. Save ~1 min.
3. **Last emergency cut:** Q&A down to 2–3 min.

**Never cut, even on stage:** Demo 4, Demo 7 caching, Slide 11.

---

## Alternative cut path (if the audience cares about external APIs)

Swap cuts #1 and #3:
- **Keep Demo 3** (Joke API live).
- **Drop Demo 6 entirely** (not just trim) and lean on slide 16.
- Save the same ~5 min from those two cuts; #2 (slide 17) and #4 (Q&A trim) still apply.

Pick this path only if you know the audience leans toward Power Platform / external integration. For a tenant-internal SPFx audience, the primary cut path above is better.
