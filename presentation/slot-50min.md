# 50-Minute Slot Variant

Use this when the talk is in a 50-minute slot instead of the canonical 60. The full deck in [slides.md](slides.md) and [demo-scripts.md](demo-scripts.md) remains the source of truth — this file lists what to **skip** and the revised timing, so the master files don't fork.

The deck (`data.pptx`) leads with Anonymous in both passes, so the cut plan below is shaped around that ordering — Anonymous is the opener and can't be dropped.

---

## Runtime target

≈36 slides shown / ≈22 min demos / ≈4 min Q&A. Total ~52 min with a small flex buffer.

---

## What to skip (in order)

### 1. Demo 5 — PnPjs + anonymous live demo. Save ~2 min.

Stay on slide 29. Walk the `Queryable` composition in VS Code, don't pivot to the browser. The audience just saw the SPFx Anonymous network request in Demo 1 — they have the concrete reference.

> **Say:** "Same external API, same network request as Demo 1. Different pipeline composition — `BrowserFetch`, `RejectOnError`, `ResolveOnData`, `JSONParse`. The PnPjs pipeline isn't SharePoint-only; you compose what you need. Slide 29 has the code."

### 2. Demo 7 — PnPjs + Graph live demo. Save ~2 min.

Stay on slide 31. Walk the code, do not switch to the browser. By this point the audience trusts PnPjs from Demo 4's URL reveal — a third "yep, same network request" beat is overkill.

### 3. Demo 3 bonus add/delete (slides 23–24). Save ~1.5 min.

In Demo 3, stick to Graph read + Graph Explorer. Skip the bonus add/delete walkthroughs. Slides 23–24 still get shown; you just don't drive them in the browser.

> **Say (when you arrive at slide 23):** "Bonus slides — Graph add and delete. Real POST, real DELETE, no header gymnastics. The field-shape mapping is the same SharePoint shape problem. Moving on."

### 4. Q&A: 6 min → 4 min. Save 2 min.

Plenty for a 50-min slot; most conferences bake Q&A into a longer Q&A panel anyway.

**Total saved: ~7.5 min.**

---

## Protect at all costs

- **Demo 1** (SPFx Anonymous — it's the opener; without it, Demo 2 lands cold)
- **Slide 25** (the pivot)
- **Demo 4** (URL reveal — the keystone)
- **Demo 8 caching** segment (most memorable visual in the talk)
- **Demo 8 batching** segment (the architectural beat)
- **Slides 35 / 36 / 37** (the free upgrades)
- **Slide 39** (when SPFx-native is the right answer — earns honesty)

---

## Do NOT cut

- **Slides 10–12** (REST anatomy + OData operators). They earn the pivot. Cutting them makes slide 25 land softer because the audience hasn't felt the URL-string pain yet.
- **Demos 2 and 3.** They set up the "this is the work you're doing yourself" feeling that makes slide 25 inevitable rather than evangelical.
- **The bonus SP REST add/delete (slides 18–19) in Demo 2.** They're cheap — the items are already on screen and the writes are 30 seconds each. Cutting them loses the "three of four verbs go through POST" beat that primes the X-HTTP-Method punchline.

---

## Day-of timing breakdown

| Section | 60-min | 50-min | Delta |
|---|---|---|---|
| Title + intro + roadmap (1–9) | ≈4 min | ≈4 min | — |
| REST anatomy (10–12) | ≈3 min | ≈3 min | — |
| Pass 1 SPFx-native (13–24, demos 1–3) | ≈18 min | ≈16 min | Demo 3 bonus add/delete cut |
| Pivot (25) | ≈2 min | ≈2 min | Never cut |
| Pass 2 PnPjs (26–37, demos 4–8) | ≈24 min | ≈20 min | Demo 5 → code walkthrough only, Demo 7 → code walkthrough only |
| Wrap (38–41) | ≈4 min | ≈4 min | — |
| Q&A | ≈6 min | ≈4 min | Trimmed |

---

## If you're still running long mid-talk

Fallback order, on stage, in real time:

1. **First emergency cut:** skip Demo 6 update/delete (slides 32–34). Slide 30's read demo carries the URL-reveal payoff; the writes echo Demo 2's beats. Save ~2 min.
2. **Second emergency cut:** shorten Demo 8 batching — skip the multipart-body Payload-tab click-through, just show the one `$batch` request in Network and move on. Save ~1 min.
3. **Last emergency cut:** Q&A down to 2–3 min.

**Never cut, even on stage:** Demo 1, Demo 4, Demo 8 caching, Slide 25.

---

## Alternative cut path (if the audience cares about external APIs)

Swap cuts #1 and #2:
- **Keep Demo 5** (PnPjs anonymous live).
- **Drop Demo 7 entirely** (not just trim) and lean on slide 31.
- Save the same ~4 min from those two cuts; #3 (Demo 3 bonus) and #4 (Q&A trim) still apply.

Pick this path only if you know the audience leans toward Power Platform / external integration. For a tenant-internal SPFx audience, the primary cut path above is better.

---

## What changed from the old cut plan

If you're coming from an earlier version of this file: the old plan cut the SPFx Anonymous demo (then Demo 3) first. That move is **off the table** in the current deck — Anonymous is now Demo 1 and the opener. Cutting it would leave the SP REST demo cold-starting the talk. The new first cut is Demo 5 (PnPjs Anonymous live) instead.
