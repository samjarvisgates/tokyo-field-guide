# Tokyo Field Guide

197 Tokyo cafés to work from, indexed by three dials — **work**, **vibes**, **food & drink** — that re-order the list rather than filter it. Five spot inks, one per section of the city, on bone stock.

Live: https://samjarvisgates.github.io/tokyo-field-guide/

## What's in it

- `index.html` — the whole app (vanilla JS, MapLibre GL from cdnjs with OpenFreeMap's Liberty style, Google Fonts). Desktop ≥1024px: index + map + unfolding plate. Below that: phone screens (index, map, entry, setting sheet).
- `data.json` — generated from `tokyo_workspots_1.xlsx` (`priority` + `phase2_signals` tabs) by `build/build.py`.
- `photos/<id>.jpg` — each place's Google Maps main photo, 420px; `photos/<id>-N.jpg` — up to five more per place (from `build/images.json`, a Google Maps scrape with images on). Both fetched by `build/fetch.py`.
- `build/apify.json` — coordinates and addresses from the phase-2 Google Maps scrape, keyed by place ID. `build/geocode_extra.json` — four hand-placed pins the scrape didn't cover.

## Order formula

```
crowd = Google and Tabelog ratings, each shrunk toward the city mean by review count
        (150 prior reviews for Google, 60 for Tabelog), z-scored and averaged
mix   = (W·work' + V·vibes + F·food) / (W + V + F)      dials 0–5; work' = sheet work-fit refined by
                                                        outlets / Wi-Fi / laptop signals and time limits
base  = (0.70·mix + 0.30·crowd) × 2                     0–10 scale
score = 6.9 + (base − 6.9) × (0.55 + 0.45·confidence)   confidence = reviews / (reviews + 120)
```

Few reviews hold a place near the middle at every setting; rows under 100 Google reviews carry a THIN EVIDENCE mark and the SOLID EVIDENCE chip hides them. ▲/▼ deltas compare against the 3·3·3 edition. The edition, the CAFÉS ONLY / SOLID EVIDENCE chips and an open plate live in the URL hash (`#w5v3f3c/038`).

"Open now" is computed live in Japan time from the sheet's hours; places whose hours could not be parsed say HOURS UNLISTED and are never shown as open.

## Rebuild

```bash
python3 build/build.py ../tokyo-working/tokyo_workspots_1.xlsx   # -> data.json
python3 build/fetch.py                                            # -> photos/  (then run build.py again)
python3 -m http.server 8765                                       # local preview
```

Design: Claude Design project "01b Field Guide — Five Inks" and "01b Motion Prototype".
