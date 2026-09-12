# Tokyo Field Guide

197 Tokyo cafés to work from, indexed by three dials — **work**, **vibes**, **food & drink** — that re-order the list rather than filter it. Five spot inks, one per section of the city, on bone stock.

Live: https://samjarvisgates.github.io/tokyo-field-guide/

## What's in it

- `index.html` — the whole app (vanilla JS, Leaflet from cdnjs, Google Fonts). Desktop ≥1024px: index + map + unfolding plate. Below that: phone screens (index, map, entry, setting sheet).
- `data.json` — generated from `tokyo_workspots_1.xlsx` (`priority` + `phase2_signals` tabs) by `build/build.py`.
- `photos/<id>.jpg` — each place's Google Maps main photo, 420px, fetched by `build/fetch.py`.
- `build/apify.json` — coordinates and addresses from the phase-2 Google Maps scrape, keyed by place ID. `build/geocode_extra.json` — four hand-placed pins the scrape didn't cover.

## Order formula

```
mix   = (W·work + V·vibes + F·food) / (W + V + F)      dials 0–5, sheet scores 1–5
score = (mix × 0.9 + uniqueness × 0.1) × 2             shown on a 0–10 scale
```

Uniqueness never leaves the mix. ▲/▼ deltas compare against the 3·3·3 edition. The edition is in the URL hash (`#w5v3f3/038`) so a setting — and an open plate — can be shared.

"Open now" is computed live in Japan time from the sheet's hours; places whose hours could not be parsed say HOURS UNLISTED and are never shown as open.

## Rebuild

```bash
python3 build/build.py ../tokyo-working/tokyo_workspots_1.xlsx   # -> data.json
python3 build/fetch.py                                            # -> photos/  (then run build.py again)
python3 -m http.server 8765                                       # local preview
```

Design: Claude Design project "01b Field Guide — Five Inks" and "01b Motion Prototype".
