# Polymarket 30D Topic Dashboard

This repo contains a static dashboard that tracks Polymarket's rolling 30-day
event volume share by market topic.

## Dashboard

The dashboard lives in `docs/` so GitHub Pages can publish it directly:

- `docs/index.html` renders the table and line chart.
- `docs/data/current.csv` contains the latest snapshot.
- `docs/data/history.csv` contains the daily time series.
- `docs/data/polymarket_30d_topic_share.xlsx` is the Excel workbook.

## Data Update

The update script uses Polymarket's public Gamma API `/events` endpoint, sorted
by `volume1mo`, and aggregates the top 2000 events by topic tag.

```bash
python3 scripts/update_polymarket_30d.py
```

Preview locally:

```bash
python3 -m http.server 8000 --directory docs
```

Then open `http://localhost:8000/`.

## GitHub Pages

`.github/workflows/polymarket-dashboard.yml` runs every day at 01:15 UTC
(09:15 Beijing time), updates the data files, commits them back to the repo,
and deploys the `docs/` folder with GitHub Pages.
