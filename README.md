# Stock Intelligence

A single-file, dark-themed **inventory intelligence dashboard** for warehouse and supply-planning teams: stock health, availability, demand, slow movers and an automated **daily transfer plan** (3-day demand → safety stock → strategic fill, with FEFO batch picking and pallet/trip capacity).

> **All data in this repository is fictional.** Materials, quantities, prices, batches, warehouse positions and saturation history are generated from a fixed random seed by [`generate_sample_data.py`](generate_sample_data.py). Any resemblance to a real company's inventory is coincidental. This project is not affiliated with or endorsed by any company.

## What's inside

| File | Purpose |
|---|---|
| `index.html` | The dashboard. Open it in a browser — no build step, no server. |
| `update_dashboard.py` | Reads two Excel workbooks, recalculates every KPI, table, chart dataset and the transfer plan, and rewrites them into `index.html`. |
| `generate_sample_data.py` | Builds the fictional workbooks in `sample_data/`. |
| `sample_data/` | Fictional `Stock Control.xlsx` and `Position Saturation.xlsx`. |

## Quick start

```bash
pip install -r requirements.txt

python generate_sample_data.py              # (re)create the fictional workbooks
python update_dashboard.py --as-of 2026-11-17   # rebuild index.html from them
```

Then open `index.html`. Use `--dry-run` to print the calculation summary without writing anything.

## Views

- **Overview** – KPI cards, stock-quality doughnut, capital at risk, Pareto of where the stock sits, saturation trend, demand concentration, and a "what to do now" list.
- **Stock Health** – critical shortages (< 1 week of cover) and overstocked materials (> 12 weeks).
- **Availability** – materials ranked by an availability/health score.
- **Demand** – most consumed materials.
- **Transfer** – daily transfer plan, ready-to-go waves (2 free trips + paid extra), safety-stock table, rupture risks, dedicated/shared position map.
- **Slow Movers** – stock idle for more than 60 days.

Tables are searchable and sortable.

## Using your own data

1. Export your data into two workbooks with the same layout as the samples (keep your real files in `data/` — it is git-ignored).
2. Run:
   ```bash
   python update_dashboard.py --xlsx data/YourStock.xlsx --sat-xlsx data/YourPositions.xlsx
   ```
3. Tune the business rules in section 1 of `update_dashboard.py` (cover thresholds, demand window, safety stock per material, trip capacity, exclusions, storage-location names).

**`Stock Control.xlsx`** sheets

| Sheet | Columns used |
|---|---|
| `In Stock - DC` (positional) | Material · Description · Batch · Special Stock · Storage Location · Storage Type · Storage Bin · GR Date · Stock Category · Available stock · Total Stock · Number of quants · Last movement · Time · Base unit · Expiry date |
| `In Stock - Externals` | Material · Material Description · Storage Location · Batch · Base Unit of Measure · Unrestricted · Transit and Transfer · In Quality Insp. · Restricted-Use Stock · Blocked · Returns · Value Unrestricted · Val. in Trans./Tfr · Value in QualInsp. · Value Restricted · Value BlockedStock |
| `Usage Material` | Material · Material Description · Requirements Date · Difference Quantity · Item Deleted · Phantom Item |

**`Position Saturation.xlsx`** (first sheet): `PROD/END | Total | Occupied | Free | Turnover` rows for shared areas, one row per dedicated material and a `Blocked` row, plus `Shared saturation` / `Dedicated saturation` summary rows and a `Date | Shared saturation | Dedicated saturation` history in columns G–I.

## Notes

- Finished goods (codes with letters or `-`), melted/tank materials and explicitly excluded codes are filtered out everywhere; the script aborts if one leaks into the output.
- A backup of the previous `index.html` is kept as `index.html.bak` (git-ignored; set `BACKUP = False` to disable).
- The page loads [Chart.js](https://www.chartjs.org/) and [three.js](https://threejs.org/) (both MIT) from CDNs and the Anton and Inter fonts (SIL OFL) from Google Fonts, so it needs an internet connection the first time.

## License

[MIT](LICENSE)
