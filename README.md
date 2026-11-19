# Treaty Performance Tracker

Excel front end over an MS Access database for capturing treaty statistics by
underwriting year and testing shares. It follows the same layout as the Universe platform
and the CRESTA Aggregate Tracker: screens → services → repositories → database,
numbered migrations, optimistic locking, an audit log, and Universe's reference lists.

## What's in the kit

| Path | What it is |
|---|---|
| `Performance_Tracker.xlsx` | Clean workbook with Universe cedants (with countries), treaty types, classes, countries and currencies; no treaties |
| `Performance_Tracker_DEMO.xlsx` | Same workbook with 5 fictional "Demo -" cedants, 9 treaties and 83 years of stats |
| `vba/*.bas` | The eight VBA modules |
| `Build-Workbook.cmd` / `.ps1` | One-click build: imports the modules and saves `Performance_Tracker.xlsm` with buttons |
| `tools/build_perf_workbook.py` | Generates both `.xlsx` files (`python tools/build_perf_workbook.py .`, needs `openpyxl`) |

## Setup (Windows)

1. Extract the zip to a normal folder (not OneDrive). Unblock the zip first: right-click → Properties → Unblock.
2. Double-click `Build-Workbook.cmd`, or import by hand:
   - Open `Performance_Tracker.xlsx` → **Developer → Visual Basic → File → Import File**, once for each of the 8 `.bas` files.
   - Save as `.xlsm`.
   - **Developer → Macros → InstallButtons → Run**.
3. Open the `.xlsm`, enable macros, click **1. Setup / create database**. This creates `TreatyPerformance.accdb`
   next to the workbook, unless a path is set in Start!C7.

## Workflow

1. **Cedants**: every cedant needs a home country. Use **Add cedant** for new ones.
2. **Treaties**: enter one row per treaty with:
   - cedant, treaty type, class, currency and reference (blank = MAIN; LAYER 1, LAYER 2 … for XL layers)
   - **stats start and end dates**
   - our share, plus optional min / max shares for the optimiser

   Then **Push treaties**.
3. **Generate stat rows**: adds one row per UW year in each treaty's period, e.g.
   01/01/2016–31/12/2025 = 10 rows (2016–2025) and 01/07/2017–30/06/2025 = 8 rows (2017–2024).
4. **Stats**: enter 100% figures in original currency:
   - Premium and Incurred claims
   - **Commission** for proportional treaties, **Brokerage** for non-proportional treaties

   The treaty type's category decides which one applies. Then **Push stats**. Generated rows with no figures stay on the sheet and aren't saved.
5. **Performance**: choose the options below. Up to 20 UW years are shown.
   - Group by: cedant, country, class, treaty type, category or treaty
   - Measure: premium, incurred claims, commission + brokerage, result, loss ratio or combined ratio
   - Basis: 100%, current share, scenario share or optimised share
   - Category filter: all, proportional or non-proportional
6. **Shares**: test shares in two ways:
   - Type scenario shares in column J (blank = current share).
   - Or set the UW years, a risk weight and an optional premium range, then **Run optimiser**.
7. **Scenarios**: compare current, scenario and optimised shares by UW year: premium, claims, result,
   loss ratio, combined ratio, average and standard deviation of annual result, worst year and years with a loss.

## Definitions

- Result = premium − incurred claims − commission − brokerage
- Loss ratio = incurred claims / premium; combined ratio = (claims + commission + brokerage) / premium
- All USD figures use the latest rate per currency on *Rates*. Only USD and the fixed pegs are pre-filled.

## Optimiser

The optimiser maximises **average annual result − λ × standard deviation of annual result**, at our share.
It uses the UW years chosen on *Shares* and has these limits:

- Each treaty stays between its min and max share. Blank min = 0%; blank max = 2 × our share, capped at 100%.
- Optionally, average annual premium at our share stays within a range.

λ = 0 maximises profit; a higher λ prefers steadier years. The method is coordinate search, plus
bisection on a price for premium when the premium range binds. On test data it came within 0.1% of an
exact solver. Suggested shares appear in column M. **Use optimised as scenario** copies them into column J
so you can adjust them.

## Data model (Access)

| Table / query | Holds |
|---|---|
| `cedant` | Cedant with a required `home_country_id` |
| `treaty` | Cedant × treaty type × class × currency × reference (unique); stats start/end; our / min / max share |
| `treaty_stat` | One row per treaty per UW year (unique): premium, incurred claims, commission, brokerage at 100% |
| `ref_*` | Countries, currencies, FX history (`rate_to_usd`), treaty types with category, classes |
| `vw_perf_by_year`, `vw_perf_by_cedant_year`, `vw_perf_by_country_year`, `vw_perf_by_class_year`, `vw_perf_by_treaty_type_year`, `vw_perf_by_treaty_year` | USD premium, claims, commission, brokerage, result, loss and combined ratio, and premium / result at our current share |
| `audit_log`, `schema_migrations` | Who changed what; applied migrations |

To change the schema, add a numbered migration in `PerfMigrate.bas` and bump `MIG_COUNT`.
Never edit a migration that has already been applied.
