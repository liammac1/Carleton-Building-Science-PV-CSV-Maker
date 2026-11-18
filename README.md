# Rooftop PV sweep (EnergyPlus)

Runs a set of solar PV setups on the house model through EnergyPlus 26.1 and compares each one
with a no-PV baseline in an Excel workbook (energy use, change vs baseline, surplus for export,
costs, GHG, heating/cooling impact and peaks, plus a Notes & Citations sheet).

```
Baseline_PV_corrected.idf           the house model (PV is stripped out to make the baseline)
<Yellowknife>.epw                   the weather file (you add this)
example_scenarios.csv               template for your own scenario list (optional)
pv_sweep.py                         the script
.github/workflows/run-pv-sweep.yml  runs everything on GitHub
```

## Option 1: run on GitHub (nothing to install)

One-time setup:
1. Create a new repository on github.com (private is fine) and upload all of these files,
   keeping the folders. With the web uploader, drag the whole folder contents in; make sure
   `.github/workflows/run-pv-sweep.yml` comes along (it is a hidden folder on Mac/Linux: press
   Cmd+Shift+. in Finder to show it, or create the file on GitHub with "Add file > Create new file"
   and paste the contents).
2. Upload the Yellowknife `.epw` weather file into `weather/` (see `weather/README.md`).

Each run:
1. Open the repository's **Actions** tab and choose **Run PV sweep** on the left.
2. Click **Run workflow**. Keep the defaults, or point it at a different model / scenario CSV.
3. When the run finishes (about 5 minutes; the first run takes slightly longer to install EnergyPlus), open it and
   download **pv-results** from the "Artifacts" section at the bottom. It is a zip with the spreadsheet.

To change the scenarios, edit `scenarios/example_scenarios.csv` on GitHub (or upload your own CSV)
and type its path in the "scenario CSV" box when you start the run. To change the model, upload a new
IDF into `model/` and type its path.

## Option 2: run on your own computer

1. Install EnergyPlus 26.1 from https://energyplus.net/downloads
2. Install Python 3.8+ and then `pip install openpyxl`
3. From this folder:

```
python pv_sweep.py --idf Baseline_PV_corrected.idf --epw <yellowknife>.epw --out pv_results.xlsx
```

The script looks for EnergyPlus in the usual install folders; if it can't find it, add
`--energyplus "C:\EnergyPlusV26-1-0\energyplus.exe"` (Windows) or the path on your machine.
Use `--list` to preview the scenarios without simulating, and `--scenarios my.csv` to run your own list.

## Changing rates or emission factors

No re-run needed: open the spreadsheet's **Assumptions** sheet and edit the blue cells. All cost,
GHG and baseline-comparison columns are formulas. To change the defaults for future runs, edit the
`ASSUMPTIONS` and `CITATIONS` blocks near the top of `pv_sweep.py`.
