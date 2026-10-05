# bli2prism

Turn Octet / BLI binding exports into a **GraphPad Prism project**, an **Excel workbook** and an **HTML QC report**.

You give it a run folder (raw sensorgram exports, the results table from the Octet analysis software, and an Excel setup
workbook that describes the plate). It gives you back:

- a Prism project (`.prism`) with one data sheet per ligand / concentration, an equilibrium table per ligand, and the
  nonlinear fits Prism applies to them;
- an Excel workbook (`Results.xlsx`) with the equilibrium tables, a **Kinetic KDs** tab (per-fit-row KDs with an Include switch, formula-driven
  summaries) and a **KD summary** tab (equilibrium KD from the EC50 fit next to the mean per-row kinetic KD, in nM);
- an HTML QC report: file mapping, time grids, CSV-vs-raw agreement, hook effects, equilibrium and kinetic fit checks.

It is a plain Python package. It runs locally and needs no network access.

> **Graphs are not generated.** Prism's graph files are an undocumented binary format. The tool writes the data sheets and
> the fit analyses; you create the graphs in Prism (or reshape a project that already has them, see *Prism project routes*).

## Try it on the demo experiment

`examples/demo_binder_screen/` is a complete run folder with **simulated data** (two binders, three analytes, 8
concentrations). See `examples/README.md` for what is in it and what the results should look like.

```bash
pip install .                       # Python 3.9+; installs openpyxl, numpy, scipy, matplotlib
bli2prism build    examples/demo_binder_screen
bli2prism workbook examples/demo_binder_screen
bli2prism qc       examples/demo_binder_screen
```

Outputs land in `examples/demo_binder_screen/rebuild/`: `Demo binding experiment.prism`, `Results.xlsx` and `qc_report.html`. Without installing: `python3 -m bli2prism ...` from the repo root.

## Your own experiment

A run folder holds:

| file | what |
|---|---|
| `A1.xls` ... `H<n>.xls` | raw sensorgram exports (Octet text format). The letter is the plate row (concentration step), the number is the sensor column |
| `kineticanalysistableresults*.csv` | the Octet analysis table: one row per ligand / analyte / concentration, with the `Sensor Location`, `X=<t>` equilibrium response and the kinetic fit columns |
| a lab workbook (`.xlsx`) | a `Setup` tab with the plate map (B5:M12) and the dilution series (N5 start concentration, R5 factor), and a `Proteins` tab with names and molecular weights |

1. `bli2prism draft <folder>` writes `bli2prism_setup.xlsx` next to them (it never overwrites). The `bli2prism` tab lists every
   plate-map entry by formula with its plate column and rows, and pre-fills roles, display names and "made of" proteins.
2. Fill the yellow cells: **Kinetic concentrations per ligand** and **Number of analytes** (1-7, equal to the plate columns with
   Role = Analyte) are manual inputs on purpose. Check roles, display names and "made of" (dropdowns list the Proteins tab).
3. `bli2prism build`, `workbook` and `qc` write to `<folder>/rebuild/`. Names in the plate map must match the results CSV exactly.

The setup workbook is flexible about plate layout: a whole column of reference sensors, a single reference sensor stacked
with the ligand sensors, several ligands stacked down one sensor column, and renamed tabs all work.

### Commands

| command | does |
|---|---|
| `draft` | write the setup workbook for a run folder |
| `build` | write the Prism project (`-n` concentrations, `--no-template`, `--blank`, `-t template`, `-o out`) |
| `workbook` | write `Results.xlsx` (`-o` to name it) |
| `qc` | write the HTML QC report |
| `link-proteins` | make the Proteins tab follow the plate map's roles: Role column, desired concentration and total volume by role |
| `upgrade`, `fix-contents` | repair a setup workbook made by an older version or edited by hand |

`build` never overwrites a file: if the output exists it writes `name_2.prism` and says so.

### Proteins tab (bench prep)

After `link-proteins` (and for new drafts): `Setup!R9` is the immobilized ligand concentration, `R10` the ligand volume per
sensor row (200 uL) and `R11` an excess factor (1.2). Each protein's role comes from the plate map; desired concentration
and total volume follow it (ligand / reference: rows x `R10` x `R11` at `R9`; analyte: `Setup!R8` x `R11` at the starting
concentration).

## Prism project routes

- **From scratch** (default when the folder has no `.prism`, or with `--no-template`): any number of ligands and
  concentrations, 1-7 analytes, data sheets and fit analyses, no graphs. Put an **empty** project saved from Prism in the
  folder and name it in the setup sheet's *Prism output file* cell to use your Prism version's own skeleton; the result is
  written to `rebuild/` and your file is never touched.
- **Template**: a populated `.prism` (with graphs) named in the *Prism template* cell. The tool reshapes it (rename, fewer
  ligands or concentrations); it cannot add more than the template holds, and layouts are dropped when the shape shrinks.
  Exactly 3 analytes.

## How the numbers are made

- **Equilibrium KD**: the EC50 of the fixed-slope sigmoid `Y = Bottom + (Top - Bottom) / (1 + 10^(LogEC50 - X))` (the fit the
  Prism analyses apply) on the `X=<t>` response from the results CSV, which is authoritative over the raw trace. Prism refits on
  open; the workbook's value is computed in Python with the same model.
- **Kinetic KDs**: the `KD (M)`, `kon`, `kdis` columns only. A row is used when Full X² < 3, Full R² > 0.9 and KD, kon and kdis
  are determinate. A used row more than 5-fold from the group's median KD is flagged but stays included. The Kinetic KDs tab's
  Include column recalculates every summary.
- **Concentrations** come from each raw file's `Conc1` header, snapped to the dilution step in the setup workbook.

## Caveats

- Names must match across the plate map, proteins and results CSV, character for character.
- The equilibrium KD is not read from Prism's results table: compare one with your Prism analysis the first time.
- Developed against Octet exports of the layout the demo shows, on macOS with Python 3.12. Excel and Prism behaviour was
  verified by hand, not by automated tests.
- The formula evaluator (`xlcalc.py`) supports the functions this tool writes, not all of Excel.

## Tests

```bash
pip install pytest && python3 -m pytest
```

The tests run on the simulated demo experiment only.

## Layout

```
bli2prism/       package (cli, setup sheet, data layer, Prism writer, workbook, QC, fits, formula evaluator)
examples/        the demo experiment and its generator (make_example.py)
tests/
```
