# Demo experiment

`demo_binder_screen/` is a complete run folder with **simulated data**: nothing in it comes from a real experiment. Use it to
try the tool end to end, or as a template for the layout your own run folder needs.

The experiment: two binders (`Binder A`, `Binder B`) immobilized on sensor columns 1 and 2, each tested against three analytes
(`Analyte X`, `Analyte Y` and `Analyte Z`) at 8 concentrations (1 µM down, 3-fold). The traces are 1:1 binding
curves with seeded noise. The true constants are in `make_example.py`:

| ligand | analyte | simulated KD | note |
|---|---|---|---|
| Binder A | Analyte X | 5 nM | |
| Binder A | Analyte Y | 1 nM | |
| Binder A | Analyte Z | none | does not bind: the QC report flags it |
| Binder B | Analyte X | 50 nM | |
| Binder B | Analyte Y | 20 nM | |
| Binder B | Analyte Z | 2 µM | weak |

## Run it

```bash
pip install .                       # or: python3 -m bli2prism ... from the repo root
bli2prism build    examples/demo_binder_screen
bli2prism workbook examples/demo_binder_screen
bli2prism qc       examples/demo_binder_screen
```

Outputs go to `examples/demo_binder_screen/rebuild/`:

- `Demo binding experiment.prism`: the Prism project, named by the setup sheet's *Prism output file* cell (no graphs);
- `Results.xlsx`: the Excel workbook (Equilibrium tabs, Kinetic KDs, KD summary);
- `qc_report.html`: the QC report.

`expected_output/` holds the `Demo binding experiment.prism`, `Results.xlsx` and `qc_report.html` that this data produces, for
comparison. The Prism project has the equilibrium sheets first (`Binder A Equilibrium`, `Binder B Equilibrium`), then the
kinetics sheets, and a nonlinear fit per ligand. It has no graphs: create them in Prism.

With no `.prism` file in the folder, `build` writes a Prism project from scratch under the name in the *Prism output file* cell.
To write into your own empty Prism project instead, save one from Prism with exactly that name, put it in the folder, and run
`build` again.

## What is in the folder

- `A1.xls` ... `H2.xls`: raw sensorgrams in the Octet export layout. The letter is the plate row (concentration step), the number is the sensor column.
- `kineticanalysistableresults.csv`: one fit row per ligand, analyte and concentration, like the Octet analysis export.
- `bli2prism_setup.xlsx`: the setup workbook (`bli2prism`, `Setup`, `Proteins` tabs), filled in. The yellow cells are the inputs. **It is also the input to the generator below**: the ligands, analytes and dilution series are read from it.

## Regenerate

```bash
python3 examples/make_example.py
```

This rewrites the raw files and the results table from the setup workbook (edit the plate map or dilution series there first).

The output is deterministic (fixed random seed), so the files do not change between runs.

## What to expect

- Binder B's kinetic KDs match the simulated values closely. Binder A's equilibrium KD for Analyte X (about 13 nM) is higher than
  its kinetic KD (about 5 nM): at 300 s a tight binder has not reached equilibrium at the low concentrations, which is the kind
  of disagreement the KD summary tab is there to show.
- `Analyte Z` on Binder A never binds, so it has no usable kinetic fits and an unreliable equilibrium fit, both flagged.
