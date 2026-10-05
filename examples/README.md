# Demo experiment

`demo_binder_screen/` is a complete run folder with **simulated data**: nothing in it comes from a real experiment. Use it to
try the tool end to end, or as a template for the layout your own run folder needs.

The experiment: two binders (`Binder A`, `Binder B`) immobilized on sensor columns 1 and 2, each tested against three analytes
(`Target X`, the complex `Target X/Partner Z`, and `Partner Z`) at 8 concentrations (1 µM down, 3-fold). The traces are 1:1 binding
curves with seeded noise. The true constants are in `make_example.py`:

| ligand | analyte | simulated KD | note |
|---|---|---|---|
| Binder A | Target X | 5 nM | |
| Binder A | Target X/Partner Z | 1 nM | |
| Binder A | Partner Z | none | does not bind: the QC report flags it |
| Binder B | Target X | 50 nM | |
| Binder B | Target X/Partner Z | 20 nM | |
| Binder B | Partner Z | 2 µM | weak |

## Run it

```bash
pip install .                       # or: python3 -m bli2prism ... from the repo root
bli2prism build    examples/demo_binder_screen
bli2prism workbook examples/demo_binder_screen
bli2prism qc       examples/demo_binder_screen
```

Outputs go to `examples/demo_binder_screen/rebuild/`. `expected_output/` holds the workbook and QC report that this data
produces, for comparison. (The Prism project is not included: it has no graphs and is written by `build`.)

With no `.prism` file in the folder, `build` writes a Prism project from scratch. To write into your own empty Prism project,
put it in the folder and enter its file name in the "Prism output file" cell of the setup workbook's `bli2prism` tab.

## What is in the folder

- `A1.xls` ... `H2.xls`: raw sensorgrams in the Octet export layout. The letter is the plate row (concentration step), the number is the sensor column.
- `kineticanalysistableresults.csv`: one fit row per ligand, analyte and concentration, like the Octet analysis export.
- `bli2prism_setup.xlsx`: the setup workbook (`bli2prism`, `Setup`, `Proteins` tabs), filled in. The yellow cells are the inputs.

## Regenerate

```bash
python3 examples/make_example.py
```

The output is deterministic (fixed random seed), so the files do not change between runs.

## What to expect

- Binder B's kinetic KDs match the simulated values closely. Binder A's equilibrium KD for Target X (about 13 nM) is higher than
  its kinetic KD (about 5 nM): at 300 s a tight binder has not reached equilibrium at the low concentrations, which is the kind
  of disagreement the KD summary tab is there to show.
- `Partner Z` on Binder A never binds, so it has no usable kinetic fits and an unreliable equilibrium fit, both flagged.
