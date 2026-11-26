# PrintCorrectSize for MicroStation V8i

A small VBA macro for MicroStation V8i that prints each model/sheet of a DGN
to a **separate PDF with the correct page size**.

It solves the usual batch-print problem: when a project contains A4, A3 and
custom-size sheets, choosing one paper size makes small sheets tiny and clips
large sheets. This macro detects the size of every model and sets the PDF page
size individually.

> Status: source code is ready for testing on real drawings. Test it on a copy
> of a DGN first. The repository intentionally contains source, not a compiled
> `.mvba` binary, so it stays readable and reviewable.

## What it does

- Prints all printable models of the active DGN, one PDF per model.
- Detects page size from:
  1. a sheet model's sheet definition,
  2. a selection set (manual override),
  3. a border/plot level,
  4. the largest closed element,
  5. the extents of all graphical elements.
- Sets the PDF page to the exact detected width × height (including custom
  sizes such as 350 × 500 mm).
- Uses MicroStation's built-in Bentley PDF driver (`pdf.pltcfg`), so no
  external PDF printer is required.
- Writes a `PrintCorrectSize.log` with the detected sizes and results.
- Provides a preview mode that only lists the detected sizes.

## Requirements

- MicroStation V8i (SELECTseries) with VBA.
- The standard Bentley PDF driver `pdf.pltcfg` installed and available through
  `MS_PLTCFG_PATH`.
- The driver must allow paper size editing. Bentley PDF drivers normally do.

## Installation

1. Download [`src/PrintCorrectSize.bas`](src/PrintCorrectSize.bas).
2. In MicroStation, open **Utilities > Macro > Project Manager**
   (in some V8i versions: **Tools > VBA > Project Manager**).
3. Create a new VBA project named `PrintCorrectSize` and save it as
   `PrintCorrectSize.mvba` in a folder that MicroStation searches for macros,
   for example:
   `...\Workspace\Standards\Macros\vba`.
4. Open the VBA editor and choose **File > Import File...**.
5. Select `PrintCorrectSize.bas`.
6. Save the VBA project.
7. Load the project in the Project Manager if it is not loaded yet.

Detailed instructions, including how to find the right folder, are in
[`docs/INSTALL.md`](docs/INSTALL.md).

## Usage

1. Open the DGN you want to print.
2. Save the file.
3. Run the preview first:

   ```text
   vba run [PrintCorrectSize]PrintCorrectSize.Preview
   ```

   Check the list of models and detected page sizes.

4. Run the main macro:

   ```text
   vba run [PrintCorrectSize]PrintCorrectSize.Main
   ```

   You can also run both from **Utilities > Macros**.

5. The PDFs are written to a `PDF` sub-folder next to the DGN:

   ```text
   <dgn-name>_<model-name>.pdf
   ```

   If only one model is printed, the name is `<dgn-name>.pdf`.

## How the page size is detected

The first matching rule wins:

| Order | Source | When it is used |
|---|---|---|
| 1 | `SHEET` | The model is a sheet model and has a valid sheet definition. |
| 2 | `SELECTION` | Something is selected in the model. Select the border or the elements that define the page. |
| 3 | `BORDER LEVEL` | Elements exist on a level whose name contains one of the hints from `BORDER_LEVEL_HINTS`. |
| 4 | `CLOSED ELEMENT` | The largest closed element (usually a rectangle) is used. |
| 5 | `EXTENTS` | Fallback: the extents of all graphical elements. |

For a sheet model, MicroStation's own sheet size is authoritative. The
selection set is useful for design models that do not have a border on a
recognised level: select the page border before running the macro.

By default, if a DGN contains sheet models, only sheet models are printed.
Design models in such a file are usually source geometry, not printable
sheets. Change `PRINT_ONLY_SHEETS_IF_PRESENT` in the source if you need every
model.

## Configuration

All settings are constants at the top of `PrintCorrectSize.bas`:

| Constant | Default | Meaning |
|---|---|---|
| `PDF_DRIVER` | `pdf.pltcfg` | Bentley PDF plot driver. |
| `OUTPUT_SUBFOLDER` | `PDF` | Sub-folder for PDFs; `""` writes next to the DGN. |
| `BORDER_LEVEL_HINTS` | `PLOT,BORDER,SHEET,RAMKA,ARKUSZ,FORMAT,OBRAMOWANIE` | Comma-separated, case-insensitive level-name fragments. |
| `USE_SELECTION_AS_PAGE` | `True` | Use a selection set as the page size override. |
| `PRINT_ONLY_SHEETS_IF_PRESENT` | `True` | If sheet models exist, print only them. |
| `MIN_PAGE_MM` | `40` | Reject detected pages smaller than this. |
| `MAX_PAGE_MM` | `10000` | Reject detected pages larger than this. |
| `PDF_TIMEOUT_SEC` | `180` | Wait time for one PDF. |
| `WRITE_LOG` | `True` | Write `PrintCorrectSize.log`. |

## Important notes

- The macro changes the active model and creates/removes a MicroStation fence.
  Save your file before running it.
- It does not modify drawing geometry. It only changes print settings and the
  temporary fence.
- It prints the active DGN. If every object is a separate DGN file, run the
  macro once per file (a safe folder-batch mode may be added later).
- The PDF page size is exact. The PDF driver may still apply its own small
  internal margins, so the drawing can be a fraction of a millimetre smaller
  than the page.
- Custom sizes depend on the driver's paper size editing support. If the
  driver ignores `PRINT XFORM` / `PRINT YFORM`, use a Bentley PDF driver that
  allows editing the paper size.

## Troubleshooting

See [`docs/TROUBLESHOOTING.md`](docs/TROUBLESHOOTING.md).

Quick checks:

- PDF is missing or the log says `FAIL`: verify that `pdf.pltcfg` is installed
  and that `MS_PLTCFG_PATH` points to its folder.
- Page size is wrong: run `Preview`, check the source in brackets, and select
  the page border before running the macro if needed.
- Extra PDFs: the DGN contains models you do not want to print. By default,
  sheet models are preferred; otherwise remove or rename unwanted models.
- Macro not found: the project is not loaded, or the project name is not
  `PrintCorrectSize`.

## Repository layout

```text
src/PrintCorrectSize.bas      Macro source (import this into a VBA project)
docs/INSTALL.md               Detailed installation instructions
docs/TEST_CHECKLIST.md        Checklist for testing on real drawings
docs/TROUBLESHOOTING.md       Common problems and fixes
README.pl.md                  Polish version of this README
LICENSE                       MIT license
```

## License

MIT. See [LICENSE](LICENSE).

MicroStation and Bentley are trademarks of Bentley Systems, Incorporated.
This project is not affiliated with or endorsed by Bentley Systems.
