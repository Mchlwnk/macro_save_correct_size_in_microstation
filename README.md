# PrintCorrectSize for MicroStation V8i

A small VBA macro for MicroStation V8i that prints each model of a DGN to a
**separate PDF with the correct page size**.

It solves the usual batch-print problem: when a project contains A4, A3 and
custom-size sheets, choosing one paper size makes small sheets tiny and clips
large sheets. This macro measures the frame in every model and sets the PDF
page size individually.

The primary workflow is: one DGN, several **design models**, each model has a
frame on a level whose name contains `RAMKA` (for example `M_RAMKA` or
` _RAMKA`).

> Status: source code is ready for testing on real drawings. Test it on a copy
> of a DGN first. The repository intentionally contains source, not a compiled
> `.mvba` binary, so it stays readable and reviewable.

## What it does

- Prints all printable models of the active DGN, one PDF per model.
- Measures the frame on the `RAMKA` level in every model.
- Detects page size from:
  1. a sheet model's sheet definition (if present),
  2. elements on a frame level (`RAMKA`, `M_RAMKA`, ` _RAMKA`),
  3. a selection set (fallback manual override),
  4. the largest closed element,
  5. the extents of all graphical elements.
- Snaps to the exact A4 or A3 page when the measured frame is within 5 mm of
  that size. Every other size keeps its measured width × height, including
  other A/B sizes.
- Sets the PDF page to the detected width × height (including custom sizes
  such as 350 × 500 mm).
- Uses MicroStation's built-in Bentley PDF driver (`pdf.pltcfg`), so no
  external PDF printer is required.
- Writes a `PrintCorrectSize.log` with the detected sizes and results.
- Provides a preview mode that only lists the detected sizes.

## Requirements

- MicroStation V8i (SELECTseries) with VBA.
- Design models with a frame on a level whose name contains `RAMKA`
  (`M_RAMKA`, ` _RAMKA`). Other frame-level names can be configured.
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

   If only one model is printed, the name is `<dgn-name>.pdf`. Spaces and
   characters that are invalid in Windows file names are replaced by `_`.

## How the page size is detected

The first matching rule wins:

| Order | Source | When it is used |
|---|---|---|
| 1 | `SHEET` | The model is a sheet model and has a valid sheet definition. |
| 2 | `BORDER LEVEL` | Elements exist on a level whose name contains one of the hints from `BORDER_LEVEL_HINTS`. The default `RAMKA` matches `M_RAMKA` and ` _RAMKA`. |
| 3 | `SELECTION` | No frame was found, but something is selected in the model. Select the border or the elements that define the page. |
| 4 | `CLOSED ELEMENT` | The largest closed element (usually a rectangle) is used. |
| 5 | `EXTENTS` | Fallback: the extents of all graphical elements. |

For a design model with a frame on `M_RAMKA` or ` _RAMKA`, the frame is
measured and that is the page. If the frame is a closed shape, the macro uses
the largest closed element on the frame level; if the frame is drawn from
lines, it uses the union of the non-text elements on that level. Text and
dimensions on the frame level are ignored so notes outside the frame do not
enlarge the page.

For a sheet model, MicroStation's own sheet size is authoritative. The
selection set is only a fallback for unusual models without a frame: select
the page border before running the macro.

By default `PRINT_ONLY_SHEETS_IF_PRESENT` is `False`, so every model is printed
when there are no sheet models. If a DGN does contain sheet models and you
want only those, set it to `True`.

## Configuration

All settings are constants at the top of `PrintCorrectSize.bas`:

| Constant | Default | Meaning |
|---|---|---|
| `PDF_DRIVER` | `pdf.pltcfg` | Bentley PDF plot driver. |
| `CHECK_PDF_DRIVER` | `True` | Check that the driver exists before printing. |
| `OUTPUT_SUBFOLDER` | `PDF` | Sub-folder for PDFs; `""` writes next to the DGN. |
| `BORDER_LEVEL_HINTS` | `RAMKA` | Comma-separated, case-insensitive level-name fragments. Matches `M_RAMKA` and ` _RAMKA`. |
| `USE_SELECTION_AS_PAGE` | `True` | Use a selection set as a fallback when no frame is found. |
| `PRINT_ONLY_SHEETS_IF_PRESENT` | `False` | If sheet models exist, print only them. |
| `STANDARD_SIZES_TO_SNAP` | `A4,A3` | Sizes that are snapped to exact ISO dimensions. |
| `STANDARD_SIZE_TOLERANCE_MM` | `5` | Maximum difference from a standard size, per side, for snapping. |
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
- A4 and A3 are snapped to their exact ISO dimensions when the measured frame
  is within 5 mm. Other sizes, including A0–A2 and B-series, keep the measured
  width × height. The preview shows both the final page and the measured size.
- The PDF driver may still apply its own small internal margins, so the
  drawing can be a fraction of a millimetre smaller than the page.
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
- Extra PDFs: the DGN contains models you do not want to print. By default all
  models are printed; remove or rename unwanted models, or set
  `PRINT_ONLY_SHEETS_IF_PRESENT` if the file contains sheet models.
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
