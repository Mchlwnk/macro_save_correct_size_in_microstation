# Troubleshooting

## The macro is not in the Macros dialog

- Make sure the `PrintCorrectSize` project is loaded in
  **Utilities > Macro > Project Manager**.
- Make sure the module was imported into that project.
- If you renamed the project, use the new name in the key-in:
  `vba run [YourProjectName]PrintCorrectSize.Main`.
- Check that macros are enabled in MicroStation.

## `vba run ...` says the macro was not found

The syntax is:

```text
vba run [ProjectName]ModuleName.ProcedureName
```

For this project:

```text
vba run [PrintCorrectSize]PrintCorrectSize.Main
```

The first `PrintCorrectSize` is the **project** name, the second is the
**module** name. Both are case-insensitive in practice, but spelling must
match.

## No PDF is created, or the log says `FAIL`

1. Check that the Bentley PDF driver is installed:
   `pdf.pltcfg`.
2. Check the configuration variable `MS_PLTCFG_PATH`. It must point to the
   folder that contains `pdf.pltcfg`.
3. Open the Print dialog manually, select the Bentley PDF driver, and try to
   print one model. If that also fails, the problem is in the driver setup,
   not in the macro.
4. Check that the output folder is writable. The macro writes to a `PDF`
   sub-folder next to the DGN.
5. Check `PrintCorrectSize.log` in the output folder for the failing model.

The timeout is 180 seconds per PDF. A very large or raster-heavy model may
need more time. Increase `PDF_TIMEOUT_SEC` in the source if needed.

## The PDF page has a standard size instead of the custom size

The macro sets the page with the documented key-ins:

```text
PRINT XFORM <width> MM
PRINT YFORM <height> MM
```

This requires a Bentley driver that allows editing the paper size. In the
Print dialog:

1. Open **File > Edit Printer Driver Configuration**.
2. Check the base properties for an option such as **Allow paper size
   editing** and enable it.
3. Save the driver configuration.
4. Run the macro again.

If your organisation uses a custom `.pltcfg` that locks the paper size, copy
the standard `pdf.pltcfg`, enable paper size editing, and set `PDF_DRIVER` in
the macro to the name of your copy.

A very large custom page can also exceed the driver's maximum form size. In
that case the driver may clamp the page or fail to print.

## The page size is wrong

Run `Preview` first. The source is shown in brackets:

- `[SHEET]` — the sheet model definition was used. Check
  **Model Properties > Sheet > Size** in MicroStation.
- `[SELECTION]` — something was selected when the macro ran. Clear the
  selection or select the correct border.
- `[BORDER LEVEL]` — the macro found elements on a level whose name matches
  `BORDER_LEVEL_HINTS` (default: `RAMKA`, so `M_RAMKA` and ` _RAMKA` match).
  If the level has a closed frame shape, the largest closed element is used;
  otherwise the union of non-text elements on that level is used. Text and
  dimensions are ignored, but other geometry on the frame level can still
  enlarge the page.
- `[CLOSED ELEMENT]` — the largest closed element was used. If the largest
  closed element is not the page border, put the border on a recognised level
  or select it manually.
- `[EXTENTS]` — no border/sheet definition was found, so all geometry was
  used. Elements outside the page will make the page larger.

### Manual override

For a design model, select the border (or all elements that define the page)
before running the macro. The selection set is used as the page rectangle.
This is the quickest way to fix a wrong automatic detection.

If the frame is inside a reference attachment rather than in the active model,
the scan does not see it. Place the frame in the master model or select the
page elements manually.

## A4 sheets are still small, or large sheets are still clipped

This means the page size was not applied to the PDF. Check:

1. The `Preview` output — does each model have a different size?
2. The PDF page size in Acrobat: **File > Properties** (or the page size shown
   in the PDF viewer).
3. The driver's paper size editing setting (see above).

If the preview already shows the same size for every model, the detection is
wrong, not the printing. Check the source in brackets and use a selection or
the border level.

### A4/A3 snapping

A measured frame is snapped to exact A4 or A3 only when **both** sides are
within 5 mm of the standard. For example:

- 212 × 299 mm → A4 (210 × 297 mm)
- 300 × 423 mm → A3 (297 × 420 mm)
- 216 × 300 mm → custom 216 × 300 mm (more than 5 mm from A4)
- 420 × 594 mm → custom 420 × 594 mm (A2 is not in the default snap list)

The preview shows both the final page and the measured size. To snap other
sizes, add them to `STANDARD_SIZES_TO_SNAP` (for example `A4,A3,A2,B1`).
Every name must exist in `PaperDimensions` in the source.

## Extra PDFs are created

The macro prints every selected model. By default all models are printed
(`PRINT_ONLY_SHEETS_IF_PRESENT = False`). If a DGN contains sheet models and
you want only those, set `PRINT_ONLY_SHEETS_IF_PRESENT = True`.

If you get PDFs for models you do not want to print, options:

- rename or remove the models you do not want to print,
- add a model filter in the source,
- or keep separate DGN files for source geometry and sheets.

A per-model include/exclude list is not part of the current version.

## One DGN per object

The current version prints the **active DGN**. If each object is a separate
`.dgn` file:

1. Open the file.
2. Run the macro.
3. Repeat for the next file.

A folder-batch mode is not included yet because opening files automatically
can discard unsaved changes. If you need it, add it only with a workflow that
saves and closes files first.

## The drawing is slightly smaller than the page

The PDF page size is exact, but the PDF driver may apply small internal
margins. This is normal for some driver configurations. Check the driver's
margins in the Printer Driver Configuration dialog.

If the drawing is much smaller, `Maximize` is fitting a page whose aspect
ratio does not match the drawing. Check that the detected size is correct.

## The drawing is clipped

The detected page is smaller than the actual drawing. Check the preview and
the source:

- for a sheet model, check the sheet size;
- for a border level, make sure the level contains the whole border and
  nothing is on a hidden/frozen level;
- select the full border manually.

## The DGN or output folder contains spaces

The macro normally converts the output folder to its Windows short path
(`PROGRA~1` style) and replaces spaces in the file name with `_`. This avoids
problems with the MicroStation key-in parser. If the volume does not provide
short paths (some network shares), the macro falls back to quoting the path.
If that still fails, use a DGN/output folder without spaces for the test.

## The macro changes my fence or active model

The macro uses a temporary MicroStation fence and activates each model while
printing. At the end it removes the fence and returns to the model that was
active when the macro started. Any fence that existed before the run is not
restored. Save your file and avoid running the macro while you need a
specific fence.

The macro does not change geometry.

## The DGN is read-only

Printing should still work, but the macro creates the `PDF` folder next to the
DGN. If that folder is not writable, the macro stops. Use a writable copy of
the DGN or change `OUTPUT_SUBFOLDER` to a writable location.

## MicroStation version

The macro targets MicroStation V8i and uses V8i VBA objects:

- `ModelReference.GetSheetDefinition`
- `Fence.DefineFromModelPoints`
- the `PRINT ...` key-ins

MicroStation CONNECT has a different print API, so this exact source is not
intended for CONNECT. The detection and file-handling parts can be adapted,
but the print calls would need to be replaced.
