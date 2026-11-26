# Test checklist

Use this checklist before publishing a new version or before using the macro
on production drawings.

## Preparation

- [ ] Work on a copy of a real DGN.
- [ ] Note the expected paper sizes of the models/sheets.
- [ ] Make sure the Bentley PDF driver `pdf.pltcfg` is installed.
- [ ] Make sure `MS_PLTCFG_PATH` resolves to the folder containing
      `pdf.pltcfg`.
- [ ] Make sure the driver allows paper size editing.

## Preview

- [ ] Run `PrintCorrectSize.Preview`.
- [ ] Every expected model is listed.
- [ ] No source/design model is listed when only sheet models should print.
- [ ] The detected sizes match the real sheets (within a fraction of a mm).
- [ ] The source in brackets makes sense:
      `SHEET`, `SELECTION`, `BORDER LEVEL`, `CLOSED ELEMENT`, `EXTENTS`.
- [ ] No model is marked `SKIPPED` by mistake.

## Print

- [ ] Run `PrintCorrectSize.Main`.
- [ ] Confirm the model list.
- [ ] The macro finishes without a runtime error dialog.
- [ ] `PrintCorrectSize.log` exists in the output folder.
- [ ] The log reports `OK` for every expected model.

## PDF verification

For each PDF:

- [ ] The file exists and is not zero bytes.
- [ ] The page size in the PDF viewer's document properties matches the
      detected size.
- [ ] The drawing is not clipped.
- [ ] The drawing fills the page (allowing for driver margins).
- [ ] Line weights, fills, text and references look correct.
- [ ] The PDF opens in Acrobat/Reader without errors.

## Edge cases

- [ ] A4 portrait frame.
- [ ] A3 landscape frame.
- [ ] Frame within 5 mm of A4, e.g. 212 × 299 mm → page must be 210 × 297 mm.
- [ ] Frame within 5 mm of A3, e.g. 300 × 423 mm → page must be 297 × 420 mm.
- [ ] Frame 6 mm off A4, e.g. 216 × 300 mm → page must stay 216 × 300 mm.
- [ ] A2 or B-series frame → page must keep the measured size.
- [ ] A custom size, e.g. 350 × 500 mm.
- [ ] A DGN with sheet models and design source models.
- [ ] A design model with a frame on `M_RAMKA`.
- [ ] A design model with a frame on ` _RAMKA` (leading space).
- [ ] A design model with a frame drawn from lines instead of a closed shape.
- [ ] A design model with no frame, using a selection set.
- [ ] A model with a reference attachment.
- [ ] A DGN whose output folder does not exist yet.
- [ ] A DGN in a path containing spaces.
- [ ] A model name containing characters that are invalid in a file name.

## After the test

- [ ] Close the PDFs and remove the test output from the repository working
      folder (PDFs and logs are ignored by `.gitignore`).
- [ ] Do not commit real project DGN files or PDFs.
