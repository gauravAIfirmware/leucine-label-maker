# Leucine Label Maker

A browser page that makes the Leucine probiotic label PDF from the brand
workbook. The person using it uploads the .xlsx, picks the batch sheet, picks
Adults (2 B CFU per strain) or Kids (1 B CFU per strain), checks the patient
list, then generates and downloads the PDF.

Everything runs in the browser. The Excel file is never uploaded anywhere.

Public site (anyone can use it, no sign-in):
https://gauravaifirmware.github.io/leucine-label-maker/

There is also a claude.ai copy at
https://claude.ai/artifact/GXHLwhodZVYZWnhUfDCpqE. claude.ai blocks the
download for some viewers, so share the public site instead.

## Files

| File | What it is |
|---|---|
| `index.html` | The page |
| `generator.js` | Label logic, ported from the backend's `utils/generateLeucineFromExcel.ts` (same template, coordinates and font sizes) |
| `assets.js` | The backend's `pdfTemplates/Leucine_template.pdf` and `assets/checkmark.png`, base64-encoded |

To run it without claude.ai, put the folder on any static web host (or open
`index.html` directly). It needs internet access for SheetJS and pdf-lib from
cdnjs.

If the label template changes, regenerate `assets.js` from the new PDF.

## Nutrition panel width

The template's nutrition panel is 6.5 × 5.7 cm. The generator prints it at
5.7 × 5.7 cm by squeezing it horizontally (text about 12% narrower), and
moves the dimension arrows to match. To change the width, edit
`PANEL_WIDTH_CM` in `generator.js`; set it to 6.5 to print the template as
designed.

## Checks it makes that the backend doesn't

- The strain columns (E to AG) must match the 29 label strains in order, or it
  refuses to generate.
- If the sheet name says "kids" or "adults" and the other option is picked, it
  asks for confirmation.
- A text cell such as "FALSE" or "No" counts as not selected. The backend
  counted any non-empty text as selected.
