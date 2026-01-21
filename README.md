# QuoteToPDF (MVP)

Static, in-browser document templating:

- Import a `.docx` or `.txt`
- Highlight text and turn it into editable “fields”
- Fill fields (text/number/date) + define constants + compute formula fields
- Export the filled document as a PDF

## Run locally

Just open `index.html` in a browser, or serve the folder:

```powershell
python -m http.server 5173
```

Then open `http://localhost:5173`.

## Notes / limitations

- `.doc` is not supported in the browser (convert to `.docx`).
- Formatting aims to be “basic but readable” (paragraphs/line breaks); it is not a pixel-perfect Word renderer.
- Everything is stored in `localStorage` for now (no sign-in backend yet).

## Offline / self-contained

The app is fully static and now vendors its third-party libraries under `vendor/`.
If you serve this folder (GitHub Pages, any static host, or `python -m http.server`), it will work offline after the first load via `service-worker.js`.

## Backup / restore

Use **Export template** / **Import template** in the UI to back up or move a template to another browser/device without a backend.

## Templates + presets

- **Templates**: document + fields + constants + settings
- **Presets**: named sets of values/constants for a template (e.g., different customers)
