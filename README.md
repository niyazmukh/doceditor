# doceditor

Static, in-browser document templating:

- Import a `.docx` or `.txt`
- Highlight text and turn it into editable fields
- Fill fields (text/number/date) + constants + formulas
- Export the filled document as a PDF

## Run locally

```powershell
python -m http.server 5173
```

Then open `http://localhost:5173`.

## Offline / self-contained

- Third-party libraries are vendored under `vendor/`
- After the first load, it works offline via `service-worker.js`

## Backup / restore

Use **Export template** / **Import template** in the UI to move templates between browsers/devices without a backend.

## Templates + presets

- **Templates**: document + fields + settings
- **Presets**: named sets of values/constants per template

