# Admission Offer Sender

Reads a candidate list, fills the college offer letter for each one, and emails it as a PDF.

## Run from source

```bash
uv run app.py          # opens http://127.0.0.1:8765
uv run test_app.py     # self-check: parse, fill, convert, build the email
```

## Build the app

```bash
./build.sh             # macOS → dist/Admission Offer Sender.app
```

PyInstaller cannot cross-compile, so **Windows builds happen on GitHub Actions**
(`.github/workflows/build-windows.yml`): run it from the Actions tab, or push a `v*` tag to get a
`.exe` attached to a release. The workflow smoke-tests the build before uploading it.
Double-clicking the app starts the server and opens the browser. It keeps running until you press
**Quit** in the top-right of the page — there is no terminal window, and Cmd+Q will not reach it.

## Requirements on the machine that runs it

- **LibreOffice** — used to convert each filled letter to PDF. Without it the app refuses to send
  and says so. Nothing else needs installing; the `.app` carries its own Python.
- A **Gmail App Password** for the sending account (Google Account → Security → 2-Step
  Verification → App passwords). A normal password will not work.

## Input file

`.xlsx` or `.csv` with a header row containing an `email` column. The letter is filled from
`Candidate Name`, `Father Name`, `Discipline`, `Matric`, `Inter`, `ECAT`; everything else on the
letter (campus, category, reporting date/time, venue) is typed into the form and applies to the
whole batch. `sample-candidates.csv` is a six-row sample.

## Dry run

Builds every letter and email without sending, and writes each complete message as a `.eml` file
you can open in Mail — attachment included. Packaged app: `~/Documents/Admission Offer Sender/preview`.
From source: `preview/` in the project folder.

## Notes

- `offer-letter.docx` is the template. Values are written into its table cells by label, so the
  wording can change without touching the code — the labels cannot.
- Roughly 0.8s per candidate goes to the PDF conversion; 100 letters take a couple of minutes.
- The sending password is used for the batch and never written to disk or logged.
