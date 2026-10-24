# Word–Excel Mail Merge (static site)

Plain HTML, CSS and JavaScript — no build step, no server. Everything runs in the browser, so uploaded files never leave the user's computer. Ready for **GitHub Pages**.

## Files
```
index.html        page
css/style.css     styles
js/app.js         user interface
js/mailmerge.js   merge engine (placeholders, formatting, Excel, records)
js/vendor/        third-party libraries, bundled so the site works offline:
                  pizzip, SheetJS (xlsx), mammoth (browser), DOMPurify
tests/            optional Node tests (not needed to run the site; safe to delete)
.nojekyll         tells GitHub Pages to serve files as-is
```

## Publish on GitHub Pages
1. Create a GitHub repository and upload **the contents of this folder** (so `index.html` is in the repository root).
2. Repository **Settings → Pages**.
3. Under *Build and deployment*, choose **Deploy from a branch**, branch `main`, folder `/ (root)`, then **Save**.
4. After a minute the site is live at `https://<your-user>.github.io/<repository>/`.

## Run locally
Open `index.html` directly, or serve the folder (any static server), e.g. `python -m http.server 8000` → http://localhost:8000

## Using it
1. Upload the Word template (.docx) with placeholders such as `<<name>>`, `<<designation>>`.
2. Upload the Excel workbook (.xlsx/.xls/.xlsm): first sheet, headers in row 1. Headers match placeholders case-insensitively, ignoring spaces, hyphens, underscores, periods and slashes.
3. Optionally type values in the placeholder boxes. Empty = no override; **Set blank** = intentionally empty value.
4. Pick a mode: **Replace** (website value replaces the Excel value in every record) or **Append** (all Excel records unchanged, plus one extra record from the website values).
5. Check the live preview (choose any record), then **Generate and Download Word Document** → `merged_document.docx`.

## Behaviour and limitations
- Inserted values take the formatting of the field-name text inside the brackets; surrounding text and spaces are untouched; the uploaded files are never modified.
- Records are separated by page breaks. Header/footer placeholders are filled from the first record only.
- A placeholder must be inside one paragraph; text boxes are not supported. Legacy `.doc` is not supported.
- The preview is an HTML approximation (no exact fonts, pagination or floating objects). The downloaded .docx is authoritative.
- Limits: 10 MB per file, 2000 records.
