# Docs Filler

**Repo created for P&C**

Fill a Word template with data from a spreadsheet — automatically producing one
finished Word document per row (e.g. one offer letter per employee).

Docs Filler is a **static web page**: you open it in a browser, choose a template
and a spreadsheet, and download a single ZIP of the finished documents. **All of
the work happens in your browser** — there is no server, nothing to install, and
your files are never uploaded anywhere.

You prepare two files:

1. A **Word template** (`.docx`) with placeholders like `{{Employee Name}}`.
2. A **spreadsheet** (`.xlsx`) where each column header matches a placeholder and
   each row is one person.

---

## The placeholder rule (the only thing you need to learn)

Anywhere you want a value to appear in the Word template, type the column name
inside **double curly braces**:

```
Dear {{Employee First Name}},

Your position will be {{Job Title}} and your remuneration will be
INR {{New CTC}} per annum.
```

Your spreadsheet then has a matching header row:

| Employee First Name | Job Title         | New CTC  |
|---------------------|-------------------|----------|
| Name1               | Head of Delivery  | 1234567  |
| Name2               | Quality Engineer  | 123456   |

Rules:

- The text inside `{{ }}` must match a **column header**.
- Matching **ignores capitalisation and extra spaces**, so `{{new ctc}}`,
  `{{New CTC}}` and a header of `New CTC` all match.
- A placeholder can be used **as many times as you like** in the template.
- Any column your template doesn't use is simply ignored — no need to trim the
  spreadsheet.
- **No numbers, no special codes.** Just the column name in `{{ }}`.

This works for **any** template and **any** spreadsheet, as long as you follow
this rule.

---

## How to use it

1. Open the page in a browser (your published GitHub Pages URL).
2. Choose your **Word template** (`.docx`) and your **spreadsheet** (`.xlsx`).
3. Click **Create documents**. A ZIP downloads with one filled document per row,
   named after your template (e.g. `Offer_Template.zip`). Each document inside is
   named after the first spreadsheet column (e.g. `Offer_Template - Jane Doe.docx`).

### Advanced options

Optional settings on the page let you override the defaults:

- **Name output files by column** — name each document after a specific column's
  value (defaults to the first column).
- **Worksheet name** — pick a specific tab when the workbook has several
  (defaults to the first sheet).

---

## Good to know

- **Numbers are tidied up automatically.** `3210000` becomes `3,210,000`, and
  amounts with decimals are rounded to 2 places (`864236.538…` → `864,236.54`).
- **Formatting is kept.** Bold, colour, font, tables, headers and footers — the
  layout of your template is preserved; only the placeholders change.
- **Safety check.** If a `{{Placeholder}}` in the template has no matching
  column, it is left untouched and listed as a warning after processing, so
  nothing goes out half-finished by mistake. **Always review the generated
  documents before sending or filing them.**
- **Privacy by design.** Everything runs locally in your browser. Your template
  and spreadsheet are never uploaded to a server, written to disk, or stored
  anywhere.

---

## Deploying it (GitHub Pages, free)

The repo is ready to publish to GitHub Pages — it's just static files.

**One-time setup:**

1. Push the repo to GitHub (`main` branch).
2. In the repo, go to **Settings → Pages → Build and deployment** and set
   **Source: GitHub Actions**.

**Deploy:** push to `main` (or run the workflow manually from the **Actions**
tab). The included [`.github/workflows/pages.yml`](.github/workflows/pages.yml)
publishes the `site/` folder. When it finishes, the run's **deploy** step shows
the live URL, e.g. `https://<you>.github.io/<repo>/`.

**Custom domain (optional):** **Settings → Pages → Custom domain** — add your
domain, create the DNS record GitHub shows you, and enable **Enforce HTTPS**.

> **Access is public.** GitHub Pages serves to anyone with the URL. That's fine
> here because it's a pure client-side tool that stores nothing — files never
> leave the user's browser.

---

## Running it locally

The site has no build step — serve the `site/` folder with any static file
server, for example:

```
python -m http.server -d site 8080
```

Then open <http://127.0.0.1:8080>. (Opening `index.html` directly via `file://`
won't work — the vendored scripts need to load over http.)

---

## Files in this project

| File / folder                    | What it is                                              |
|----------------------------------|--------------------------------------------------------|
| `site/index.html`                | The page (UI + styling).                               |
| `site/app.js`                    | All the logic — reads the spreadsheet, fills the template, builds the ZIP, in the browser. |
| `site/vendor/jszip.min.js`       | Reads/writes the `.docx` and `.zip` (vendored, no CDN).|
| `site/vendor/xlsx.full.min.js`   | Reads the `.xlsx` (SheetJS, vendored, no CDN).         |
| `site/2i-Logo.png`               | Logo shown in the header.                              |
| `site/.nojekyll`                 | Tells GitHub Pages not to run Jekyll over the assets.  |
| `.github/workflows/pages.yml`    | GitHub Actions workflow that publishes `site/`.        |
