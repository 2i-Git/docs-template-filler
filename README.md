# Docs Filler

Fill a Word template with data from a spreadsheet — automatically producing one
finished Word document per row (e.g. one offer letter per employee).

You prepare two files:

1. A **Word template** (`.docx`) with placeholders like `{{Employee Name}}`.
2. A **spreadsheet** (`.xlsx`) where each column header matches a placeholder and
   each row is one person.

The tool creates one filled-in document per row in an `output` folder.

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

## How to run it

### First time only — install the tools

Open **Terminal** (Mac) or **Command Prompt** (Windows), go to this folder, and
run:

```
pip install -r requirements.txt
```

### Every time — generate the documents

Put your template and spreadsheet in the `input` folder, then run (both
`--template` and `--data` are required):

```
python fill_docs.py --template "input/Offer_Template.docx" --data "input/content.xlsx"
```

Finished documents are written to the `output/` folder by default (change it
with `--outdir`).

You'll see one line per document created, for example:

```
  ✓ Offer_Template - Test Name1.docx
  ✓ Offer_Template - Test Name2.docx

Done. 2 document(s) written to 'output/'.
```

### Using your own file names

```
python fill_docs.py --template "input/MyLetter.docx" --data "input/staff.xlsx" --outdir "letters"
```

Name each output file after a particular column (default is the first column):

```
python fill_docs.py --template "input/MyLetter.docx" --data "input/staff.xlsx" --name-column "Employee Name"
```

See every option with:

```
python fill_docs.py --help
```

---

## Good to know

- **Numbers are tidied up automatically.** `3210000` becomes `3,210,000`, and
  amounts with decimals are rounded to 2 places (`864236.538…` → `864,236.54`).
- **Formatting is kept.** Bold, colour, font, tables — the layout of your
  template is preserved; only the placeholders change.
- **Multiple sheets?** Pick one with `--sheet "SheetName"` (the first sheet is
  used by default).
- **Safety check.** If a `{{Placeholder}}` in the template has no matching
  column, the tool leaves it untouched and prints a warning listing it, so
  nothing goes out half-finished by mistake.

---

## Files in this project

| File / folder                    | What it is                                             |
|----------------------------------|--------------------------------------------------------|
| `fill_docs.py`                   | The command-line script.                               |
| `docs_filler/core.py`            | The shared filling logic (used by the script and app). |
| `webapp/`                        | The web app (FastAPI) — see below.                     |
| `Dockerfile`                     | Builds the web app into a container for the cloud.     |
| `deploy/README-deploy.md`        | Step-by-step Azure + Entra ID deployment guide.        |
| `requirements.txt`               | The libraries it needs.                                |
| `input/Offer_Template.docx`      | Example template using the `{{ }}` convention.         |
| `input/content.xlsx`             | Example spreadsheet (one header row, one row/person).  |
| `input/`                         | Put your template and spreadsheet here.                |
| `output/`                        | Where the command-line script writes finished docs.    |

---

## Run it as a web app (no install for your colleagues)

There's also a small web app so non-technical colleagues can use it from a
browser: upload the template and spreadsheet, watch a progress bar, and download
a single ZIP of the finished documents. **Files are processed in memory and never
saved on the server.**

### Try it locally

```
pip install -r requirements.txt
uvicorn webapp.main:app --reload
```

Then open <http://127.0.0.1:8000> in your browser.

### Deploy it for the whole team (Azure + Entra ID sign-in)

The app is designed to run on **Azure Container Apps**:

- **Scales to zero** — it costs ≈ £0 when nobody is using it and spins up on
  demand.
- **Microsoft Entra ID sign-in** — using 2i's existing Entra tenant, so only 2i
  staff can open it (handled by Azure, no passwords in the app).
- **No stored data** — no storage account or database is attached; files exist
  only in memory during a job.

Deployment is a one-time manual setup plus an automated pipeline for updates:

- **One-time setup.** Create the Azure resources (resource group, registry,
  Container Apps environment, the app, and Entra sign-in) with step-by-step `az`
  commands in [`deploy/README-deploy.md`](deploy/README-deploy.md).
- **Automated updates: CI/CD.** `azure-pipelines.yml` is the **Azure DevOps**
  pipeline: on a merge to `main` it runs the tests and, only if they pass, builds
  the image in ACR and rolls out a new revision — using the Azure CLI, so the
  pipeline only needs **Contributor** on the resource group. See
  [`deploy/README-pipeline.md`](deploy/README-pipeline.md).

## Tests

```
pip install -r requirements-dev.txt
pytest
```

The suite covers the core filling logic (including placeholders split across
runs, number formatting, and error cases) and the web endpoints end-to-end.
