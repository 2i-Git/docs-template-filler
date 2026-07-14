# Deploying Docs Filler to Azure Static Web Apps (free)

The `site/` folder is a fully self-contained, browser-only version of the app —
all document filling happens client-side, so it's just static files. That means
it runs on the **Azure Static Web Apps free tier** (custom domain, free SSL, and
a global CDN included) and publishes straight from the Azure DevOps pipeline.

You do **steps 1–3 once**. After that, every push to `main` deploys automatically.

---

## Prerequisites

- An Azure subscription (the Static Web App itself is free).
- The [Azure CLI](https://learn.microsoft.com/cli/azure/install-azure-cli), or
  just the Azure Portal.
- This repo connected to an Azure DevOps pipeline that runs `azure-pipelines.yml`.

---

## 1. Create the free Static Web App (once)

### Option A — Azure CLI (fastest)

```bash
# Create the resource group first if you don't have one:
az group create --name docs-filler-rg --location westeurope

az staticwebapp create \
  --name docs-filler \
  --resource-group docs-filler-rg \
  --location westeurope \
  --sku Free
```

- `--location` is only where the app's metadata lives — content is served from
  the global CDN regardless.
- **Don't** pass `--source` / `--repository-url`. Leaving them off creates a
  "manual / Other" deployment that publishes from your Azure DevOps pipeline
  instead of GitHub.

### Option B — Azure Portal

Create a resource → **Static Web App** → **Plan type: Free** → **Deployment
source: Other** → **Review + create**.

---

## 2. Copy the deployment token

```bash
az staticwebapp secrets list \
  --name docs-filler \
  --resource-group docs-filler-rg \
  --query "properties.apiKey" -o tsv
```

(Portal: the app's **Overview** → **Manage deployment token**.)

---

## 3. Add the token to the pipeline (once)

In Azure DevOps → **Pipelines** → your pipeline → **Edit** → **Variables** →
**New variable**:

- **Name:** `deployment_token`
- **Value:** the token from step 2
- ✅ **Keep this value secret**

The pipeline (`azure-pipelines.yml`) already reads `$(deployment_token)` — there
is nothing else to wire up.

---

## 4. Deploy

Merge / push to `main`. The pipeline runs two stages:

1. **Test** — Bandit + `pytest` over the retained Python.
2. **Deploy** (only on a real merge to `main`, not PR builds) — the
   `AzureStaticWebApp@0` task publishes the `site/` folder. No container, no ACR.

---

## 5. Verify

The deploy step prints the live URL in its log, or fetch it with:

```bash
az staticwebapp show \
  --name docs-filler \
  --resource-group docs-filler-rg \
  --query "defaultHostname" -o tsv
```

Open `https://<that-host>/`, upload `input/Offer_Template.docx` +
`input/content.xlsx`, and confirm that `Offer_Template.zip` downloads with one
document per row inside.

---

## Optional

### Custom domain + free SSL

In the app's **Custom domains** blade, add your domain, create the CNAME record
it shows you, and SSL is provisioned automatically.

### Try it locally before you push

The site has no build step — serve the folder with any static file server:

```bash
python -m http.server -d site 8080
```

Then open <http://127.0.0.1:8080>. (Opening `index.html` directly via `file://`
won't work — the vendored scripts need to load over http.)

---

## Notes

- **Access is public** (anyone with the URL can use it). That's safe here because
  it's a pure client-side tool that stores nothing — files never leave the
  user's browser. Restricting sign-in to your own Entra tenant would require the
  Static Web Apps **Standard** plan (a custom Entra provider); the free
  preconfigured provider can't be tenant-locked. If you need org-only access,
  keep using the FastAPI container app instead (see `README.md`).
- **The pipeline triggers on `main`.** The `site/` folder, the updated
  `azure-pipelines.yml`, and the README changes must be committed and merged to
  `main` before the first deploy can run. The pipeline YAML also has to exist on
  the branch the pipeline is configured to watch.
- **Decommissioning the old container app** (optional): once the Static Web App
  is live you can delete the Container App / ACR resources to stop those costs.
  This repo keeps the container code in place but no longer deploys it.
