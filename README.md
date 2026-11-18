# Technical Operations Dashboard

Public, static TechOps dashboard hosted on GitHub Pages. A GitHub Actions workflow downloads the configured SharePoint workbook, converts every worksheet to JSON, and publishes the dashboard.

## Setup
1. In repository Settings → Pages, set the build and deployment source to **GitHub Actions**.
2. Open Actions and run **Sync SharePoint workbook and publish dashboard** once using Run workflow.
3. After a successful deployment, open Settings → Pages for the public URL.

The workflow is scheduled every five minutes, but GitHub may delay scheduled jobs. The dashboard checks its published data snapshot every five minutes; that browser refresh does not itself trigger a workbook download.

## Public data notice
This repository and the deployed website are public. The sync publishes every populated row and column from every worksheet, including IMEIs and any other workbook fields. The source SharePoint URL is an Anyone link and is stored in the workflow script. Publish only with the workbook owner's authorization. Do not add passwords, tokens, or private credentials.

Node.js is used only on GitHub-hosted runners; no local server or local Node installation is required.