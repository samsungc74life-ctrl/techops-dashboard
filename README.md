# Technical Operations Dashboard

Public static dashboard hosted on GitHub Pages. GitHub Actions downloads the SharePoint workbook through Microsoft Graph, converts every worksheet to JSON, and publishes the site.

## One-time Microsoft setup

The SharePoint sharing page is not a direct XLSX download: it returns HTML. Automated download therefore uses Microsoft Graph authentication.

Ask your Microsoft 365 / Entra administrator to:
1. Register an application in Microsoft Entra ID.
2. Add Microsoft Graph **Application** permission sufficient to read the workbook (typically `Files.Read.All` or `Sites.Read.All`, depending on tenant policy) and grant admin consent. If the tenant's shared-item endpoint requires a broader permission, the administrator should confirm the least-privileged supported option.
3. Provide the Tenant ID, Application (client) ID, and a client secret.

Then add these repository Actions secrets under Settings → Secrets and variables → Actions:
- `MS_TENANT_ID`
- `MS_CLIENT_ID`
- `MS_CLIENT_SECRET`
- `SHAREPOINT_SHARE_URL` (the complete Anyone-with-the-link view URL)

Never commit the client secret or other credentials to this public repository.

## Publish

1. In Settings → Pages, choose **GitHub Actions** as the source.
2. Run **Sync SharePoint workbook and publish dashboard** from the Actions tab.
3. After a successful run, open Settings → Pages for the public site URL.

The workflow is scheduled every five minutes, but GitHub may delay scheduled runs. The browser refreshes the published snapshot every five minutes; that does not itself download the workbook.

## Public data notice

This repository and website are public. Every populated row and column from every worksheet is published, including IMEIs and other workbook fields. Anyone may view, copy, or redistribute the published data. Publish only with the workbook owner's authorization.
