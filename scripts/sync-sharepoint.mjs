import fs from 'node:fs/promises';
import path from 'node:path';
import XLSX from 'xlsx';

const tenantId = process.env.MS_TENANT_ID;
const clientId = process.env.MS_CLIENT_ID;
const clientSecret = process.env.MS_CLIENT_SECRET;
const shareUrl = process.env.SHAREPOINT_SHARE_URL;
if (!tenantId || !clientId || !clientSecret || !shareUrl) {
  throw new Error('Missing GitHub Actions secrets. Configure MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET and SHAREPOINT_SHARE_URL.');
}

async function getAccessToken() {
  const body = new URLSearchParams({
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default',
    grant_type: 'client_credentials'
  });
  const response = await fetch(`https://login.microsoftonline.com/${encodeURIComponent(tenantId)}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body,
    signal: AbortSignal.timeout(30000)
  });
  const data = await response.json();
  if (!response.ok || !data.access_token) {
    throw new Error('Microsoft identity token request failed: ' + (data.error_description || response.status));
  }
  return data.access_token;
}

function encodeSharingUrl(url) {
  return 'u!' + Buffer.from(url, 'utf8').toString('base64')
    .replace(/=/g, '').replace(/\\+/g, '-').replace(/\\//g, '_');
}

const token = await getAccessToken();
const shareToken = encodeSharingUrl(shareUrl);
const response = await fetch(`https://graph.microsoft.com/v1.0/shares/${shareToken}/driveItem/content`, {
  headers: { Authorization: `Bearer ${token}` },
  redirect: 'follow',
  signal: AbortSignal.timeout(60000)
});
if (!response.ok) {
  const detail = (await response.text()).slice(0, 500);
  throw new Error(`Microsoft Graph workbook download failed (HTTP ${response.status}): ${detail}`);
}
const buffer = Buffer.from(await response.arrayBuffer());
if (buffer.length < 4 || buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
  throw new Error('Microsoft Graph returned a non-XLSX response. Check the SharePoint sharing link and app permissions.');
}

const workbook = XLSX.read(buffer, { type: 'buffer', cellDates: true, raw: false });
const sheets = workbook.SheetNames.map(name => ({
  name,
  rows: XLSX.utils.sheet_to_json(workbook.Sheets[name], { defval: '', raw: false })
}));
if (!sheets.length) throw new Error('Workbook contains no worksheets.');
const payload = { workbookName: 'SharePoint Excel workbook', fetchedAt: new Date().toISOString(), sheets };
const out = path.resolve('site/data');
await fs.mkdir(out, { recursive: true });
await fs.writeFile(path.join(out, 'workbook.json'), JSON.stringify(payload));
console.log('Prepared ' + sheets.length + ' worksheets and ' + sheets.reduce((n, s) => n + s.rows.length, 0) + ' rows.');
