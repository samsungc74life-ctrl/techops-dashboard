import fs from 'node:fs/promises';
import path from 'node:path';
import XLSX from 'xlsx';

const SHAREPOINT_URL = 'https://trackingco-my.sharepoint.com/:x:/g/personal/data05_iotcube_sa/IQB8yhdEMfeHTLpHIDK1DCMJAVprCHIt-s5rs16XXjgOE_Y?e=Q59voD';
const source = new URL(SHAREPOINT_URL);
const variants = [source.toString()];
const download = new URL(source); download.searchParams.set('download','1'); variants.push(download.toString());
const clean = new URL(source); clean.search='?download=1'; variants.push(clean.toString());
let buffer, errors=[];
for (const url of [...new Set(variants)]) {
  try {
    const response = await fetch(url,{redirect:'follow',headers:{'User-Agent':'Mozilla/5.0 (compatible; TechOpsDashboardSync/1.0)','Accept':'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/octet-stream,*/*'},signal:AbortSignal.timeout(45000)});
    if (!response.ok) throw new Error('HTTP '+response.status);
    const bytes=Buffer.from(await response.arrayBuffer());
    if (bytes.length<4 || bytes[0]!==0x50 || bytes[1]!==0x4b) throw new Error('Response is not an XLSX ZIP file ('+(response.headers.get('content-type')||'unknown')+')');
    buffer=bytes; break;
  } catch(e) { errors.push(e.message); }
}
if (!buffer) throw new Error('SharePoint download failed: '+errors.join(' | '));
const workbook=XLSX.read(buffer,{type:'buffer',cellDates:true,raw:false});
const sheets=workbook.SheetNames.map(name=>({name,rows:XLSX.utils.sheet_to_json(workbook.Sheets[name],{defval:'',raw:false})}));
if (!sheets.length) throw new Error('Workbook contains no worksheets.');
const payload={workbookName:'SharePoint Excel workbook',fetchedAt:new Date().toISOString(),sheets};
const out=path.resolve('site/data'); await fs.mkdir(out,{recursive:true}); await fs.writeFile(path.join(out,'workbook.json'),JSON.stringify(payload));
console.log('Prepared '+sheets.length+' worksheets and '+sheets.reduce((n,s)=>n+s.rows.length,0)+' rows.');
