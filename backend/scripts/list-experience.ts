import { PrismaClient } from '@prisma/client';
import { getValidAccessToken } from '../src/lib/graphAuth.ts';

const prisma = new PrismaClient();

async function fetchJson(url:string, token:string){
  const res = await fetch(url, {headers:{Authorization:`Bearer ${token}`}});
  const txt = await res.text();
  let data:any;
  try{data=JSON.parse(txt);}catch{data=txt;}
  return {status:res.status, data};
}

(async()=>{
  const token = await getValidAccessToken();
  const wb = await prisma.excelWorkbook.findFirst();
  if(!wb){console.error('Workbook not found');process.exit(1);} 
  const realId = wb.driveItemId;
  const base = `https://graph.microsoft.com/v1.0/me/drive/items/${realId}/workbook`;
  const tblRes = await fetchJson(`${base}/tables`, token);
  if(tblRes.status!==200){console.error('tables error',tblRes.status);process.exit(1);} 
  const tables = tblRes.data.value || [];
  console.log('Tables:', tables.map((t:any)=>({name:t.name,id:t.id})));
  for(const t of tables){
    const colsRes = await fetchJson(`${base}/tables/${t.id}/columns`, token);
    console.log(`Columns for table ${t.name}:`, (colsRes.data.value||[]).map((c:any)=>c.name));
  }
  await prisma.$disconnect();
})();

