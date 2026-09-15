const crypto = require('crypto');

const ENV = ['SP_TENANT_ID','SP_CLIENT_ID','SP_CLIENT_SECRET','SP_SITE_ID','SP_LIST_ID'];

function json(res, status, body) {
  res.status(status).setHeader('Content-Type','application/json; charset=utf-8').send(JSON.stringify(body));
}
function clean(v, max=120) { return String(v ?? '').trim().replace(/[<>]/g,'').slice(0,max); }
function int(v) { const n=Number(v); return Number.isInteger(n) && n>=0 ? n : 0; }
function hash(record) { return crypto.createHash('sha256').update(JSON.stringify(record)).digest('hex'); }

async function token() {
  const body = new URLSearchParams({client_id:process.env.SP_CLIENT_ID,client_secret:process.env.SP_CLIENT_SECRET,scope:'https://graph.microsoft.com/.default',grant_type:'client_credentials'});
  const r = await fetch(`https://login.microsoftonline.com/${process.env.SP_TENANT_ID}/oauth2/v2.0/token`, {method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body});
  if(!r.ok) throw new Error('Microsoft authentication failed');
  return (await r.json()).access_token;
}
async function graph(path, options={}) {
  const t=await token();
  const r=await fetch(`https://graph.microsoft.com/v1.0${path}`,{...options,headers:{Authorization:`Bearer ${t}`,'Content-Type':'application/json',...(options.headers||{})}});
  const text=await r.text(); let data={}; try{data=JSON.parse(text)}catch{}
  if(!r.ok){const e=new Error(data?.error?.message||`SharePoint request failed (${r.status})`);e.status=r.status;throw e}
  return data;
}
async function listItems(filter='') {
  const q=filter?`&$filter=${encodeURIComponent(filter)}`:'';
  const data=await graph(`/sites/${encodeURIComponent(process.env.SP_SITE_ID)}/lists/${encodeURIComponent(process.env.SP_LIST_ID)}/items?$expand=fields&$top=200${q}`);
  return (data.value||[]).map(x=>({id:String(x.id),...x.fields}));
}
function toRecord(x){return {id:String(x.id),technicianId:x.TechnicianID||'',technicianName:x.TechnicianName||'',punchIn:x.PunchIn||null,punchOut:x.PunchOut||null,workingMinutes:int(x.WorkingMinutes),status:x.Status||'Present',location:x.Location||'',deviceInfo:x.DeviceInfo||'',repaired:int(x.BatteriesRepaired),replaced:int(x.CasesReplaced),recordHash:x.RecordHash||''};}
async function createRecord(r){
  const fields={Title:`${r.technicianId} - ${r.technicianName}`,TechnicianID:r.technicianId,TechnicianName:r.technicianName,PunchIn:r.punchIn,PunchOut:null,WorkingMinutes:0,Status:'Present',Location:r.location||'',DeviceInfo:r.deviceInfo||'',BatteriesRepaired:r.repaired,CasesReplaced:r.replaced};
  fields.RecordHash=hash(fields);
  const data=await graph(`/sites/${encodeURIComponent(process.env.SP_SITE_ID)}/lists/${encodeURIComponent(process.env.SP_LIST_ID)}/items`,{method:'POST',body:JSON.stringify({fields})});
  return toRecord({id:data.id,...data.fields});
}
async function updateRecord(id, r){
  const fields={PunchOut:r.punchOut,WorkingMinutes:r.workingMinutes,Status:'Completed',BatteriesRepaired:r.repaired,CasesReplaced:r.replaced,Location:r.location||'',DeviceInfo:r.deviceInfo||''};
  fields.RecordHash=hash(fields);
  const data=await graph(`/sites/${encodeURIComponent(process.env.SP_SITE_ID)}/lists/${encodeURIComponent(process.env.SP_LIST_ID)}/items/${encodeURIComponent(id)}/fields`,{method:'PATCH',body:JSON.stringify(fields)});
  return toRecord({id,...data});
}
async function handler(req,res){
  if(req.method!=='POST') return json(res,405,{error:'Method not allowed'});
  if(!ENV.every(k=>process.env[k])) return json(res,503,{error:'SharePoint is not configured. Add the required Vercel environment variables.'});
  try{
    const body=typeof req.body==='string'?JSON.parse(req.body):req.body||{};
    const action=clean(body.action,30);
    if(action==='list'){
      const rows=await listItems();
      rows.sort((a,b)=>new Date(b.PunchIn||0)-new Date(a.PunchIn||0));
      return json(res,200,{records:rows.map(toRecord)});
    }
    if(!['punch-in','punch-out'].includes(action)) return json(res,400,{error:'Invalid attendance action'});
    const technicianId=clean(body.technicianId,40),technicianName=clean(body.technicianName,100);
    if(!technicianId||!technicianName) return json(res,400,{error:'Technician ID and Technician Name are required.'});
    const now=new Date();
    const repaired=int(body.repaired), replaced=int(body.replaced), location=clean(body.location,200), deviceInfo=clean(body.deviceInfo,300);
    const existing=(await listItems(`fields/TechnicianID eq '${technicianId.replace(/'/g,"''")}'`)).sort((a,b)=>new Date(b.PunchIn||0)-new Date(a.PunchIn||0));
    const active=existing.find(x=>!x.PunchOut);
    if(action==='punch-in'){
      if(active) return json(res,409,{error:'Duplicate Punch-In blocked. This technician already has an open attendance record.',record:toRecord(active)});
      const record=await createRecord({technicianId,technicianName,punchIn:now.toISOString(),location,deviceInfo,repaired,replaced});
      return json(res,200,{record});
    }
    if(!active) return json(res,409,{error:'No open Punch-In was found for this technician.'});
    const punchIn=new Date(active.PunchIn);
    if(Number.isNaN(punchIn.getTime())||now<=punchIn) return json(res,409,{error:'Invalid attendance sequence. Punch-Out must be later than Punch-In.'});
    const workingMinutes=Math.floor((now-punchIn)/60000);
    if(workingMinutes<0 || workingMinutes>24*60) return json(res,409,{error:'Unrealistic working duration detected. Attendance was not updated.'});
    const record=await updateRecord(active.id,{punchOut:now.toISOString(),workingMinutes,repaired,replaced,location:location||active.Location,deviceInfo:deviceInfo||active.DeviceInfo});
    return json(res,200,{record});
  }catch(e){console.error(e);return json(res,e.status===429?429:500,{error:e.message||'Unable to process attendance.'})}
}
module.exports=handler;
