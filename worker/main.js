import {excelToWordTemplate,generateExcelReport} from '../src/xlsx.js';
import {inspectTemplate,generateReport,appendPhotoSection} from '../src/docx.js';
import {selectPhotos,currentMonth,periodLabel,validMonth} from '../src/policy.js';
const json=(v,status=200)=>new Response(JSON.stringify(v),{status,headers:{'Content-Type':'application/json','Cache-Control':'no-store'}});
const get=async(b,k)=>(await b.get(k))?.json();
const put=(b,k,v)=>b.put(k,JSON.stringify(v),{httpMetadata:{contentType:'application/json'}});
async function records(b,prefix){let cursor,keys=[];do{const r=await b.list({prefix,cursor,limit:1000});keys.push(...r.objects);cursor=r.truncated?r.cursor:undefined;}while(cursor);const out=[];for(let i=0;i<keys.length;i+=20)out.push(...await Promise.all(keys.slice(i,i+20).map(x=>get(b,x.key))));return out.filter(Boolean);}
async function handle(req,env){
 const url=new URL(req.url),path=url.pathname;
 const assets={'/':[__HTML__,'text/html'],'/app.js':[__CLIENT__,'text/javascript'],'/style.css':[__CSS__,'text/css'],'/manual':[__MANUAL__,'text/html']};
 if(assets[path])return new Response(assets[path][0],{headers:{'Content-Type':assets[path][1]+'; charset=utf-8','Cache-Control':'no-store','X-Content-Type-Options':'nosniff'}});
 if(!path.startsWith('/api/'))return json({error:'Not found'},404);
 if(req.method!=='GET'&&req.headers.get('Origin')&&req.headers.get('Origin')!==url.origin)return json({error:'Cross-site request rejected'},403);
 const b=env.BUCKET;if(!b)return json({error:'Project storage is not configured.'},503);
 if(path==='/api/state')return json({locations:(await records(b,'locations/')).sort((a,b)=>a.name.localeCompare(b.name)),reports:(await records(b,'reports/')).sort((a,b)=>b.createdAt.localeCompare(a.createdAt)),currentMonth:currentMonth(),companyCamConnected:false});
 if(path==='/api/templates/inspect'&&req.method==='POST'){const bytes=new Uint8Array(await req.arrayBuffer()),excel=req.headers.get('x-template-format')==='xlsx';return json({...inspectTemplate(excel?excelToWordTemplate(bytes,'Monthly reporting log'):bytes),excel});}
 if(path==='/api/locations'&&req.method==='POST'){
  const f=await req.formData(),file=f.get('template'),m=JSON.parse(f.get('metadata')||'{}');
  if(!m.name?.trim()||m.name.length>200)throw Error('Enter a project name under 200 characters.');new Intl.DateTimeFormat('en',{timeZone:m.timeZone});
  if(!(file instanceof File)||!(/\.(docx|xlsx)$/i.test(file.name)))throw Error('Upload a Word .docx master template.');
  const original=new Uint8Array(await file.arrayBuffer()),isExcel=/\.xlsx$/i.test(file.name);let bytes=isExcel?excelToWordTemplate(original,m.name):original;let inspection=inspectTemplate(bytes);if(!inspection.bodyImageCount&&!inspection.tokens.includes('{{PHOTOS}}')){const donor=(await records(b,'locations/')).find(l=>l.inspection?.photoTables?.length);if(donor){const f=await b.get(donor.templateKey);bytes=appendPhotoSection(bytes,new Uint8Array(await f.arrayBuffer()));inspection=inspectTemplate(bytes);m.mapping={...m.mapping,photoTableIndices:inspection.tables.filter(t=>t.images>=2).map(t=>t.index)};}}const id=m.id&&/^[\w-]+$/.test(m.id)?m.id:crypto.randomUUID();
  const old=await get(b,`locations/${id}.json`),mapping=m.mapping||{photoTableIndices:inspection.tables.filter(t=>t.images>=2).map(t=>t.index),periodText:m.periodText||''};
  if(!mapping.photoTableIndices?.length&&!inspection.tokens.includes('{{PHOTOS}}')&&!Number.isInteger(mapping.photoIndex))throw Error('No photo layout found. Map a photo table or add a {{PHOTOS}} paragraph in Word.');
  if(mapping.periodText&&!Number.isInteger(mapping.periodIndex)){const p=inspection.paragraphs.find(p=>p.text.includes(mapping.periodText));if(p)mapping.periodIndex=p.index;}const l={id,format:isExcel?'xlsx':'docx',excelTemplateKey:isExcel?`excel-templates/${id}.xlsx`:null,name:m.name.trim(),client:m.client||'',code:m.code||'',companyCamProjectId:m.companyCamProjectId||'',timeZone:m.timeZone,templateName:file.name,templateKey:`templates/${id}/${crypto.randomUUID()}.docx`,mapping,createdAt:old?.createdAt||new Date().toISOString(),updatedAt:new Date().toISOString(),inspection:{bodyImageCount:inspection.bodyImageCount,photoTables:inspection.tables.filter(t=>mapping.photoTableIndices?.includes(t.index))}};
  if(isExcel)await b.put(l.excelTemplateKey,original);await b.put(l.templateKey,bytes);await put(b,`locations/${id}.json`,l);return json(l,201);
 }
 const lm=path.match(/^\/api\/locations\/([\w-]+)\/(photos|template)$/);
 if(lm){const [,id,kind]=lm,l=await get(b,`locations/${id}.json`);if(!l)return json({error:'Project not found'},404);
  if(kind==='template'){const f=await b.get(l.templateKey);return new Response(f.body,{headers:{'Content-Type':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':'attachment; filename="master-template.docx"'}});}
  if(req.method==='GET')return json(await records(b,`photos/${id}/`));
  if(req.method==='POST'){
   const f=await req.formData(),file=f.get('photo'),m=JSON.parse(f.get('metadata')||'{}');
   if(!(file instanceof File)||file.size>2000000)throw Error('Photo must be a normalized JPEG under 2 MB.');
   const bytes=new Uint8Array(await file.arrayBuffer());if(bytes[0]!==255||bytes[1]!==216)throw Error('Use a JPEG photo.');
   if(!m.capturedAt||!Number.isFinite(new Date(m.capturedAt).getTime()))throw Error('Set the capture date. Upload dates are not capture dates.');
   if(!Number.isInteger(m.width)||!Number.isInteger(m.height)||m.width<1||m.height<1||m.width>2400||m.height>2400)throw Error('Invalid photo dimensions.');
   const hash=Array.from(new Uint8Array(await crypto.subtle.digest('SHA-256',bytes))).map(x=>x.toString(16).padStart(2,'0')).join('');
   const p={id:crypto.randomUUID(),projectId:id,name:String(m.name||file.name).slice(0,250),caption:String(m.caption||'').slice(0,500),activity:String(m.activity||'Work activity').slice(0,100),capturedAt:new Date(m.capturedAt).toISOString(),relevant:m.relevant===true,width:m.width,height:m.height,hash,source:'manual'};
   await b.put(`images/${p.id}.jpg`,bytes);await put(b,`photos/${id}/${p.id}.json`,p);return json(p,201);
  }
 }
 const pm=path.match(/^\/api\/photos\/([\w-]+)\/([\w-]+)$/);
 if(pm){const [,project,id]=pm,k=`photos/${project}/${id}.json`,p=await get(b,k);if(!p)return json({error:'Photo not found'},404);
  if(req.method==='PATCH'){const d=await req.json();for(const k of ['caption','activity','relevant','capturedAt'])if(d[k]!==undefined)p[k]=d[k];if(typeof p.relevant!=='boolean'||!p.capturedAt||!Number.isFinite(new Date(p.capturedAt).getTime()))throw Error('Invalid photo details.');p.caption=String(p.caption||'').slice(0,500);p.activity=String(p.activity||'').slice(0,100);await put(b,k,p);return json(p);}
  if(req.method==='DELETE'){await b.delete([k,`images/${id}.jpg`]);return json({ok:true});}
 }
 const im=path.match(/^\/api\/images\/([\w-]+)$/);if(im){const f=await b.get(`images/${im[1]}.jpg`);return f?new Response(f.body,{headers:{'Content-Type':'image/jpeg','Cache-Control':'private, max-age=3600'}}):json({error:'Not found'},404);}
 if(path==='/api/reports'&&req.method==='POST'){
  const d=await req.json();if(!/^[\w-]+$/.test(d.projectId||'')||!validMonth(d.month))throw Error('Choose a project and report month.');if(d.reviewed!==true)throw Error('Review monthly details before generating.');
  const l=await get(b,`locations/${d.projectId}.json`);if(!l)return json({error:'Project not found'},404);
  const selection=selectPhotos(await records(b,`photos/${l.id}/`),l.id,d.month,l.timeZone);if(!selection.photos.length)throw Error('No eligible photos for this project and month. Add captions and mark relevant work photos.');if(!d.summary?.trim())throw Error('Enter this month’s work summary.');
  const photos=[];for(const p of selection.photos){const f=await b.get(`images/${p.id}.jpg`);if(!f)throw Error('Selected photo is missing.');photos.push({...p,bytes:new Uint8Array(await f.arrayBuffer())});}
  const t=await b.get(l.templateKey),period=periodLabel(d.month),values={REPORT_PERIOD:period,REPORT_MONTH:period.split(' ')[0],REPORT_YEAR:d.month.slice(0,4),LOCATION:l.name,WORK_SUMMARY:d.summary.trim()};
  const bytes=generateReport(new Uint8Array(await t.arrayBuffer()),l.mapping,values,photos),id=crypto.randomUUID();
  const record={id,projectId:l.id,projectName:l.name,month:d.month,period,photoCount:photos.length,relevantCount:selection.relevantCount,limit:selection.limit,photoIds:photos.map(p=>p.id),summary:d.summary.trim(),templateName:l.templateName,createdAt:new Date().toISOString(),filename:`${l.name.replace(/[^a-zA-Z0-9_-]/g,'_')}_${d.month}.docx`};
  if(l.excelTemplateKey){const original=await b.get(l.excelTemplateKey);const excel=generateExcelReport(new Uint8Array(await original.arrayBuffer()),period,photos,l.name);await b.put(`documents/${id}.xlsx`,excel);record.excelFilename=record.filename.replace(/\.docx$/,'.xlsx');}await b.put(`documents/${id}.docx`,bytes);await put(b,`reports/${id}.json`,record);return json(record,201);
 }
 const rm=path.match(/^\/api\/reports\/([\w-]+)\/(download|excel)$/);if(rm){const r=await get(b,`reports/${rm[1]}.json`);if(!r)return json({error:'Report not found'},404);const excel=rm[2]==='excel';if(excel&&!r.excelFilename)return json({error:'No Excel output for this report'},404);const f=await b.get(`documents/${r.id}.${excel?'xlsx':'docx'}`);return new Response(f.body,{headers:{'Content-Type':excel?'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':'application/vnd.openxmlformats-officedocument.wordprocessingml.document','Content-Disposition':`attachment; filename="${excel?r.excelFilename:r.filename}"`,'Cache-Control':'no-store'}});}
 return json({error:'Not found'},404);
}
export default {async fetch(req,env){try{return await handle(req,env);}catch(e){console.error(e.message);return json({error:e.message||'Operation failed'},400);}}};
