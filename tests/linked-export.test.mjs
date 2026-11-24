import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {unzipSync,strFromU8} from 'fflate';
import {DOMParser} from '@xmldom/xmldom';
import worker from '../dist/server/index.js';
import {demoTemplate,generateReport} from '../src/docx.js';
import {writeReportZip} from '../web/batch-reports.js';

function memoryBucket(){
 const data=new Map();
 return {data,put:async(k,v)=>data.set(k,v),get:async k=>data.has(k)?{body:data.get(k),json:async()=>JSON.parse(data.get(k)),arrayBuffer:async()=>data.get(k)}:null,
  list:async({prefix})=>({objects:[...data.keys()].filter(k=>k.startsWith(prefix)).map(key=>({key}))}),
  delete:async keys=>{for(const key of [keys].flat())data.delete(key);}};
}
const jpeg=await readFile(new URL('./fixtures/test-photo.jpg',import.meta.url));
test('consolidates the unlinked master, retains its photos and history, and exports each linked site once',async()=>{
 const b=memoryBucket(),api=(path,opts)=>worker.fetch(new Request('https://app.test/api'+path,opts),{BUCKET:b});
 const put=(key,value)=>b.put(key,JSON.stringify(value));
 const site={name:'Abu Dhabi',code:'82A13402',templateName:'Ajwan.docx',mapping:{},timeZone:'Asia/Dubai'};
 for(const [id,link] of [['old',''],['keep','42'],['other','43']]){
  const l={...site,id,companyCamProjectId:link,templateKey:'templates/'+id+'/v1.docx'};
  if(id==='other')l.code='82A13403';
  await put('locations/'+id+'.json',l);await b.put(l.templateKey,demoTemplate());
 }
 await put('reports/history.json',{id:'history',projectId:'old',createdAt:'2026-01-01'});
 for(const id of ['old','other']){
  await put('photos/'+id+'/'+id+'.json',{id,projectId:id,relevant:true,capturedAt:'2026-11-10T12:00:00Z',caption:'Manual Watering',width:320,height:200});
  await b.put('images/'+id+'.jpg',jpeg);
 }
 assert.equal((await (await api('/state')).json()).duplicateMasters,1);
 await api('/masters/consolidate',{method:'POST'});
 assert.equal(b.data.has('locations/old.json'),false);assert.equal(b.data.has('templates/old/v1.docx'),false);
 assert.ok(b.data.has('reports/history.json'));assert.ok(b.data.has('photos/keep/old.json'));
 const state=await (await api('/state')).json();assert.equal(state.duplicateMasters,0);assert.equal(state.locations.length,2);
 const f=new FormData();f.append('template',new File([demoTemplate()],'Ajwan.docx'));f.append('metadata',JSON.stringify(site));
 assert.equal((await (await api('/locations',{method:'POST',body:f})).json()).id,'keep');
 const batchId=crypto.randomUUID(),entries=[];
 for(const l of state.locations){
  const post=()=>api('/reports',{method:'POST',body:JSON.stringify({projectId:l.id,month:'2026-11',linkedProjectId:l.companyCamProjectId,batchId,reviewed:true})});
  const report=await (await post()).json();assert.equal(report.summary,'Manual Watering.');assert.equal((await (await post()).json()).id,report.id);
  const response=await api('/reports/'+report.id+'/download');
  entries.push({path:l.code+'.docx',url:'data:application/octet-stream;base64,'+Buffer.from(await response.arrayBuffer()).toString('base64')});
 }
 const chunks=[];await writeReportZip(entries,chunk=>chunks.push(Buffer.from(chunk)));
 const files=unzipSync(Buffer.concat(chunks));assert.equal(Object.keys(files).length,2);
 for(const file of Object.values(files))assert.match(strFromU8(unzipSync(file)['word/document.xml']),/Manual Watering/);
});

test('photo cells use fixed frames, opposite crops for portrait and landscape, and aligned captions',()=>{
 const photos=[{width:900,height:400},{width:300,height:900},{width:400,height:400},{width:800,height:600},{width:300,height:600}].map((dimensions,i)=>({...dimensions,bytes:jpeg,caption:'Activity '+i}));
 const file=generateReport(demoTemplate(),{},{LOCATION:'Test',REPORT_PERIOD:'November 2026',WORK_SUMMARY:'Work'},photos);
 const xml=strFromU8(unzipSync(file)['word/document.xml']),doc=new DOMParser().parseFromString(xml,'application/xml');
 const nodes=(el,tag)=>Array.from(el.getElementsByTagNameNS('*',tag));
 const boxes=nodes(doc,'extent'),crops=nodes(doc,'srcRect');assert.equal(boxes.length,5);
 assert.equal(new Set(boxes.map(b=>b.getAttribute('cx')+'/'+b.getAttribute('cy'))).size,1);
 assert.ok(Number(crops[0].getAttribute('l'))>0);assert.ok(Number(crops[1].getAttribute('t'))>0);
 const rows=nodes(doc,'tr');assert.deepEqual(rows.map(r=>nodes(r,'tc').length),[4,4,4,4]);
 assert.equal(nodes(rows[1],'tc')[2].textContent,'Activity 2');assert.equal(nodes(rows[3],'tc')[1].textContent,'');
});
