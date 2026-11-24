import test from 'node:test';import assert from 'node:assert/strict';import {readFile} from 'node:fs/promises';
import {zipSync,unzipSync} from 'fflate';import {demoTemplate} from '../src/docx.js';import worker from '../dist/server/index.js';import {createHandler} from '../server/vercel-handler.js';
test('Vercel adapter completes large master upload, photo selection, persistence and streamed Word download',async()=>{
 const values=new Map();const storage={
  put:async(path,value,options)=>{assert.equal(options.access,'private');values.set(path,typeof value==='string'?new TextEncoder().encode(value):new Uint8Array(value));return {pathname:path};},
  get:async(path,options)=>{assert.equal(options.access,'private');const bytes=values.get(path);return bytes?{statusCode:200,stream:new Response(bytes).body,blob:{size:bytes.length,contentType:'application/octet-stream'}}:null;},
  list:async({prefix})=>({blobs:[...values.keys()].filter(k=>k.startsWith(prefix)).map(pathname=>({pathname})),hasMore:false}),
  del:async(paths)=>{for(const p of Array.isArray(paths)?paths:[paths])values.delete(p);}
 };
 const env={BLOB_READ_WRITE_TOKEN:'synthetic'};const handler=createHandler({worker,storage,env});
 const call=(path,opts={})=>handler(new Request('https://app.test'+path,{...opts,headers:{...opts.headers}}));
 for(const p of ['/','/api/index','/api/index?__route=']){const home=await call(p);assert.equal(home.status,200);assert.ok((await home.text()).includes('Field Report Agent'));}
 const z=unzipSync(demoTemplate());z['word/media/unused.png']=new Uint8Array(5*1024*1024);const bytes=zipSync(z,{level:0});assert.ok(bytes.length>4.5*1024*1024);
 const path='incoming/12345678-abcd.docx';values.set(path,bytes);
 const result=await call('/api/locations',{method:'POST',headers:{'Content-Type':'application/json','x-staged-upload':'1'},body:JSON.stringify({uploadPath:path,filename:'test-master.docx',metadata:JSON.stringify({name:'Test site',timeZone:'Asia/Dubai'})})});assert.equal(result.status,201);const site=await result.json();assert.equal(values.has(path),false);
 const photo=new FormData();photo.append('photo',new Blob([await readFile(new URL('./fixtures/test-photo.jpg',import.meta.url))],{type:'image/jpeg'}),'test.jpg');photo.append('metadata',JSON.stringify({width:320,height:200,capturedAt:'2026-11-05T12:00:00Z',relevant:true,caption:'Pruning',activity:'Pruning'}));assert.equal((await call(`/api/locations/${site.id}/photos`,{method:'POST',body:photo})).status,201);
 const generated=await call('/api/reports',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({projectId:site.id,month:'2026-11',reviewed:true,summary:'Pruning completed.'})});assert.equal(generated.status,201);const report=await generated.json();assert.equal(report.photoCount,1);
 const download=await call(`/api/reports/${report.id}/download`);assert.equal(download.status,200);assert.ok(download.body instanceof ReadableStream);assert.ok(unzipSync(new Uint8Array(await download.arrayBuffer()))['word/document.xml']);
 const state=await(await call('/api/state')).json();assert.equal(state.locations.length,1);assert.equal(state.reports.length,1);
});

test('CompanyCam import requires review before Word generation and excludes a previous project link',async()=>{
 const values=new Map(),storage={put:async(k,v)=>{values.set(k,typeof v==='string'?new TextEncoder().encode(v):new Uint8Array(v));},get:async k=>values.has(k)?{statusCode:200,stream:new Response(values.get(k)).body,blob:{size:values.get(k).length}}:null,list:async({prefix})=>({blobs:[...values.keys()].filter(k=>k.startsWith(prefix)).map(pathname=>({pathname})),hasMore:false}),del:async()=>{}};
 const jpg=await readFile(new URL('./fixtures/test-photo.jpg',import.meta.url));
 const env={BLOB_READ_WRITE_TOKEN:'synthetic',COMPANYCAM_API_KEY:'synthetic'};
 const upstream=async url=>{const u=new URL(url);if(u.hostname==='companycam.imgix.net')return new Response(jpg);return Response.json({data:u.pathname.includes('/photos/')?{id:'10',project_id:'42',captured_at:'2026-11-05T00:00:00Z',description:{plain_text_content:'Pruning'},uris:[{type:'web',url:'https://companycam.imgix.net/photo.jpg'}]}:{id:u.pathname.endsWith('/43')?'43':'42',name:'CompanyCam site'},errors:[],meta:{has_next:false,next_cursor:null}});};
 const handler=createHandler({worker,storage,env,fetchImpl:upstream});
 const call=(p,opts={})=>handler(new Request('https://app.test/api'+p,{...opts,headers:{...opts.headers}})),post=(p,data)=>call(p,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)});
 const f=new FormData();f.append('template',new File([demoTemplate()],'master.docx'));f.append('metadata',JSON.stringify({name:'Site',timeZone:'Asia/Dubai'}));const created=await call('/locations',{method:'POST',body:f});assert.equal(created.status,201);const l=await created.json();
 assert.equal((await post('/companycam/link',{locationId:l.id,companyCamProjectId:'42'})).status,200);
 const imported=await post('/companycam/import',{locationId:l.id,month:'2026-11',photoId:'10'});assert.equal(imported.status,201);const p=(await imported.json()).photo;
 const generate=()=>post('/reports',{projectId:l.id,month:'2026-11',reviewed:true,summary:'Pruning completed.'});assert.equal((await generate()).status,400);
 const updateCaption = data => call(`/photos/${l.id}/${p.id}`, {
  method:'PATCH', headers:{'Content-Type':'application/json'}, body:JSON.stringify(data)
 });
 assert.equal((await updateCaption({caption:'An unlisted caption'})).status,400);
 const suggested=await updateCaption({caption:'Lawn Mowing',captionSuggestion:true});
 assert.equal((await suggested.json()).relevant,false);
 await updateCaption({caption:'LAWN MOWING'});
 const again=await updateCaption({caption:'Pruning',captionSuggestion:true});
 assert.equal((await again.json()).caption,'LAWN MOWING');
 const review=await call(`/photos/${l.id}/${p.id}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({relevant:true,caption:'Tree Pruning',activity:'Pruning'})});assert.equal(review.status,200);assert.equal((await review.json()).reviewRequired,false);
 const generated=await generate();assert.equal(generated.status,201);const report=await generated.json();assert.equal(report.photoCount,1);const download=await call(`/reports/${report.id}/download`);assert.equal(download.status,200);assert.ok(unzipSync(new Uint8Array(await download.arrayBuffer()))['word/document.xml']);
 assert.equal((await post('/companycam/link',{locationId:l.id,companyCamProjectId:'43'})).status,200);assert.deepEqual(await(await call(`/locations/${l.id}/photos`)).json(),[]);assert.equal((await generate()).status,400);assert.ok(values.has(`field-report/photos/${l.id}/${p.id}.json`));
});
