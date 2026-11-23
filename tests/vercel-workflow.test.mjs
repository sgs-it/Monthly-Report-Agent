import test from 'node:test';import assert from 'node:assert/strict';import {readFile} from 'node:fs/promises';
import {zipSync,unzipSync} from 'fflate';import {demoTemplate} from '../src/docx.js';import worker from '../dist/server/index.js';import {createHandler} from '../server/vercel-handler.js';
test('Vercel adapter completes large master upload, photo selection, persistence and streamed Word download',async()=>{
 const values=new Map();const storage={
  put:async(path,value,options)=>{assert.equal(options.access,'private');values.set(path,typeof value==='string'?new TextEncoder().encode(value):new Uint8Array(value));return {pathname:path};},
  get:async(path,options)=>{assert.equal(options.access,'private');const bytes=values.get(path);return bytes?{statusCode:200,stream:new Response(bytes).body,blob:{size:bytes.length,contentType:'application/octet-stream'}}:null;},
  list:async({prefix})=>({blobs:[...values.keys()].filter(k=>k.startsWith(prefix)).map(pathname=>({pathname})),hasMore:false}),
  del:async(paths)=>{for(const p of Array.isArray(paths)?paths:[paths])values.delete(p);}
 };
 const env={APP_PASSWORD:'synthetic-test-password-123',BLOB_READ_WRITE_TOKEN:'synthetic'};const headers={Authorization:'Basic '+Buffer.from('team:'+env.APP_PASSWORD).toString('base64')};const handler=createHandler({worker,storage,env});
 const call=(path,opts={})=>handler(new Request('https://app.test'+path,{...opts,headers:{...headers,...opts.headers}}));
 const z=unzipSync(demoTemplate());z['word/media/unused.png']=new Uint8Array(5*1024*1024);const bytes=zipSync(z,{level:0});assert.ok(bytes.length>4.5*1024*1024);
 const path='incoming/12345678-abcd.docx';values.set(path,bytes);
 const result=await call('/api/locations',{method:'POST',headers:{'Content-Type':'application/json','x-staged-upload':'1'},body:JSON.stringify({uploadPath:path,filename:'test-master.docx',metadata:JSON.stringify({name:'Test site',timeZone:'Asia/Dubai'})})});assert.equal(result.status,201);const site=await result.json();assert.equal(values.has(path),false);
 const photo=new FormData();photo.append('photo',new Blob([await readFile(new URL('./fixtures/test-photo.jpg',import.meta.url))],{type:'image/jpeg'}),'test.jpg');photo.append('metadata',JSON.stringify({width:320,height:200,capturedAt:'2026-11-05T12:00:00Z',relevant:true,caption:'Test work activity',activity:'Pruning'}));assert.equal((await call(`/api/locations/${site.id}/photos`,{method:'POST',body:photo})).status,201);
 const generated=await call('/api/reports',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({projectId:site.id,month:'2026-11',reviewed:true,summary:'Pruning completed.'})});assert.equal(generated.status,201);const report=await generated.json();assert.equal(report.photoCount,1);
 const download=await call(`/api/reports/${report.id}/download`);assert.equal(download.status,200);assert.ok(download.body instanceof ReadableStream);assert.ok(unzipSync(new Uint8Array(await download.arrayBuffer()))['word/document.xml']);
 const state=await(await call('/api/state')).json();assert.equal(state.locations.length,1);assert.equal(state.reports.length,1);
});
