import test from 'node:test';import assert from 'node:assert/strict';
import {createHandler,routeRequest} from '../server/vercel-handler.js';import {BlobBucket} from '../server/blob-bucket.js';
const env={BLOB_READ_WRITE_TOKEN:'synthetic-storage-token'};
test('the ordinary address and workspace APIs open with no credentials, even with old login settings',async()=>{
 const handler=createHandler({env:{...env,APP_PASSWORD:'old-password-setting',TEAM_LINK_KEY:'old-team-setting'},storage:{},worker:{fetch:()=>new Response('workspace')}});
 for(const path of ['/','/api/state']){
  const res=await handler(new Request('https://test'+path));
  assert.equal(res.status,200);
  assert.equal(res.headers.has('www-authenticate'),false);
  assert.equal(res.headers.has('set-cookie'),false);
  assert.equal(await res.text(),'workspace');
 }
});
test('reconstructs rewritten API paths and preserves payload and query',async()=>{const r=routeRequest(new Request('https://test/api/index?__route=api/reports&filter=current',{method:'POST',body:'body'}));assert.equal(new URL(r.url).pathname,'/api/reports');assert.equal(new URL(r.url).search,'?filter=current');assert.equal(await r.text(),'body');});
test('maps the Vercel function entry URL to the application home page',()=>{for(const p of ['/api/index','/api/index.js','/api','/api/'])assert.equal(new URL(routeRequest(new Request('https://test'+p)).url).pathname,'/');});
test('rejects cross-site changes before reaching storage',async()=>{
 const handler=createHandler({env,storage:{},worker:{}});
 const res=await handler(new Request('https://test/api/reports',{method:'POST',headers:{Origin:'https://other.test'}}));
 assert.equal(res.status,403);
});
test('upload tokens are restricted to bounded, incoming template files',async()=>{let settings;const handleUpload=async o=>{settings=await o.onBeforeGenerateToken('incoming/12345678-abcd.docx');return {ok:true};};const h=createHandler({env,storage:{},worker:{},handleUpload});const r=await h(new Request('https://test/api/uploads',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({type:'blob.generate-client-token'})}));assert.equal(r.status,200);assert.equal(settings.maximumSizeInBytes,12*1024*1024);assert.equal(settings.allowOverwrite,false);});
test('private storage adapter preserves keys, bypasses stale cache, streams downloads and paginates',async()=>{const calls=[];const sdk={put:async(...args)=>{calls.push(args);},get:async(k,o)=>{assert.equal(o.access,'private');assert.equal(o.useCache,false);return {statusCode:200,stream:new Response('{"name":"Site"}').body,blob:{size:15}};},list:async()=>({blobs:[{pathname:'field-report/locations/site.json'}],hasMore:false,cursor:undefined}),del:async()=>{}};const b=new BlobBucket(sdk,'test');await b.put('locations/site.json','{}');assert.equal(calls[0][0],'field-report/locations/site.json');assert.equal(calls[0][2].access,'private');assert.deepEqual(await(await b.get('locations/site.json')).json(),{name:'Site'});assert.deepEqual((await b.list({prefix:'locations/'})).objects,[{key:'locations/site.json'}]);assert.throws(()=>b.key('../secret'));});
test('large staged masters bypass the function request body but still use the master parser',async()=>{let size=0;const sdk={get:async()=>({statusCode:200,stream:new Response(new Uint8Array(6*1024*1024)).body,blob:{size:6*1024*1024,contentType:'application/zip'}})};const h=createHandler({env,storage:sdk,worker:{fetch:async r=>{size=(await r.arrayBuffer()).byteLength;return Response.json({ok:true});}}});const r=await h(new Request('https://test/api/templates/inspect',{method:'POST',headers:{'Content-Type':'application/json','x-staged-upload':'1'},body:JSON.stringify({uploadPath:'incoming/12345678-abcd.docx'})}));assert.equal(r.status,200);assert.equal(size,6*1024*1024);});
