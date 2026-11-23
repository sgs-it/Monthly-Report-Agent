import test from 'node:test';
import assert from 'node:assert/strict';
import {readFile} from 'node:fs/promises';
import {createCompanyCamHandler,normalizePhoto} from '../server/companycam.js';
import {createHandler} from '../server/vercel-handler.js';
import {selectPhotos} from '../src/policy.js';
const location={id:'site',name:'Test site',timeZone:'Asia/Dubai',companyCamProjectId:'42'};
const photo={id:'10',project_id:'42',captured_at:'2026-11-05T12:00:00Z',description:{plain_text_content:'Shrub pruning'},uris:[{type:'web',url:'https://companycam.imgix.net/test.jpg'}]};
const response=(data,meta={has_next:false,next_cursor:null})=>Response.json({data,errors:[],meta});
function bucket(){const values=new Map([['locations/site.json',JSON.stringify(location)]]);return {values,get:async key=>values.has(key)?{json:async()=>JSON.parse(values.get(key))}:null,put:async(key,value)=>{values.set(key,value);}};}
const request=(path,data)=>new Request('https://app.test/api/companycam/'+path,data?{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(data)}:{});
test('connection status distinguishes missing, rejected and working keys without disclosing credentials',async()=>{
 let calls=0;const missing=createCompanyCamHandler({env:{},fetchImpl:async()=>{calls++;}});assert.equal((await(await missing(request('status'))).json()).configured,false);assert.equal(calls,0);
 const bad=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'synthetic-private-key'},fetchImpl:async()=>new Response('synthetic-private-key upstream detail',{status:401})});const status=await(await bad(request('status'))).json();assert.equal(status.connected,false);assert.ok(!JSON.stringify(status).includes('synthetic-private-key'));assert.match(status.message,/rejected/);
 const good=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'synthetic-private-key'},fetchImpl:async(url,options)=>{assert.equal(url.origin,'https://app.companycam.com');assert.equal(options.headers.Authorization,'Bearer synthetic-private-key');assert.equal(options.redirect,'error');return response([]);}});assert.equal((await(await good(request('status'))).json()).connected,true);
});
test('project cursor is forwarded as a query value and incomplete pagination fails visibly',async()=>{
 const h=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async url=>{assert.equal(url.searchParams.get('after'),'abc&limit=1');assert.equal(url.searchParams.get('limit'),'100');return response([{id:'42',name:'Site'}],{has_next:true,next_cursor:'next-page'});}});const p=await(await h(request('projects?after='+encodeURIComponent('abc&limit=1')))).json();assert.equal(p.next,'next-page');assert.equal(p.projects[0].id,'42');
 const incomplete=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async()=>response([],{has_next:true,next_cursor:null})});assert.equal((await incomplete(request('projects'))).status,502);
});
test('linking checks the remote project and changes only its association',async()=>{
 const b=bucket(),h=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async()=>response({id:'43',name:'Second site'})});const r=await h(request('link',{locationId:'site',companyCamProjectId:'43'}),b);assert.equal(r.status,200);const l=JSON.parse(b.values.get('locations/site.json'));assert.equal(l.companyCamProjectId,'43');assert.equal(l.name,location.name);assert.equal(l.timeZone,location.timeZone);
});
test('month filtering uses capture dates and the exact site timezone, with all pages exposed',async()=>{
 const b=bucket(),h=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async url=>{assert.equal(url.pathname,'/public_api/v1/projects/42/photos');assert.ok(url.searchParams.has('start_date'));return response([photo,{...photo,id:'11',captured_at:'2026-10-31T21:00:00Z'},{...photo,id:'12',captured_at:'2026-11-30T21:00:00Z'},{...photo,id:'13',captured_at:null,created_at:'2026-11-05T00:00:00Z'}],{has_next:true,next_cursor:'page2'});}});const p=await(await h(request('photos?locationId=site&month=2026-11'),b)).json();assert.deepEqual(p.photos.map(x=>x.id),['10','11']);assert.equal(p.excluded,2);assert.equal(p.next,'page2');
});
test('foreign-project list results stop the import rather than selecting unrelated work',async()=>{
 const h=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async()=>response([{...photo,project_id:'other'}])});assert.equal((await h(request('photos?locationId=site&month=2026-11'),bucket())).status,502);
});
test('import resizes real photos, stores provenance, requires review, and preserves decisions on retry',async()=>{
 const b=bucket(),jpg=await readFile(new URL('./fixtures/test-photo.jpg',import.meta.url));let downloads=0;
 const h=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async(url,options)=>{if(String(url).startsWith('https://app.companycam.com/'))return response(photo);downloads++;assert.equal(options.headers,undefined);assert.equal(options.redirect,'error');return new Response(jpg,{headers:{'Content-Type':'image/jpeg'}});}});
 const data={locationId:'site',month:'2026-11',photoId:'10'},r=await h(request('import',data),b);assert.equal(r.status,201);const imported=(await r.json()).photo;assert.equal(imported.source,'companycam');assert.equal(imported.companyCamProjectId,'42');assert.equal(imported.relevant,false);assert.equal(imported.reviewRequired,true);assert.equal(imported.width,320);assert.equal(imported.height,200);assert.equal(b.values.get(`images/${imported.id}.jpg`)[0],255);assert.equal(selectPhotos([imported],'site','2026-11','Asia/Dubai').photos.length,0);
 imported.relevant=true;imported.caption='Reviewed pruning work';b.values.set(`photos/site/${imported.id}.json`,JSON.stringify(imported));const retry=await(await h(request('import',data),b)).json();assert.equal(retry.alreadyImported,true);assert.equal(retry.photo.caption,'Reviewed pruning work');assert.equal(downloads,1);assert.equal(selectPhotos([retry.photo],'site','2026-11','Asia/Dubai').photos.length,1);
});
test('import rejects a different project, capture month, missing capture date, or unsafe image host before download',async()=>{
 for(const change of [{project_id:'99'},{captured_at:'2026-10-05T00:00:00Z'},{captured_at:null},{uris:[{type:'web',url:'https://127.0.0.1/secret'}]}]){let downloads=0;const h=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async url=>{if(String(url).startsWith('https://app.companycam.com/'))return response({...photo,...change});downloads++;throw Error('Unexpected download');}});const r=await h(request('import',{locationId:'site',month:'2026-11',photoId:'10'}),bucket());assert.ok(r.status>=400);assert.equal(downloads,0);}
});
test('rate limits and oversized source images surface errors without writing a partial photo record',async()=>{
 const rate=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async()=>new Response('',{status:429})});assert.equal((await rate(request('projects'))).status,429);
 const b=bucket(),large=createCompanyCamHandler({env:{COMPANYCAM_API_KEY:'test'},fetchImpl:async url=>String(url).startsWith('https://app.companycam.com/')?response(photo):new Response('oversize',{headers:{'Content-Length':String(26*1024*1024)}})});assert.equal((await large(request('import',{locationId:'site',month:'2026-11',photoId:'10'}),b)).status,413);assert.equal(b.values.size,1);
});
test('CompanyCam endpoints retain team authentication and same-origin mutation checks',async()=>{
 let calls=0;const env={APP_PASSWORD:'synthetic-test-password-123',COMPANYCAM_API_KEY:'test'},h=createHandler({worker:{},storage:{},env,fetchImpl:async()=>{calls++;return response([]);}});assert.equal((await h(request('status'))).status,401);assert.equal(calls,0);const auth={Authorization:'Basic '+Buffer.from('team:'+env.APP_PASSWORD).toString('base64')};assert.equal((await h(new Request('https://app.test/api/companycam/status',{headers:auth}))).status,200);assert.equal(calls,1);assert.equal((await h(new Request('https://app.test/api/companycam/link',{method:'POST',headers:{...auth,Origin:'https://evil.test'},body:'{}'}))).status,403);assert.equal(calls,1);
});
test('image normalization rejects invalid image content',async()=>{await assert.rejects(normalizePhoto(Buffer.from('not an image')));});
