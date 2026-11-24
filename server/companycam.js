import {createHash} from 'node:crypto';
import sharp from 'sharp';
import {allowedCaption} from '../src/captions.js';
import {capturedMonth,validMonth} from '../src/policy.js';

const BASE='https://app.companycam.com/public_api/v1';
const json=(data,status=200)=>Response.json(data,{status,headers:{'Cache-Control':'no-store'}});
class ConnectorError extends Error {constructor(message,status=400){super(message);this.status=status;}}
function id(value){if(typeof value!=='string'||!/^\w[\w-]{0,199}$/.test(value))throw new ConnectorError('Choose a valid project or photo.');return value;}
function cursor(value){if(value&&value.length>2000)throw new ConnectorError('Invalid pagination cursor.');return value||undefined;}
function month(value){if(!validMonth(value||''))throw new ConnectorError('Choose a valid reporting month.');return value;}
function page(payload){if(!Array.isArray(payload.data)||typeof payload.meta?.has_next!=='boolean'||(payload.meta.has_next&&!payload.meta.next_cursor))throw new ConnectorError('CompanyCam returned incomplete pagination. Retry the import.',502);return {items:payload.data,next:payload.meta.has_next?String(payload.meta.next_cursor):null};}
function capture(photo){const value=photo.captured_at;return typeof value==='string'&&Number.isFinite(Date.parse(value))?new Date(value).toISOString():null;}
function caption(photo){const value=photo.description?.plain_text_content?.trim();return allowedCaption(value)?value:'';}
function imageURL(photo){
 const values=[...(photo.uris||[]).filter(x=>/web|large|original/i.test(x.type||'')),...(photo.uris||[])];
 for(const entry of values){try{const u=new URL(entry.url||entry.uri);if(u.protocol==='https:'&&!u.username&&!u.password&&!u.port&&(u.hostname==='companycam.imgix.net'||u.hostname.endsWith('.companycam.com')))return u.href;}catch{}}
 throw new ConnectorError('This photo has no supported CompanyCam image URL. Import it manually.',422);
}
async function boundedBytes(response,max){
 if(Number(response.headers.get('content-length'))>max)throw new ConnectorError('CompanyCam photo exceeds the 25 MB import limit.',413);
 const reader=response.body?.getReader();if(!reader)throw new ConnectorError('CompanyCam returned an empty photo.',502);
 const chunks=[];let length=0;try{while(true){const {done,value}=await reader.read();if(done)break;length+=value.length;if(length>max)throw new ConnectorError('CompanyCam photo exceeds the 25 MB import limit.',413);chunks.push(value);}}catch(e){await reader.cancel().catch(()=>{});throw e;}return Buffer.concat(chunks,length);
}
export async function normalizePhoto(bytes){
 const {data,info}=await sharp(bytes,{limitInputPixels:40_000_000,animated:false}).rotate().resize({width:1600,height:1600,fit:'inside',withoutEnlargement:true}).jpeg({quality:82}).toBuffer({resolveWithObject:true});
 if(data.length>2_000_000)throw new ConnectorError('The resized photo is too large. Upload a smaller copy manually.',413);
 return {bytes:data,width:info.width,height:info.height};
}
export function createCompanyCamHandler({env,fetchImpl=fetch,normalize=normalizePhoto}){
 async function call(path,params={}){
  if(!env.COMPANYCAM_API_KEY)throw new ConnectorError('Add COMPANYCAM_API_KEY to Vercel Production, then redeploy.',503);
  const url=new URL(BASE+path);for(const [key,value]of Object.entries(params))if(value!==undefined)url.searchParams.set(key,value);
  let r;try{r=await fetchImpl(url,{headers:{Authorization:'Bearer '+env.COMPANYCAM_API_KEY,Accept:'application/json'},redirect:'error',signal:AbortSignal.timeout(25_000)});}catch{throw new ConnectorError('CompanyCam could not be reached. Please retry.',502);}
  if(!r.ok){const messages={401:'CompanyCam rejected the key. Check that it is an active Application Key, save it in Production, and redeploy.',403:'The CompanyCam key cannot read this resource. Check its project and photo permissions.',404:'The CompanyCam project or photo is no longer available.',429:'CompanyCam rate limit reached. Wait a moment, then retry.'};throw new ConnectorError(messages[r.status]||'CompanyCam request failed. Please retry.',r.status===429?429:r.status===401||r.status===403?502:r.status===404?404:502);}
  let data;try{data=await r.json();}catch{throw new ConnectorError('CompanyCam returned an unexpected response.',502);}if(data.errors?.length)throw new ConnectorError('CompanyCam could not complete the request.',502);return data;
 }
 async function location(bucket,locationId){if(!bucket)throw new ConnectorError('Configure the app’s private report storage first.',503);const record=await bucket.get(`locations/${id(locationId)}.json`);if(!record)throw new ConnectorError('Report project not found.',404);return record.json();}
 return async(request,bucket)=>{
  const url=new URL(request.url),path=url.pathname;
  try{
   if(path==='/api/companycam/status'&&request.method==='GET'){
    if(!env.COMPANYCAM_API_KEY)return json({configured:false,connected:false,message:'Add COMPANYCAM_API_KEY in Vercel and redeploy.'});
    try{const result=await call('/projects',{limit:1});if(!Array.isArray(result.data))throw new ConnectorError('CompanyCam returned an unexpected project response.',502);return json({configured:true,connected:true,message:'Connected · on-device photo captions'});}catch(e){return json({configured:true,connected:false,message:e.message});}
   }
   if(path==='/api/companycam/projects'&&request.method==='GET'){const p=page(await call('/projects',{limit:100,after:cursor(url.searchParams.get('after'))}));return json({projects:p.items.map(p=>({id:String(p.id),name:String(p.name||p.id)})),next:p.next});}
   if(path==='/api/companycam/link'&&request.method==='POST'){
    const d=await request.json(),l=await location(bucket,d.locationId),remoteId=id(d.companyCamProjectId),remote=(await call('/projects/'+remoteId)).data;
    if(String(remote?.id)!==remoteId)throw new ConnectorError('CompanyCam returned a different project.',502);
    l.companyCamProjectId=remoteId;l.companyCamProjectName=String(remote.name||remoteId);l.updatedAt=new Date().toISOString();await bucket.put(`locations/${l.id}.json`,JSON.stringify(l));return json(l);
   }
   if(path==='/api/companycam/photos'&&request.method==='GET'){
    const l=await location(bucket,url.searchParams.get('locationId')),m=month(url.searchParams.get('month')),remoteId=id(l.companyCamProjectId);
    // Widen API dates by a day at each end; exact eligibility uses the site's time zone below.
    const start=new Date(m+'-01T00:00:00Z'),end=new Date(start);start.setUTCDate(0);end.setUTCMonth(end.getUTCMonth()+1);end.setUTCDate(2);
    const p=page(await call(`/projects/${remoteId}/photos`,{limit:100,after:cursor(url.searchParams.get('after')),start_date:start.toISOString(),end_date:end.toISOString()}));
    if(p.items.some(photo=>String(photo.project_id)!==remoteId))throw new ConnectorError('CompanyCam returned photos from a different project. Import stopped.',502);
    const eligible=p.items.filter(photo=>capturedMonth(capture(photo),l.timeZone)===m);
    return json({photos:eligible.map(photo=>({id:String(photo.id),capturedAt:capture(photo),caption:caption(photo)})),scanned:p.items.length,excluded:p.items.length-eligible.length,next:p.next});
   }
   if(path==='/api/companycam/import'&&request.method==='POST'){
    const d=await request.json(),l=await location(bucket,d.locationId),m=month(d.month),remoteId=id(l.companyCamProjectId),photoId=id(d.photoId);
    const photo=(await call('/photos/'+photoId)).data;
    if(String(photo?.id)!==photoId||String(photo.project_id)!==remoteId)throw new ConnectorError('Photo does not belong to the linked CompanyCam project.');
    const capturedAt=capture(photo);if(capturedMonth(capturedAt,l.timeZone)!==m)throw new ConnectorError('Photo was not captured in the selected reporting month.');
    const localId='cc-'+createHash('sha256').update(l.id+'\0'+remoteId+'\0'+photoId).digest('hex').slice(0,40),key=`photos/${l.id}/${localId}.json`,existing=await bucket.get(key);
    if(existing)return json({photo:await existing.json(),alreadyImported:true});
    const image=await fetchImpl(imageURL(photo),{redirect:'error',signal:AbortSignal.timeout(30_000)});
    if(!image.ok)throw new ConnectorError('CompanyCam image download failed. Retry this photo.',502);
    const normalized=await normalize(await boundedBytes(image,25*1024*1024));
    const record={id:localId,projectId:l.id,name:'CompanyCam '+photoId,caption:caption(photo),activity:'',capturedAt,relevant:false,width:normalized.width,height:normalized.height,hash:createHash('sha256').update(normalized.bytes).digest('hex'),source:'companycam',companyCamPhotoId:photoId,companyCamProjectId:remoteId,importedAt:new Date().toISOString(),reviewRequired:true};
    await bucket.put(`images/${localId}.jpg`,normalized.bytes);await bucket.put(key,JSON.stringify(record));return json({photo:record,alreadyImported:false},201);
   }
   return json({error:'CompanyCam endpoint or method not found.'},404);
  }catch(e){if(e instanceof ConnectorError)return json({error:e.message},e.status);console.error('CompanyCam operation failed:',e.name);return json({error:'CompanyCam import could not finish. Retry to resume; imported photos are kept.'},502);}
 };
}
