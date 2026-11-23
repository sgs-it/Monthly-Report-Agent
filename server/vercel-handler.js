import {createHash,timingSafeEqual} from 'node:crypto';
import {BlobBucket} from './blob-bucket.js';
const json=(value,status=200)=>Response.json(value,{status,headers:{'Cache-Control':'no-store'}});
const equal=(a,b)=>timingSafeEqual(createHash('sha256').update(a).digest(),createHash('sha256').update(b).digest());
export function authenticated(request,env){
 if(!env.APP_PASSWORD||env.APP_PASSWORD.length<16)return false;
 const header=request.headers.get('authorization')||'';if(!header.startsWith('Basic '))return false;
 try{const decoded=Buffer.from(header.slice(6),'base64').toString('utf8'),colon=decoded.indexOf(':');return colon>=0&&equal(decoded.slice(0,colon),env.APP_USERNAME||'team')&&equal(decoded.slice(colon+1),env.APP_PASSWORD);}catch{return false;}
}
export function routeRequest(request){
 const url=new URL(request.url);if(url.searchParams.has('__route')){const route=url.searchParams.get('__route');if(route.startsWith('/')||route.includes('\\'))throw Error('Invalid route');url.pathname='/'+route;url.searchParams.delete('__route');return new Request(url,request);}if(['/api/index','/api/index.js','/api','/api/'].includes(url.pathname)){url.pathname='/';return new Request(url,request);}return request;
}
export function createHandler({worker,storage,handleUpload,env}){
 return async incoming=>{
  try{
   const request=routeRequest(incoming),url=new URL(request.url),path=url.pathname;
   let uploadBody;
   if(path==='/api/uploads'&&request.method==='POST')uploadBody=await request.json();
   // The SDK verifies the signature of completion callbacks; these do not mint upload tokens.
   if(uploadBody?.type==='blob.upload-completed'){
    try{return json(await handleUpload({body:uploadBody,request,token:env.BLOB_READ_WRITE_TOKEN,onBeforeGenerateToken:async()=>{throw Error('Not authorized');},onUploadCompleted:async()=>{}}));}catch{return json({error:'Invalid upload callback.'},400);}
   }
   if(!env.APP_PASSWORD||env.APP_PASSWORD.length<16)return new Response('Setup required: add APP_PASSWORD (at least 16 characters) in Vercel environment variables, then redeploy.',{status:503,headers:{'Cache-Control':'no-store'}});
   if(!authenticated(request,env))return new Response('Sign in with the team username and password.',{status:401,headers:{'WWW-Authenticate':'Basic realm="Field Report Agent", charset="UTF-8"','Cache-Control':'no-store'}});
   if(request.method!=='GET'&&request.method!=='HEAD'&&request.headers.get('Origin')&&request.headers.get('Origin')!==url.origin)return json({error:'Cross-site request rejected.'},403);
   if(path==='/api/uploads'){
    if(!uploadBody)return json({error:'Use POST to upload.'},405);
    if(!env.BLOB_READ_WRITE_TOKEN)return json({error:'Connect a private Vercel Blob store with BLOB_READ_WRITE_TOKEN, then redeploy.'},503);
    const response=await handleUpload({request,body:uploadBody,token:env.BLOB_READ_WRITE_TOKEN,onBeforeGenerateToken:async pathname=>{
     if(!authenticated(request,env)||!/^incoming\/[a-f0-9-]+\.(docx|xlsx)$/i.test(pathname))throw Error('Invalid or unauthorized template upload.');
     return {allowedContentTypes:['application/vnd.openxmlformats-officedocument.wordprocessingml.document','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet','application/zip','application/octet-stream'],maximumSizeInBytes:12*1024*1024,validUntil:Date.now()+10*60*1000,addRandomSuffix:false,allowOverwrite:false,ttlDays:1};
    },onUploadCompleted:async()=>{}});
    return json(response);
   }
   let effective=request;
   if(request.headers.get('x-staged-upload')==='1'){
    if(!['/api/templates/inspect','/api/locations'].includes(path)||request.method!=='POST')return json({error:'Invalid staged upload operation.'},400);
    const data=await request.json();if(!/^incoming\/[a-f0-9-]+\.(docx|xlsx)$/i.test(data.uploadPath||''))return json({error:'Invalid staged template path.'},400);
    const result=await storage.get(data.uploadPath,{access:'private',token:env.BLOB_READ_WRITE_TOKEN,useCache:false});
    if(!result||result.statusCode!==200)return json({error:'Template upload is not available. Choose the file again.'},400);
    if(result.blob.size>12*1024*1024)return json({error:'Template exceeds 12 MB.'},413);
    const bytes=await new Response(result.stream).arrayBuffer(),headers=new Headers(request.headers);headers.delete('content-type');headers.delete('content-length');headers.delete('x-staged-upload');
    if(path==='/api/locations'){const f=new FormData();f.append('metadata',data.metadata||'{}');f.append('template',new File([bytes],String(data.filename||'master.docx'),{type:result.blob.contentType||'application/octet-stream'}));effective=new Request(url,{method:'POST',headers,body:f});}
    else effective=new Request(url,{method:'POST',headers,body:bytes});
    const response=await worker.fetch(effective,{BUCKET:new BlobBucket(storage,env.BLOB_READ_WRITE_TOKEN)});
    if(path==='/api/locations'&&response.ok)await storage.del(data.uploadPath,{token:env.BLOB_READ_WRITE_TOKEN});
    return response;
   }
   if(path.startsWith('/api/')&&!env.BLOB_READ_WRITE_TOKEN)return json({error:'Connect a private Vercel Blob store with BLOB_READ_WRITE_TOKEN, then redeploy.'},503);
   return await worker.fetch(effective,{BUCKET:new BlobBucket(storage,env.BLOB_READ_WRITE_TOKEN)});
  }catch(error){console.error('Request failed:',error.name);return json({error:'The request could not be completed. Check the private Blob connection and server logs.'},500);}
 };
}
