import http from 'node:http';
import {readFile,writeFile,mkdir,readdir,rm} from 'node:fs/promises';import {resolve,dirname,join} from 'node:path';
import worker from '../dist/server/index.js';
const root=resolve('.local-data');await mkdir(root,{recursive:true});
class DiskBucket{
 path(key){const p=resolve(root,key);if(!p.startsWith(root+'\\')&&!p.startsWith(root+'/'))throw Error('Invalid storage key');return p;}
 async put(key,value){const p=this.path(key);await mkdir(dirname(p),{recursive:true});await writeFile(p,typeof value==='string'?value:new Uint8Array(value));}
 async get(key){try{const bytes=await readFile(this.path(key));return {body:bytes,arrayBuffer:async()=>bytes.buffer.slice(bytes.byteOffset,bytes.byteOffset+bytes.byteLength),json:async()=>JSON.parse(bytes.toString())};}catch(e){if(e.code==='ENOENT')return null;throw e;}}
 async delete(keys){for(const key of Array.isArray(keys)?keys:[keys])await rm(this.path(key),{force:true});}
 async list({prefix='',cursor,limit=1000}){const files=await readdir(root,{recursive:true,withFileTypes:true});const keys=files.filter(d=>d.isFile()).map(d=>join(d.parentPath,d.name).slice(root.length+1).replaceAll('\\','/')).filter(k=>k.startsWith(prefix)).sort();const start=cursor?Number(cursor):0;return {objects:keys.slice(start,start+limit).map(key=>({key})),truncated:start+limit<keys.length,cursor:String(start+limit)};}
}
const bucket=new DiskBucket();http.createServer(async(req,res)=>{try{const chunks=[];let size=0;for await(const chunk of req){size+=chunk.length;if(size>15*1024*1024){res.writeHead(413);res.end('Request too large');return;}chunks.push(chunk);}const r=await worker.fetch(new Request(`http://127.0.0.1:4173${req.url}`,{method:req.method,headers:req.headers,body:['GET','HEAD'].includes(req.method)?undefined:Buffer.concat(chunks)}),{BUCKET:bucket});res.writeHead(r.status,Object.fromEntries(r.headers));res.end(Buffer.from(await r.arrayBuffer()));}catch(e){res.writeHead(500);res.end(e.message);}}).listen(4173,'127.0.0.1',()=>console.log('Local: http://127.0.0.1:4173'));
