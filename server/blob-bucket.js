const PREFIX='field-report/';
export class BlobBucket {
 constructor(sdk,token){this.sdk=sdk;this.token=token;}
 key(key){if(!/^[a-zA-Z0-9_./-]+$/.test(key)||key.split('/').some(x=>x==='..'||x===''))throw Error('Invalid storage path.');return PREFIX+key;}
 async create(key,value){return this.sdk.put(this.key(key),value,{token:this.token,access:'private',addRandomSuffix:false,allowOverwrite:false,cacheControlMaxAge:60});}
 async put(key,value){const body=typeof value==='string'?value:Buffer.from(value);return this.sdk.put(this.key(key),body,{token:this.token,access:'private',addRandomSuffix:false,allowOverwrite:true,cacheControlMaxAge:60});}
 async get(key){const result=await this.sdk.get(this.key(key),{token:this.token,access:'private',useCache:false});if(!result||result.statusCode!==200)return null;return {etag:result.blob.etag,body:result.stream,arrayBuffer:()=>new Response(result.stream).arrayBuffer(),json:()=>new Response(result.stream).json(),size:result.blob.size};}
 async list({prefix='',cursor,limit=1000}){const result=await this.sdk.list({token:this.token,prefix:PREFIX+prefix,cursor,limit});return {objects:result.blobs.map(b=>({key:b.pathname.slice(PREFIX.length)})),truncated:result.hasMore,cursor:result.cursor};}
 async delete(keys,etag){if(etag)return this.sdk.del(this.key(keys),{token:this.token,ifMatch:etag});return this.sdk.del((Array.isArray(keys)?keys:[keys]).map(k=>this.key(k)),{token:this.token});}
}
