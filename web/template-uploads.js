import {upload} from '@vercel/blob/client';
const staged=new WeakMap();
async function stage(file){
 if(file.size>12*1024*1024)throw Error('The master template must be smaller than 12 MB.');
 if(!staged.has(file)){
  const extension=file.name.split('.').at(-1).toLowerCase();if(!['docx','xlsx'].includes(extension))throw Error('Choose a Word or Excel master.');
  staged.set(file,upload(`incoming/${crypto.randomUUID()}.${extension}`,file,{access:'private',contentType:file.type||'application/octet-stream',handleUploadUrl:'/api/uploads',multipart:true}).catch(error=>{staged.delete(file);throw error;}));
 }
 return (await staged.get(file)).pathname;
}
export async function prepareTemplateRequest(path,options){
 const body=options.body;
 if(path==='/templates/inspect'&&body instanceof File){return {...options,headers:{...options.headers,'Content-Type':'application/json','x-staged-upload':'1'},body:JSON.stringify({uploadPath:await stage(body)})};}
 if(path==='/locations'&&body instanceof FormData&&body.get('template') instanceof File){const file=body.get('template');return {...options,headers:{...options.headers,'Content-Type':'application/json','x-staged-upload':'1'},body:JSON.stringify({uploadPath:await stage(file),filename:file.name,metadata:body.get('metadata')})};}
 return options;
}
