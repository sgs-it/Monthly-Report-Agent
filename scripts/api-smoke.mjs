import {readFile,readdir,writeFile,mkdir} from 'node:fs/promises';import {join} from 'node:path';
import {unzipSync,strFromU8} from 'fflate';
const origin=process.env.REPORT_APP_URL||'http://127.0.0.1:4173';
async function call(path,options={}){const r=await fetch(origin+'/api'+path,options);if(!r.ok)throw Error(await r.text());return r;}
const names=await readdir('Master report templates'),word=names.find(n=>n.includes('Al Bateen')&&n.endsWith('.docx')),excel=names.find(n=>n.endsWith('.XLSX'));
for(const name of [word,excel]){
 const form=new FormData();form.append('template',new Blob([await readFile(join('Master report templates',name))]),name);form.append('metadata',JSON.stringify({name:'QA '+(name===word?'Word site':'Excel site'),timeZone:'Asia/Dubai',periodText:'October 2026'}));const l=await (await call('/locations',{method:'POST',body:form})).json();
 const photoForm=new FormData(),image=await readFile(new URL('../tests/fixtures/test-photo.jpg',import.meta.url));photoForm.append('photo',new Blob([image],{type:'image/jpeg'}),'test-photo.jpg');photoForm.append('metadata',JSON.stringify({name:'QA photo',capturedAt:'2026-11-10T12:00:00Z',width:320,height:200,caption:'Test work activity',activity:'Pruning',relevant:true}));await call(`/locations/${l.id}/photos`,{method:'POST',body:photoForm});
 const report=await(await call('/reports',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({projectId:l.id,month:'2026-11',reviewed:true,summary:'Test work activity.'})})).json();const output=new Uint8Array(await(await call(`/reports/${report.id}/download`)).arrayBuffer());
 await mkdir('qa',{recursive:true});await writeFile(`qa/${name===word?'word':'excel'}-report.docx`,output);const zip=unzipSync(output);if(!strFromU8(zip['word/document.xml']).includes('Test work activity'))throw Error('Word photo caption missing');
 if(name===excel){const bytes=new Uint8Array(await(await call(`/reports/${report.id}/excel`)).arrayBuffer());const z=unzipSync(bytes);if(!z['xl/workbook.xml'])throw Error('Invalid Excel output');await writeFile('qa/excel-report.xlsx',bytes);}
 console.log(`Passed ${l.format} API upload → select → generate → download (${report.photoCount} photo)`);
}
