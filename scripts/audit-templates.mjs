import {readFile,readdir,writeFile,mkdir} from 'node:fs/promises';
import {join} from 'node:path';
import {inspectTemplate} from '../src/docx.js';
await mkdir('qa',{recursive:true});
const folder=process.argv[2]||'Master report templates';const result=[];
for(const file of await readdir(folder))if(file.endsWith('.docx')){
 try {const bytes=new Uint8Array(await readFile(join(folder,file)));const t=inspectTemplate(bytes);
 const periods=[...new Set(t.paragraphs.map(p=>p.text).join('\n').match(/(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d{2}/gi)||[])];
 result.push({file,size:bytes.length,paragraphs:t.paragraphs.length,tables:t.tables,periods,bodyImageCount:t.bodyImageCount,photoTables:t.tables.filter(x=>x.images>=2).map(x=>x.index)});
 }catch(e){result.push({file,error:e.message});}
}
await writeFile('qa/template-audit.json',JSON.stringify(result,null,2));
console.log(JSON.stringify({count:result.length,errors:result.filter(r=>r.error),largest:Math.max(...result.map(r=>r.size||0)),withoutPhotoTables:result.filter(r=>!r.error&&!r.photoTables.length).map(r=>r.file),periods:[...new Set(result.flatMap(r=>r.periods||[]))]},null,2));
