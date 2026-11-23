import {readFile,readdir,writeFile,mkdir} from 'node:fs/promises';import {join} from 'node:path';import {generateReport,inspectTemplate} from '../src/docx.js';import {unzipSync,strFromU8} from 'fflate';
await mkdir('qa',{recursive:true});const results=[];const image=new Uint8Array([255,216,255,217]);
for(const name of await readdir('Master report templates'))if(/\.docx$/i.test(name))try{
 const source=new Uint8Array(await readFile(join('Master report templates',name))),i=inspectTemplate(source),photoTables=i.tables.filter(t=>t.images>=2).map(t=>t.index);
 if(!photoTables.length){results.push({name,status:'needs-mapping'});continue;}
 const period=(i.paragraphs.map(p=>p.text).join('\n').match(/(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d{2}/i)||[])[0]||'';
 const output=generateReport(source,{photoTableIndices:photoTables,periodText:period},{REPORT_PERIOD:'November 2026',REPORT_MONTH:'November',REPORT_YEAR:'2026',LOCATION:'Test site',WORK_SUMMARY:'Test activity'},Array.from({length:5},()=>({bytes:image,width:100,height:100,caption:'QA activity photo'})));
 const xml=strFromU8(unzipSync(output)['word/document.xml']);if((xml.match(/QA activity photo<\/w:t>/g)||[]).length!==5)throw Error('Caption count mismatch');
 results.push({name,status:'passed',outputBytes:output.length});
}catch(e){results.push({name,status:'failed',error:e.message});}
await writeFile('qa/smoke-results.json',JSON.stringify(results,null,2));console.log(JSON.stringify({total:results.length,passed:results.filter(r=>r.status==='passed').length,other:results.filter(r=>r.status!=='passed')},null,2));if(results.some(r=>r.status==='failed'))process.exitCode=1;
