import {readFile,readdir} from 'node:fs/promises';import {join} from 'node:path';import {inspectTemplate} from '../src/docx.js';
import {createInterface} from 'node:readline';
if(process.stdin.isTTY)process.stdin.setRawMode(true);
const rl=createInterface({input:process.stdin});const input=await new Promise(resolve=>rl.once('line',resolve));rl.close();if(process.stdin.isTTY)process.stdin.setRawMode(false);const config=JSON.parse(input||'{}'),url=config.url||'http://127.0.0.1:4173';const headers=config.token?{'OAI-Sites-Authorization':'Bearer '+config.token}:{};
const state=await(await fetch(url+'/api/state',{headers})).json();if(!state.locations)throw Error(state.error||'Could not read template library');
const files=(await readdir('Master report templates')).filter(n=>/\.(docx|xlsx)$/i.test(n));let added=0,skipped=0,failed=[];
for(const [index,name]of files.entries()){
 if(state.locations.some(l=>l.templateName===name)){skipped++;continue;}
 try{const bytes=await readFile(join('Master report templates',name)),isExcel=/\.xlsx$/i.test(name);let project=name.replace(/\.(docx|xlsx)$/i,'').replace(/^(?:Indoor )?(?:Landscaping|Indoor Plants|Landscape) Monthly Report\s*-\s*/i,'').replace(/[- ]+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d{2}$/i,'').trim(),periodText='';
 if(!isExcel){const i=inspectTemplate(new Uint8Array(bytes)),ps=i.paragraphs,idx=ps.findIndex(p=>/^location\s*:?$/i.test(p.text.trim()));if(idx>=0&&ps[idx+1]?.text)project=ps[idx+1].text;periodText=(ps.map(p=>p.text).join('\n').match(/(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d{2}/i)||[])[0]||'';}
 const form=new FormData();form.append('template',new Blob([bytes]),name);form.append('metadata',JSON.stringify({name:project,timeZone:'Asia/Dubai',periodText}));const r=await fetch(url+'/api/locations',{method:'POST',headers,body:form});if(!r.ok)throw Error((await r.json()).error);added++;if(added%10===0)console.log(`Imported ${added} templates (${index+1}/${files.length})`);
 }catch(e){failed.push({name,error:e.message});}
}
console.log(JSON.stringify({total:files.length,added,skipped,failed}));if(failed.length)process.exitCode=1;
