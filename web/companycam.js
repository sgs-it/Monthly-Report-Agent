import {periodLabel} from '../src/policy.js';
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function setupCompanyCam({api,getContext,reload,toast,onError}){
 let status={connected:false,message:'Checking connection…'},projects=[],next=null,importing=false,stop=false;
 const dialog=document.createElement('dialog');dialog.innerHTML='<form><div class="dialog-heading"><h2>Link CompanyCam project</h2><button type="button" id="cc-close" class="secondary">Close</button></div><p>Choose the exact CompanyCam project for this site. The master template stays attached to this report project.</p><label>CompanyCam project<select id="cc-project" required></select></label><p id="cc-project-status" role="status"></p><div class="dialog-actions"><button type="button" id="cc-more" class="secondary">Load more projects</button><button id="cc-link" class="primary">Link project</button></div></form>';document.body.append(dialog);
 const $=s=>dialog.querySelector(s);
 $('#cc-close').onclick=()=>dialog.close();
 function options(){const previous=$('#cc-project').value;$('#cc-project').innerHTML=projects.map(p=>`<option value="${esc(p.id)}">${esc(p.name)} (${esc(p.id)})</option>`).join('');if(projects.some(p=>p.id===previous))$('#cc-project').value=previous;$('#cc-more').hidden=!next;$('#cc-project-status').textContent=projects.length+' projects loaded'+(next?' · Load more to see the next page.':'');$('#cc-link').disabled=!projects.length;}
 async function projectPage(after){const data=await api('/companycam/projects'+(after?'?after='+encodeURIComponent(after):''));if(data.next&&data.next===after)throw Error('CompanyCam pagination did not advance. Retry.');projects.push(...data.projects);next=data.next;options();}
 $('#cc-more').onclick=async()=>{const b=$('#cc-more');b.disabled=true;try{await projectPage(next);}catch(e){$('#cc-project-status').textContent=e.message;}finally{b.disabled=false;}};
 let linkingLocationId='';
 async function openLink(){const l=getContext().location;if(!l)return;linkingLocationId=l.id;projects=[];next=null;options();$('#cc-project-status').textContent='Loading CompanyCam projects…';dialog.showModal();try{await projectPage();if(l.companyCamProjectId&&projects.some(p=>p.id===l.companyCamProjectId))$('#cc-project').value=l.companyCamProjectId;}catch(e){$('#cc-project-status').textContent=e.message;}}
 dialog.querySelector('form').onsubmit=async e=>{e.preventDefault();$('#cc-link').disabled=true;try{await api('/companycam/link',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({locationId:linkingLocationId,companyCamProjectId:$('#cc-project').value})});dialog.close();await reload();toast('CompanyCam project linked.');}catch(e){$('#cc-project-status').textContent=e.message;}finally{$('#cc-link').disabled=false;}};
 async function check(){status={connected:false,message:'Checking connection…'};render();try{status=await api('/companycam/status');}catch(e){status={connected:false,message:e.message};}render();}
 async function importPhotos(){
  const {location:l,month}=getContext();if(!l?.companyCamProjectId||importing)return;
  importing=true;stop=false;let imported=0,existing=0,excluded=0,after,complete=false;const seen=new Set();render();
  try{
   do{
    const params=new URLSearchParams({locationId:l.id,month});if(after)params.set('after',after);
    const page=await api('/companycam/photos?'+params);
    excluded+=page.excluded;
    for(const photo of page.photos){if(stop)break;const result=await api('/companycam/import',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({locationId:l.id,month,photoId:photo.id})});if(result.alreadyImported)existing++;else imported++;const progress=document.querySelector('#cc-import-progress');if(progress)progress.textContent=`${imported} imported · ${existing} already present · ${excluded} outside month or missing capture date`;}
    if(stop)break;after=page.next;if(after){if(seen.has(after))throw Error('CompanyCam repeated a page. Import stopped; retry to resume.');seen.add(after);}
   }while(after);
   complete=!stop;toast(`${complete?'Import complete':'Import stopped'}: ${imported} added, ${existing} already present. Review captions and mark work photos before generating.`);
  }catch(e){onError(Error(`Import incomplete: ${imported} added and saved. ${e.message} Retry to resume.`));}
  finally{importing=false;await reload();render();}
 }
 function render(){
  const badge=document.querySelector('#companycam-status'),note=document.querySelector('#companycam-note'),pill=document.querySelector('#connection-mode');
  if(badge)badge.textContent=status.connected?'Connected':status.configured?'Connection needs attention':status.message==='Checking connection…'?'Checking…':'Not connected';
  if(note)note.textContent=status.message;
  if(pill)pill.textContent=status.connected?'CompanyCam + manual uploads':'Manual uploads';
  const overview=document.querySelector('#companycam-overview');if(overview){overview.innerHTML=`<p class="eyebrow">CONNECTION STATUS</p><h3>CompanyCam</h3><p>${esc(status.message)}</p><p>Link a CompanyCam project in the photo library, import the reporting month, then review the work photos.</p><button class="secondary" data-cc-check>Check connection</button>`;overview.querySelector('[data-cc-check]').onclick=check;}
  const controls=document.querySelector('#companycam-controls'),l=getContext().location;
  if(controls&&l){controls.innerHTML=`<div class="selection-summary"><strong>CompanyCam</strong> · ${esc(l.companyCamProjectName||l.companyCamProjectId||'No project linked')}<br><small>${esc(status.message)} Imported photos need a caption and work-activity review.</small><div class="dialog-actions"><button class="secondary" id="cc-check">Check connection</button><button class="secondary" id="cc-connect" ${!status.connected||importing?'disabled':''}>${l.companyCamProjectId?'Change linked project':'Link CompanyCam project'}</button><button class="primary" id="cc-import" ${!status.connected||!l.companyCamProjectId||importing?'disabled':''}>${importing?'Importing…':'Import '+esc(periodLabel(getContext().month))+' photos'}</button>${importing?'<button class="secondary" id="cc-stop">Stop import</button>':''}</div><p id="cc-import-progress" role="status"></p></div>`;controls.querySelector('#cc-check').onclick=check;controls.querySelector('#cc-connect').onclick=openLink;controls.querySelector('#cc-import').onclick=importPhotos;const cancel=controls.querySelector('#cc-stop');if(cancel)cancel.onclick=()=>{stop=true;cancel.disabled=true;cancel.textContent='Stopping…';};}
 }
 return {render,check};
}
