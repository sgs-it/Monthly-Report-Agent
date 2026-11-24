import {createProjectPicker} from './project-picker.js';
import {periodLabel} from '../src/policy.js';
import {allowedCaption} from '../src/captions.js';
import {suggestCaption} from './local-caption.js';
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
export function setupCompanyCam({api,getContext,reload,toast,onError}){
 let status={connected:false,message:'Checking connection…'},importing=false,stop=false;
 const pickProject=createProjectPicker({api,reload,toast,onError});
 const openLink=()=>pickProject(getContext().location);
 async function check(){status={connected:false,message:'Checking connection…'};render();try{status=await api('/companycam/status');}catch(e){status={connected:false,message:e.message};}render();}
 async function importPhotos(){
  const {location:l,month}=getContext();if(!l?.companyCamProjectId||importing)return;
  importing=true;stop=false;let imported=0,existing=0,excluded=0,after,complete=false;const seen=new Set();render();
  let deviceAvailable=true,suggestions=0;
  try{
   do{
    const params=new URLSearchParams({locationId:l.id,month});if(after)params.set('after',after);
    const page=await api('/companycam/photos?'+params);
    excluded+=page.excluded;
    for(const photo of page.photos){
     if(stop)break;
     const result=await api('/companycam/import',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({locationId:l.id,month,photoId:photo.id})});
     if(result.alreadyImported)existing++;else imported++;
     const saved=result.photo,progress=document.querySelector('#cc-import-progress');
     if(deviceAvailable&&!saved.captionChosen&&(!saved.caption||!allowedCaption(saved.caption))){
      if(progress)progress.textContent='Suggesting a caption on this device. First use downloads the image model.';
      let suggestion;
      try{suggestion=await suggestCaption(saved.id);}
      catch{deviceAvailable=false;}
      if(suggestion){
       await api(`/photos/${l.id}/${saved.id}`,{method:'PATCH',headers:{'Content-Type':'application/json'},body:JSON.stringify({caption:suggestion,captionSuggestion:true})});
       suggestions++;
      }
     }
     if(progress)progress.textContent=`${imported} imported · ${existing} already present · ${suggestions} captions suggested · ${excluded} outside month or missing capture date`;
    }
    if(stop)break;after=page.next;if(after){if(seen.has(after))throw Error('CompanyCam repeated a page. Import stopped; retry to resume.');seen.add(after);}
   }while(after);
   complete=!stop;toast(`${complete?'Import complete':'Import stopped'}: ${imported} added, ${existing} already present. ${deviceAvailable?'Review the captions before generating.':'On-device captioning was unavailable. Choose captions from the list or retry the import.'}`);
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
