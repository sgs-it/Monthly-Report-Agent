import {ZipWriter} from '@zip.js/zip.js/index-native.js';
import {projectLabel} from '../src/project-name.js';
import {currentMonth,periodLabel} from '../src/policy.js';

// Word and Excel files are already compressed. Stream them into the archive one at a time.
export async function writeReportZip(entries,write){
  const output=new WritableStream({write});
  const zip=new ZipWriter(output,{level:0,zip64:true,useWebWorkers:false,bufferedWrite:false});
  for(const {path,url} of entries){
    const response=await fetch(url);
    if(!response.ok||!response.body)throw Error('Could not download '+path+'. Retry the export.');
    await zip.add(path,response.body);
  }
  await zip.close();
}

export function prepareBatch({api,getSites,syncPhotos,onReview,onReport}){
  const dialog=document.querySelector('#batch-dialog'),month=dialog.querySelector('[type=month]');
  const sync=dialog.querySelector('[name=sync]'),review=dialog.querySelector('[name=review]'),form=dialog.querySelector('form');
  const status=dialog.querySelector('[role=status]'),issues=dialog.querySelector('ul');
  const start=dialog.querySelector('[type=submit]'),close=dialog.querySelector('[data-close]');
  const runs=new Map();
  let working=false;
  const linked=()=>getSites().filter(site=>site.companyCamProjectId);
  function describe(){status.textContent=linked().length+' linked projects · Uses the reviewed work photos saved for the selected month.';issues.replaceChildren();review.checked=false;}
  month.onchange=describe;
  dialog.oncancel=e=>{if(working)e.preventDefault();};
  form.onsubmit=async e=>{
    e.preventDefault();if(working)return;
    const sites=linked(),period=month.value;
    if(!sites.length){status.textContent='Link a project in Integrations first.';return;}
    working=true;start.disabled=close.disabled=sync.disabled=review.disabled=month.disabled=true;
    const filename='Linked-project-reports-'+period+'.zip',entries=[];
    let writer;
    try{
      const destination=window.showSaveFilePicker?await window.showSaveFilePicker({suggestedName:filename,types:[{description:'ZIP archive',accept:{'application/zip':['.zip']}}]}):null;
      issues.replaceChildren();
      if(!runs.has(period))runs.set(period,{id:crypto.randomUUID(),synced:new Set()});
      const run=runs.get(period);
      for(const [i,site] of sites.entries()){
        const label=projectLabel(site);
        status.textContent='Generating '+(i+1)+' of '+sites.length+' · '+label;
        try{
          const syncKey=site.id+'|'+site.companyCamProjectId;
          if(sync.checked&&!run.synced.has(syncKey)){
            const importing='Syncing '+(i+1)+' of '+sites.length+' · '+label;
            status.textContent=importing;await syncPhotos(site,period,n=>{status.textContent=importing+' · '+n+' photos checked';});
            run.synced.add(syncKey);status.textContent='Generating · '+label;
          }
          const report=await api('/reports',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({projectId:site.id,month:period,batchId:run.id,linkedProjectId:site.companyCamProjectId,reviewed:true})});
          onReport(report);
          const stem=String(i+1).padStart(3,'0')+' - '+label.replace(/[<>:"/\\|?*\u0000-\u001f]/g,'_').slice(0,100)+' - '+period;
          entries.push({path:stem+'.docx',url:'/api/reports/'+report.id+'/download'});
          if(report.excelFilename)entries.push({path:stem+'.xlsx',url:'/api/reports/'+report.id+'/excel'});
        }catch(error){
          const item=document.createElement('li'),reviewSite=document.createElement('button');
          item.textContent=label+' — '+error.message+' ';
          reviewSite.type='button';reviewSite.className='text-button';reviewSite.textContent='Review project';
          reviewSite.onclick=()=>{if(!working){dialog.close();onReview(site.id,period);}};
          item.append(reviewSite);issues.append(item);
        }
      }
      if(issues.children.length){status.textContent=(sites.length-issues.children.length)+' of '+sites.length+' reports ready. Resolve the sites listed below and retry to download the complete ZIP.';return;}
      status.textContent='Packing '+entries.length+' files for '+periodLabel(period)+'…';
      const parts=[];
      if(destination)writer=await destination.createWritable();
      await writeReportZip(entries,data=>writer?writer.write(data):parts.push(data));
      if(writer){await writer.close();writer=null;}
      else{
        const url=URL.createObjectURL(new Blob(parts,{type:'application/zip'})),a=document.createElement('a');
        a.href=url;a.download=filename;a.click();setTimeout(()=>URL.revokeObjectURL(url),60000);
      }
      status.textContent='Exported '+sites.length+' linked projects for '+periodLabel(period)+'.';
      runs.delete(period);
    }catch(error){
      if(writer)await writer.abort().catch(()=>{});
      status.textContent=error.name==='AbortError'?'Export cancelled.':error.message;
    }finally{working=false;start.disabled=close.disabled=sync.disabled=review.disabled=month.disabled=false;}
  };
  return selectedMonth=>{month.value=selectedMonth||currentMonth();describe();dialog.showModal();};
}
