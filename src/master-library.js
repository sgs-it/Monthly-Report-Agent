const normalized=value=>String(value||'').normalize('NFKC').trim().toLowerCase().replace(/\s+/g,' ');

export function masterIdentity(site){
  const name=normalized(site.templateName),siteName=normalized(site.wordIdentity?.location||site.name);
  return name&&siteName?JSON.stringify([name,siteName,normalized(site.wordIdentity?.jobCode||site.code)]):'';
}

export function masterLibrary(sites){
  const groups=new Map(),locations=[],duplicates=[];
  for(const site of sites){
    const key=masterIdentity(site)||site.id;
    if(!groups.has(key))groups.set(key,[]);
    groups.get(key).push(site);
  }
  for(const group of groups.values()){
    group.sort((a,b)=>String(b.updatedAt||b.createdAt||'').localeCompare(String(a.updatedAt||a.createdAt||''))||a.id.localeCompare(b.id));
    const linked=group.filter(site=>site.companyCamProjectId);
    const active=linked.length?linked:[group[0]];
    locations.push(...active);
    // Two independently linked sites may share a master. Both links belong to the team.
    for(const site of group)if(!active.includes(site))duplicates.push({site,keep:active[0]});
  }
  return {locations:locations.sort((a,b)=>a.name.localeCompare(b.name)),duplicates};
}

export async function masterId(identity){
  const digest=await crypto.subtle.digest('SHA-256',new TextEncoder().encode(identity));
  return 'master-'+Array.from(new Uint8Array(digest),n=>n.toString(16).padStart(2,'0')).join('');
}
