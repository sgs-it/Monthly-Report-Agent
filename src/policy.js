export const DEFAULT_TIMEZONE = 'Asia/Calcutta';
export function currentMonth(timeZone = DEFAULT_TIMEZONE, now = new Date()) {
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-CA', {timeZone,year:'numeric',month:'2-digit'}).formatToParts(now).map(p=>[p.type,p.value]));
  return `${parts.year}-${parts.month}`;
}
export function validMonth(month) { return /^20\d{2}-(0[1-9]|1[0-2])$/.test(month); }
export function capturedMonth(value, timeZone = DEFAULT_TIMEZONE) {
  if (!value || !Number.isFinite(new Date(value).getTime())) return null;
  return currentMonth(timeZone, new Date(value));
}
export function periodLabel(month) {
  if (!validMonth(month)) throw new Error('Choose a valid report month.');
  return new Intl.DateTimeFormat('en', {month:'long',year:'numeric',timeZone:'UTC'}).format(new Date(`${month}-15T12:00:00Z`));
}
export function photoLimit(relevantCount) { return relevantCount > 200 ? 75 : 50; }
export function selectPhotos(photos, projectId, month, timeZone = DEFAULT_TIMEZONE) {
  if (!validMonth(month)) throw new Error('Choose a valid report month.');
  const hashes = new Set();
  const relevant = photos.filter(p => {
    if (p.projectId !== projectId || capturedMonth(p.capturedAt,timeZone) !== month || !p.relevant || !p.caption?.trim()) return false;
    if (hashes.has(p.hash || p.id)) return false;
    hashes.add(p.hash || p.id); return true;
  });
  // Round-robin activities and dates, ranked within each group, to avoid a report of repeated scenes.
  const groups = new Map();
  relevant.sort((a,b)=>(b.score||0)-(a.score||0)||a.capturedAt.localeCompare(b.capturedAt)).forEach(p=>{
    const key = `${p.activity||'Work activity'}:${new Date(p.capturedAt).toISOString().slice(0,10)}`;
    if(!groups.has(key)) groups.set(key,[]); groups.get(key).push(p);
  });
  const chosen=[]; const limit=photoLimit(relevant.length);
  while(chosen.length<limit && groups.size) for(const [key,group] of groups){if(chosen.length>=limit)break;chosen.push(group.shift());if(!group.length)groups.delete(key);}
  return {relevantCount:relevant.length,limit,photos:chosen.sort((a,b)=>a.capturedAt.localeCompare(b.capturedAt))};
}
