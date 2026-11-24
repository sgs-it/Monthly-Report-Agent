export function filenameLocation(filename='') {
  return filename.replace(/\.(docx|xlsx)$/i,'')
    .replace(/^.*?Monthly Report\s*[-–]?\s*/i,'')
    .replace(/[-\s]+(?:January|February|March|April|May|June|July|August|September|October|November|December)\s+20\d\d$/i,'')
    .trim();
}

export function projectLabel(project) {
  if (!project) return '';
  const {location} = project.wordIdentity || {};
  const site = location || (/^(location|job\s*code|job\s*number)$/i.test(project.name) ? filenameLocation(project.templateName) : project.name);
  const code = projectJobCode(project);
  return code && !site.endsWith(code) ? site+' — '+code : site;
}

export function projectJobCode(project) {
  return project?.wordIdentity?.jobCode || project?.code || (project && project.format!=='xlsx' ? filenameJobCode(project.templateName) : '');
}

// Some masters have service text in the job-code cell but end their filename with the job number.
export function filenameJobCode(filename) {
  const match = filenameLocation(filename).match(/(?:^|[-–\s])(\d{2}[AB](?:&[AB])?\d{4,}(?:\s*[,/&]\s*(?:\d{2}[AB])?\d+)*)$/i);
  return match?.[1] || '';
}
