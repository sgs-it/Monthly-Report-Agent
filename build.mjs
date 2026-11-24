import {build} from 'esbuild';
import {mkdir,readFile,writeFile} from 'node:fs/promises';
await mkdir('dist/server',{recursive:true});await mkdir('public',{recursive:true});
const client=await build({entryPoints:['web/app.js'],bundle:true,write:false,minify:true,format:'iife',target:'es2022'});
const names=['house','folder','files','file-text','file-up','file-plus-2','link','building-2','layers','arrow-up-right','leaf','camera'];
const symbols=await Promise.all(names.map(async name=>{
  const svg=await readFile('node_modules/lucide-static/icons/'+name+'.svg','utf8');
  return '<symbol id="rg-icon-'+name+'" viewBox="0 0 24 24">'+svg.slice(svg.indexOf('>',svg.indexOf('<svg'))+1,svg.lastIndexOf('</svg>'))+'</symbol>';
}));
const sprite='<svg xmlns="http://www.w3.org/2000/svg" aria-hidden="true" width="0" height="0" style="position:absolute"><defs>'+symbols.join('')+'</defs></svg>';
const html=(await readFile('web/index.html','utf8')).replace('<!--RG_ICONS-->',sprite);
const garden=await readFile('web/garden.webp');
const css=(await readFile('web/style.css','utf8')).replace('__GARDEN_IMAGE__','data:image/webp;base64,'+garden.toString('base64'));
const manual=await readFile('web/manual.html','utf8');
await build({entryPoints:['worker/main.js'],outfile:'dist/server/index.js',bundle:true,format:'esm',platform:'browser',target:'es2022',minify:false,define:{__HTML__:JSON.stringify(html),__CSS__:JSON.stringify(css),__CLIENT__:JSON.stringify(client.outputFiles[0].text),__MANUAL__:JSON.stringify(manual)}});
await writeFile('public/health.txt','Field Report Agent build ready\n');
console.log('Built Field Report Agent for Vercel (no .openai manifest required)');
