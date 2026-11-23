import assert from 'node:assert/strict';import {readFile} from 'node:fs/promises';
const config=JSON.parse(await readFile('vercel.json','utf8'));assert.equal(config.outputDirectory,'public');assert.ok(config.rewrites.some(r=>r.destination.startsWith('/api/index')));
const app=await import('../api/index.js');assert.equal(typeof app.default.fetch,'function');
const r=await app.default.fetch(new Request('https://example.test/'));assert.equal(r.status,503);console.log('Vercel entry point and protected setup validated; no .openai manifest needed.');
