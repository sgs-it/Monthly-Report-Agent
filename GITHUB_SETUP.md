# GitHub source package

Extract this ZIP and upload its contents to your repository. Uploading the ZIP alone does not extract it into a runnable GitHub project.

## Run locally

Requires Node.js 22 or newer.

```text
npm ci
npm run build
npm test
npm run validate
npm run dev
```

Open http://localhost:4173. The server creates a local .local-data folder.

## Included and excluded

Application source, the manual editor, package lock file, tests, and a synthetic test image are included. Customer master templates, uploaded photos, generated reports, API keys, Git history, and node_modules are excluded. Import the master files through the app or place them in the ignored Master report templates directory to run optional template tests. Tests requiring those private Excel masters skip when the directory is absent.

## Hosting

GitHub stores the source; uploading it does not publish or move the current hosted data. The included .openai/hosting.json references the existing Sites project. The deployed app uses a Cloudflare Worker with the BUCKET R2 binding and Sites authentication. GitHub Pages alone cannot run this backend. If hosting elsewhere, configure storage and authentication before exposing its API. Local storage and hosted R2 storage are independent.

CompanyCam integration is not implemented in this version. No CompanyCam key is included.
