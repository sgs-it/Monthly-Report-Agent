# Deploy the report agent on Vercel

This package is for Vercel. It does not read or require `.openai/hosting.json`.

## Update GitHub

Extract the ZIP and upload the contents of its `field-report-agent` folder into your repository. Keep `package.json`, `vercel.json`, `api`, `server`, `src`, `worker`, `web`, and `scripts` at the same repository level. Replace the previous files with these versions. Do not upload the ZIP itself as the application source.

## Vercel build settings

- Framework Preset: Other
- Root Directory: the directory containing `package.json` (usually the repository root)
- Node.js version: 24.x
- Build Command: `npm run build`
- Output Directory: `public`
- Install Command: `npm ci`

`vercel.json` supplies routing and build defaults. Reset any conflicting project-level build/output overrides to these values. The app uses a Node.js Vercel Function; it is not a static-only site.

## Configure persistent storage

1. Open your Vercel project and select Storage.
2. Select Create Storage and choose Blob.
3. Choose Private access.
4. Connect it to this project and select the option to add a read-write token environment variable. Use the default prefix so its name is `BLOB_READ_WRITE_TOKEN`.
5. Include the Production environment, plus Preview if you use preview deployments.

Templates, normalized photos, JSON metadata, and generated reports are stored in this private Blob store. Browser template uploads go directly to Blob through authenticated, short-lived upload tokens; this allows your masters larger than 4.5 MB. Downloads are streamed through the authenticated backend.

Official guide: https://vercel.com/docs/vercel-blob/client-upload

## Configure team sign-in

In Project Settings > Environment Variables, add:

| Name | Value |
| --- | --- |
| `APP_USERNAME` | `team`, or your chosen shared username |
| `APP_PASSWORD` | Your own password of at least 16 characters |
| `BLOB_READ_WRITE_TOKEN` | The token created by the private Blob connection |

Add them to the environments you deploy. Keep the actual values out of GitHub. `.env.example` contains only empty placeholders. The application returns a setup message until the password is configured; it does not expose the report library anonymously.

Redeploy after saving the environment variables. When you open the site, enter your shared username and password in the browser's sign-in prompt. This version uses a shared team login; individual accounts and per-user roles are not included.

## Import the masters

Open Projects & templates > Import template folder and choose your local master folder. The Vercel deployment starts with a separate library. The existing private Sites deployment and its R2 files are unchanged and are not automatically transferred to Vercel.

## CompanyCam

The CompanyCam connector is still pending. This package fixes Vercel hosting and storage; it does not activate CompanyCam or use a CompanyCam key yet. Do not put an API key into source files.

## Local development

Install Node.js 24. Copy `.env.example` to `.env.local` and fill in the team login and a development private Blob token. Then run:

```text
npm ci
npm run build
npm test
npm run validate
npm run dev
```

Open http://localhost:4173. This Vercel development entry point also uses the private Blob store; it does not use the original application's local-disk adapter. Use a separate development store when testing with disposable data. Blob completion callbacks are not required for saving the report but cannot reach localhost without a tunnel.

## Verification

The package is tested locally without `.openai/hosting.json`, including protected routing, staged large template uploads, private storage operations, report generation, and streamed Word downloads. Tests use simulated Blob storage. A real deployment and live Blob credentials have not been exercised in your Vercel account. The optional private-master Excel test skips when the master folder is absent.
