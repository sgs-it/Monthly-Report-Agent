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

Templates, normalized photos, JSON metadata, and generated reports are stored in this private Blob store. Browser template uploads go directly to Blob through short-lived upload tokens; this allows your masters larger than 4.5 MB. Downloads are streamed through the report backend.

Official guide: https://vercel.com/docs/vercel-blob/client-upload

## Open the workspace

Open the normal deployment URL. Anyone with access to this address can view, change, and download workspace data and import photos from the connected CompanyCam account.

## Import the masters

Open Master files > Import template folder and choose your local master folder. The Vercel deployment starts with a separate library. The existing private Sites deployment and its R2 files are unchanged and are not automatically transferred to Vercel.

## CompanyCam

CompanyCam imports suggest a caption from the approved landscaping list using a model downloaded to the browser on first use. Review the suggestion or choose another caption on the Generate report screen. This requires no captioning API key.

Add COMPANYCAM_API_KEY as a Secret in Vercel Project Settings > Environment Variables for Production, then deploy this updated source. Use an active CompanyCam Application Key with permission to read projects and photos. No CompanyCam credential is sent to the browser. A direct Application Key connection does not use an OAuth redirect URL.

The app checks CompanyCam and displays Connected only after an API request succeeds. In Integrations, select your report project, click Link CompanyCam project, and choose the matching remote project. Open Generate monthly report from the dashboard, choose the reporting month, and import the photos. Type a project name in the CompanyCam search field and select the matching address. Load more retrieves another page of the current search; photo pages are imported until the complete result is reached. Repeated imports preserve existing photos and review decisions. Interrupted imports show an error; retry to resume.

Only photos whose captured_at falls in the selected month in the site's time zone are imported. Upload dates are never substituted. Photos without capture dates are excluded. Each imported photo starts unmarked: review it, supply or check its caption, and mark work activity / completed work. The existing policy then selects at most 50 photos, or 75 when more than 200 unique relevant photos have been reviewed. The browser suggests approved captions using a locally running image model. Review the activity before including each photo.

Images are resized to JPEG on the server. Sources must use CompanyCam HTTPS image hosts; a photo without a supported image URL must be uploaded manually. Imports do not write to CompanyCam. This package continues to use the existing private Vercel Blob storage; it does not migrate storage to Supabase.

If the status says Not connected, use Check connection and follow the returned error. Check that the variable is named exactly COMPANYCAM_API_KEY, contains the Application Key rather than an OAuth client ID, is saved for Production, and has been included in a fresh deployment.

## Local development

Install Node.js 24. Copy `.env.example` to `.env.local` and fill in a development private Blob token. Then run:

```text
npm ci
npm run build
npm test
npm run validate
npm run dev
```

Open http://localhost:4173. This Vercel development entry point also uses the private Blob store; it does not use the original application's local-disk adapter. Use a separate development store when testing with disposable data. Blob completion callbacks are not required for saving the report but cannot reach localhost without a tunnel.

## Verification

The package is tested locally without `.openai/hosting.json`, including direct routing, staged large template uploads, private storage operations, report generation, and streamed Word downloads. Tests use simulated Blob storage. A real deployment and live Blob credentials have not been exercised in your Vercel account. The optional private-master Excel test skips when the master folder is absent.
