# Field Report Agent for Vercel

See [VERCEL_SETUP.md](VERCEL_SETUP.md) for the complete deployment steps.

This version runs the report backend in a Node.js Vercel Function and stores templates, photos, JSON metadata, and generated reports in a PRIVATE Vercel Blob store. It does not require an OpenAI Sites hosting manifest.

## Deploy

1. Upload the extracted source files to your GitHub repository. Keep package.json, vercel.json, api, server, src, worker, web, and scripts at the repository root.
2. In Vercel, use Framework Other, Node.js 24.x, Build Command npm run build, Output Directory public, and Install Command npm ci.
3. Connect a private Blob store with the read-write token option so BLOB_READ_WRITE_TOKEN is added to the project.
4. Deploy and open your production website address. Anyone with the address can use the workspace.
5. Import your masters from Projects & templates > Import template folder.

The existing private Sites app and its files are separate. This package does not migrate that R2 library or modify the existing hosted app. Customer templates, photos, reports, node_modules, Git history, and secrets are excluded from the ZIP.

## Reports

Word masters control the document layout, tables, headers, footers, and static content. Excel masters retain their workbook layout for Excel output and receive a separate Word rendering of worksheet sections. Original templates are kept as the basis for generated copies.

Only relevant work photos from the chosen project and capture month are eligible. The normal maximum is 50 photos; it becomes 75 only when more than 200 unique relevant photos are available. CompanyCam imports use captured_at in each site's time zone. Imported photos require caption and work-activity review before report selection. Manual uploads still require capture dates. Monthly narrative, statuses, readings, signatures, and other unsupported facts still need review and template mapping.

## CompanyCam

Add COMPANYCAM_API_KEY as a Production Secret in Vercel, then deploy this source. The app verifies the key through CompanyCam and shows its connection status. Open Photo library, choose your report project, link the exact CompanyCam project, select the month, and import its photos. Review the imported work photos and their captions before generating a report. Photo pagination continues until all pages are processed; interrupted imports can be retried without overwriting existing review decisions. The key stays on the server. No key is included in this package. See VERCEL_SETUP.md for details and troubleshooting.

The browser suggests approved captions using a locally running image model. Review the activity before including each photo. Storage in this package continues to use private Vercel Blob; Supabase migration is not included.

## Development and checks

Use Node.js 24. Copy .env.example to .env.local and configure a separate development private Blob store. Run npm ci, npm run build, npm test, npm run validate, then npm run dev. The development entry point uses private Blob storage.

The build, direct access, storage adapter, staged uploads larger than 4.5 MB, report generation, streamed downloads, and CompanyCam connection/month/project/pagination/retry rules are tested locally. Tests use simulated external services and a real JPEG image processor. The private-master Excel test skips when the customer master folder is absent. Your actual CompanyCam key and live Vercel deployment have not been tested. Automatic visual Word rendering was unavailable during the earlier document checks.
