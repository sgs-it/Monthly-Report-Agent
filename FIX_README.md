# Fix the home page 404

Replace these two files at their existing GitHub paths:

- `vercel.json`
- `server/vercel-handler.js`

The updated configuration explicitly sends `/` to the Vercel function. The handler treats the function entry URL as the app's home page. Existing private storage, team login, and API routes remain in place.

In Vercel > Project Settings > Environment Variables, set `APP_PASSWORD` to your own password of at least 16 characters and include Production. `APP_USERNAME` defaults to `team` if it is not set. Keep `BLOB_READ_WRITE_TOKEN` from your private Blob connection.

Commit the file changes to GitHub and deploy that newest commit after saving the environment variables. The deployment must include both the source changes and the saved environment settings. Open the new deployment from Vercel and confirm the browser sign-in prompt appears.

The old deployment's `/api/index` currently reaches the backend and reports a missing/short password; its root URL returns a platform 404. These live checks establish the two problems, but the fixed configuration still needs deployment and verification in your Vercel account.
