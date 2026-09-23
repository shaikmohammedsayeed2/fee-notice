# Fee Notice Studio

Open this folder in a terminal and run `npm install`, then `npm run dev`.

## Netlify deployment

Import the repository in Netlify and set the base directory to `fee-notice-studio`. Netlify will read `netlify.toml`, run `npm run build`, and publish `dist` automatically. No server functions, forms, databases, analytics, or uploads are used.

The app reads the Excel file through the browser's local File API and creates the PDF in memory on the user's device. Netlify only serves the public application files; it never receives the Excel sheet or the generated PDF.

Generated-by: client=codex, model_id=gpt-5.6-terra, reasoning_effort=medium
