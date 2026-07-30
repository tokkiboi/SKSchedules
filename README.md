# SKSchedules

Static Git repository export of the StyleKorean Logistics Planner originally hosted at:

<https://stylekorean-logistics-planner.alex481942.chatgpt.site/>

## Run locally

Serve the repository root with any static HTTP server:

```bash
python3 -m http.server 8080
```

Then open <http://localhost:8080/>.

## Deployment

GitHub Pages deploys the repository root from the `main` branch.

## Export limitation

The public application reads schedule data directly from its configured Google Sheets. Status edits in the original hosted application post to a private `/api/status` server endpoint; that server implementation is not present in the public deployment and cannot be recovered from the compiled browser assets.
