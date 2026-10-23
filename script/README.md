# Google Apps Script (backend)

`Code.gs` here is the source of truth for the script behind the Google Sheet.
Pushing a change to `script/Code.gs` deploys it automatically (see `.github/workflows/deploy-script.yml`)
to the existing web app deployment, so the app's address never changes.

Bump `SCRIPT_VERSION` at the top of `Code.gs` with every change; it becomes the deployment description,
shows in the app (⚙), and is written to `LiveVersion` on the Sheet's Settings tab.
