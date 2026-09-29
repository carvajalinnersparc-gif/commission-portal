# Commission Portal

Inner SPARC Realty agent and administration portal built with Google Apps Script and HTML Service.

## Project files

- `Code.js`: web app entry point and commission backend.
- `Portal.html`: active portal interface, styles, and client scripts.
- `Auth.js`: authentication and access management.
- `PortalData.js`: portal data queries.
- `EmailCampaign.js` and `EmailCampaigner.html`: email campaign features.
- `appsscript.json`: Apps Script project settings.

`Portal_CSS.html` and `Portal_JS.html` are legacy files; the active portal is self-contained in `Portal.html`.

## Deployment

Upload the source files to the intended Apps Script project, then update the web app deployment to a new version. Confirm the spreadsheet IDs and project settings match the intended environment before deployment. The local `.clasp.json` deployment mapping is excluded from Git.

## Analytics

See [ANALYTICS.md](ANALYTICS.md) for the Google Analytics measurement ID, configuration, and verification steps.
