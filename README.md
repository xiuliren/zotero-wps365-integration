# Zotero WPS 365 Integration

Zotero WPS 365 Integration is a Zotero integration plugin that communicates between WPS 365 and Zotero via the Connector.

## Back-end

The Apps Script code should be deployed as an API executable and its URL set in `zotero_config.js`.

Field codes and doc preferences are stored as `NamedRanges` 
within the document by serializing them as range names.
All text updates are done in the back-end.
The document updates via the back-end are batched in the connector
to reduce the processing time because of server latency.
To learn more read the [Apps Script reference](https://developers.google.com/apps-script/reference/document/document).

## Front-end

The connector front-end adds Zotero UI elements to the WPS 365 editor
and acts as a glue-layer between the back-end and the front end. The connector
also:
- Inserts initial citation placeholders because we have no access to the user
cursor from the back-end when the Apps Script code is deployed as an API
executable
- Performs citation conversions to footnotes because the Apps Script has no
API to insert footnotes into the document, only edit or remove them.

## Apps Script Development and Deployment

You can develop and deploy Apps Script code using [clasp](https://developers.google.com/apps-script/guides/clasp)

```bash
cd src/apps-script
clasp login
# Authenticate in the browser

clasp clone <scriptId>
# If you had unpushed changes, they are overwritten, so
git checkout -- .

# Make some changes in Code.js
clasp push
# Changes available in dev version
```