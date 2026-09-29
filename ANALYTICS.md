# Google Analytics setup

The portal uses Google Analytics 4 (GA4) measurement ID `G-F9V2PHXFBB` by default.

1. Open the Google Analytics Web data stream for `G-F9V2PHXFBB`.
2. In that stream, turn **Enhanced measurement off**. The portal sends its own page views; automatic form, click, and history collection is unnecessary for this authenticated app.
3. No script property is required. To use a different stream, set `GA4_MEASUREMENT_ID` in the Apps Script project's **Project Settings > Script properties**.
4. Upload the updated `Code.js` and `Portal.html` to Apps Script and update the web app deployment to a new version.
5. Open the deployed portal, sign in, and navigate between screens. Check the GA4 Realtime report for page views. Ad blockers and browser privacy settings can prevent collection.

Page views use allowlisted screen names. Analytics URLs exclude session tokens, login tokens, sale IDs, and other incoming URL parameters. Consecutive views of the same screen are deduplicated. No names, emails, commission amounts, or form values are explicitly sent. Google signals and ad personalization signals are disabled.

Set the script property `GA4_MEASUREMENT_ID` to `off` to disable tracking on subsequent page loads. Removing the property restores the default ID. Live data collection must be verified after deploying.

Reference: [Google's GA4 page-view documentation](https://developers.google.com/analytics/devguides/collection/ga4/views).
