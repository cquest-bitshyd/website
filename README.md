We have created this repo to host c-quest website.

This code is connected to a google sheet as a database for the events:https://docs.google.com/spreadsheets/d/1geWhEdA7mwI2qLUwtXMpKKcG0UA_HLsibcCV-Pfe7g4/edit?usp=sharing

Image proxy setup (Google Apps Script)
- Create a new Apps Script project and paste the code from the file [apps-script-proxy.js](apps-script-proxy.js).
- Deploy it as a web app with "Execute as: Me" and "Who has access: Anyone".
- Copy the web app URL and set it before the page script runs:
  ```html
  <script>
    window.CQUEST_IMAGE_PROXY_BASE = 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec';
  </script>
  ```
- If you leave the value empty, the site will keep using the current Drive thumbnail fallback.
