/*
  CW Enquiry Assistant - settings
  Put the two n8n "Production URL" webhook addresses here (from the Webhook node in each workflow).
  Nothing secret goes in this file: it is public on GitHub Pages.
  The access key is entered by each person once, inside the add-in (gear icon).
*/
window.CW_CONFIG = {
  ANALYSE_URL: "https://cromwilkes.app.n8n.cloud/webhook/cw-enquiry-analyse",
  HUBSPOT_URL: "https://cromwilkes.app.n8n.cloud/webhook/cw-enquiry-hubspot",
  MAX_BODY_CHARS: 20000
};
