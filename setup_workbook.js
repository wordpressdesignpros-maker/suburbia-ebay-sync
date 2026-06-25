// setup_workbook.js
// One-time: uploads the redesigned suburbia-ebay-auto.xlsx (committed in this
// repo) over the existing OneDrive file, replacing its content (same file ID).
// Runs on GitHub, so no PC involvement. The OneDrive file must be CLOSED.

const fs = require("fs");

async function msToken() {
  const tenant = process.env.MS_TENANT_ID || "consumers";
  const p = { client_id: process.env.MS_CLIENT_ID, grant_type: "refresh_token", refresh_token: process.env.MS_REFRESH_TOKEN, scope: "Files.ReadWrite offline_access" };
  if (process.env.MS_CLIENT_SECRET) p.client_secret = process.env.MS_CLIENT_SECRET;
  const r = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams(p) });
  if (!r.ok) throw new Error(`MS token failed: ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}

async function main() {
  const token = await msToken();
  const data = fs.readFileSync("suburbia-ebay-auto.xlsx");
  const r = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${process.env.ONEDRIVE_FILE_ID}/content`, {
    method: "PUT", headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/octet-stream" }, body: data,
  });
  if (!r.ok) throw new Error(`Upload failed: ${r.status} ${await r.text()}`);
  console.log("Workbook uploaded to OneDrive (" + data.length + " bytes).");
}
main().catch((e) => { console.error(e); process.exit(1); });
