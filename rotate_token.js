// rotate_token.js — turn a one-time eBay authorisation code into a refresh
// token and store it encrypted in tokens/<slot>.enc (see vault.js).
//
// Run by the "Rotate eBay token" workflow. Inputs via env:
//   ACCOUNT    slot 1-6 (may be "4 lightingdepot"; the leading number is used)
//   AUTH_CODE  the code from the consent redirect: raw, percent-encoded, or the
//              whole redirect URL. Codes expire about 5 minutes after consent.
const { SCOPES, RU_NAME, storeRefreshToken } = require("./vault");

const ACCOUNT_NAMES = ["superfly", "aqualightingsolutions", "autolightingsolutions", "lightingdepot", "premiumlightingsolutions", "vividlighting"];
const TOKEN_URL = "https://api.ebay.com/identity/v1/oauth2/token";

async function tokenRequest(body) {
  const creds = Buffer.from(`${process.env.EBAY_APP_ID}:${process.env.EBAY_CERT_ID}`).toString("base64");
  const r = await fetch(TOKEN_URL, {
    method: "POST",
    headers: { Authorization: `Basic ${creds}`, "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(body),
  });
  const text = await r.text();
  if (!r.ok) throw new Error(`eBay token endpoint -> ${r.status} ${text.slice(0, 300)}`);
  return JSON.parse(text);
}

async function main() {
  const slot = parseInt(String(process.env.ACCOUNT || "").trim(), 10);
  if (!(slot >= 1 && slot <= 6)) throw new Error(`ACCOUNT must be 1-6, got "${process.env.ACCOUNT}"`);
  let code = String(process.env.AUTH_CODE || "").trim();
  const m = code.match(/[?&]code=([^&\s]+)/);
  if (m) code = m[1];
  if (/%[0-9A-Fa-f]{2}/.test(code)) code = decodeURIComponent(code);
  if (!code.startsWith("v^")) throw new Error("AUTH_CODE does not look like an eBay authorisation code");

  // 1. exchange the code (one use only, ~5 min lifetime)
  const j = await tokenRequest({ grant_type: "authorization_code", code, redirect_uri: RU_NAME });
  if (!j.refresh_token) throw new Error("eBay returned no refresh_token");
  console.log(`::add-mask::${j.refresh_token}`);
  console.log(`::add-mask::${j.access_token}`);

  // 2. prove the new refresh token works with the sync's own scopes before storing it
  const t = await tokenRequest({ grant_type: "refresh_token", refresh_token: j.refresh_token, scope: SCOPES });
  if (!t.access_token) throw new Error("Refresh test returned no access token");
  console.log(`::add-mask::${t.access_token}`);

  const days = Math.round((j.refresh_token_expires_in || 0) / 86400);
  const expires = new Date(Date.now() + (j.refresh_token_expires_in || 0) * 1000).toISOString().slice(0, 10);
  const file = storeRefreshToken(slot, j.refresh_token, { account: ACCOUNT_NAMES[slot - 1], updated: new Date().toISOString().slice(0, 10), expires });
  console.log(`Stored encrypted refresh token for slot ${slot} (${ACCOUNT_NAMES[slot - 1]}) in ${file}; valid ${days} days, until ${expires}.`);
}

main().catch((e) => { console.error(e.message || e); process.exit(1); });
