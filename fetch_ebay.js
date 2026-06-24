// fetch_ebay.js
//
// Polls eBay's Sell Fulfillment API for new orders across multiple accounts
// and appends them to a Microsoft OneDrive Excel workbook via Graph API.
//
// State is held inside the spreadsheet itself: for each account we look up
// the most recent orderDate already on the Sales tab and only fetch newer
// orders. No external state file required.
//
// Required environment variables (set as GitHub Secrets):
//   EBAY_APP_ID           - eBay developer App ID (Client ID)
//   EBAY_CERT_ID          - eBay developer Cert ID (Client Secret)
//   EBAY_REFRESH_TOKEN_1  - refresh token for eBay account 1
//   EBAY_REFRESH_TOKEN_2  - refresh token for eBay account 2
//   ... up to EBAY_REFRESH_TOKEN_6
//   MS_CLIENT_ID          - Azure AD app client ID
//   MS_CLIENT_SECRET      - Azure AD app client secret
//   MS_TENANT_ID          - Azure AD tenant ID (or "consumers" for personal account)
//   MS_REFRESH_TOKEN      - Microsoft Graph refresh token
//   ONEDRIVE_FILE_ID      - drive item ID of the spreadsheet

const ACCOUNTS = [
  { idx: 1, name: "Account 1" },
  { idx: 2, name: "Account 2" },
  { idx: 3, name: "Account 3" },
  { idx: 4, name: "Account 4" },
  { idx: 5, name: "Account 5" },
  { idx: 6, name: "Account 6" },
];

const TABLE_NAME = "tblSales";

// ---------------------------------------------------------------------------
// eBay OAuth helpers
// ---------------------------------------------------------------------------

async function ebayAccessToken(refreshToken) {
  const creds = Buffer
    .from(`${process.env.EBAY_APP_ID}:${process.env.EBAY_CERT_ID}`)
    .toString("base64");
  const body = new URLSearchParams({
    grant_type: "refresh_token",
    refresh_token: refreshToken,
    scope: "https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly",
  });
  const r = await fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "POST",
    headers: {
      Authorization: `Basic ${creds}`,
      "Content-Type": "application/x-www-form-urlencoded",
    },
    body,
  });
  if (!r.ok) throw new Error(`eBay token refresh failed: ${r.status} ${await r.text()}`);
  const j = await r.json();
  return j.access_token;
}

async function ebayFetchOrders(accessToken, sinceIso) {
  const params = new URLSearchParams({
    filter: `creationdate:[${sinceIso}..]`,
    limit: "200",
  });
  const orders = [];
  let url = `https://api.ebay.com/sell/fulfillment/v1/order?${params}`;

  while (url) {
    const r = await fetch(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
    });
    if (!r.ok) throw new Error(`eBay order fetch failed: ${r.status} ${await r.text()}`);
    const j = await r.json();
    if (j.orders) orders.push(...j.orders);
    url = j.next || null;
  }
  return orders;
}

// ---------------------------------------------------------------------------
// Microsoft Graph helpers
// ---------------------------------------------------------------------------

async function msAccessToken() {
  const tenant = process.env.MS_TENANT_ID || "consumers";
  const params = {
    client_id: process.env.MS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: process.env.MS_REFRESH_TOKEN,
    scope: "Files.ReadWrite offline_access",
  };
  // Public client (no secret) is used by default. Only send client_secret
  // if one is configured (confidential client / dedicated Azure app).
  if (process.env.MS_CLIENT_SECRET) {
    params.client_secret = process.env.MS_CLIENT_SECRET;
  }
  const body = new URLSearchParams(params);
  const r = await fetch(
    `https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`,
    {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
    },
  );
  if (!r.ok) throw new Error(`Microsoft token refresh failed: ${r.status} ${await r.text()}`);
  const j = await r.json();
  return j.access_token;
}

async function msReadTableRows(token, fileId, tableName) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}/range`;
  const r = await fetch(url, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!r.ok) throw new Error(`Read table failed: ${r.status} ${await r.text()}`);
  const j = await r.json();
  return j.values || [];
}

async function msAppendRows(token, fileId, tableName, rows) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/items/${fileId}/workbook/tables/${encodeURIComponent(tableName)}/rows`;
  const r = await fetch(url, {
    method: "POST",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ values: rows }),
  });
  if (!r.ok) throw new Error(`Append rows failed: ${r.status} ${await r.text()}`);
  return r.json();
}

// ---------------------------------------------------------------------------
// Determine "lastSeen" per account from existing Sales tab data
// ---------------------------------------------------------------------------

// Normalise a date cell to an ISO 8601 string. Excel/Graph may return a date
// either as an ISO string OR as an Excel serial number (days since 1899-12-30,
// where 25569 == 1970-01-01). Anything unparseable returns null and is ignored.
function toIso(val) {
  if (val === null || val === undefined || val === "") return null;
  const num =
    typeof val === "number"
      ? val
      : /^\d+(\.\d+)?$/.test(String(val).trim())
      ? parseFloat(val)
      : NaN;
  if (!isNaN(num) && num > 20000 && num < 80000) {
    // plausible Excel serial date
    return new Date(Math.round((num - 25569) * 86400 * 1000)).toISOString();
  }
  const d = new Date(val);
  return isNaN(d.getTime()) ? null : d.toISOString();
}

function lastSeenFromSheet(values) {
  const lastSeen = {};
  if (!values || values.length < 2) return lastSeen;
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const iso = toIso(row[0]);
    const account = row[1];
    if (!iso || !account) continue;
    const cur = lastSeen[account];
    if (!cur || iso > cur) lastSeen[account] = iso;
  }
  return lastSeen;
}

function defaultSince() {
  return new Date(Date.now() - 7 * 86400000).toISOString();
}

// ---------------------------------------------------------------------------
// Build rows from eBay orders (one row per line item, fees split proportionally)
// ---------------------------------------------------------------------------

function ordersToRows(orders, accountName) {
  const rows = [];
  for (const o of orders) {
    const lis = o.lineItems || [];
    const feeTotal = parseFloat(o.totalFeeBasisAmount?.value || 0);
    const sumItems = lis.reduce(
      (s, li) => s + parseFloat(li.lineItemCost?.value || 0), 0,
    ) || 1;
    for (const li of lis) {
      const it = parseFloat(li.lineItemCost?.value || 0);
      const post = parseFloat(li.deliveryCost?.shippingCost?.value || 0);
      const qty = li.quantity || 1;
      rows.push([
        o.creationDate,
        accountName,
        o.orderId,
        li.title || "",
        li.sku || "",
        qty,
        it / qty,
        it,
        post,
        it + post,
        Math.round((it / sumItems) * feeTotal * 100) / 100,
        o.orderPaymentStatus === "PAID" ? "Paid" : "Pending",
      ]);
    }
  }
  return rows;
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------

async function main() {
  const fileId = process.env.ONEDRIVE_FILE_ID;
  if (!fileId) throw new Error("ONEDRIVE_FILE_ID env var not set");

  const msToken = await msAccessToken();
  const sheet = await msReadTableRows(msToken, fileId, TABLE_NAME);
  const lastSeen = lastSeenFromSheet(sheet);
  const fallback = defaultSince();

  const allRows = [];
  for (const acc of ACCOUNTS) {
    const refresh = process.env[`EBAY_REFRESH_TOKEN_${acc.idx}`];
    if (!refresh) {
      console.log(`Skipping ${acc.name} — no refresh token configured`);
      continue;
    }
    try {
      const since = lastSeen[acc.name] || fallback;
      console.log(`${acc.name}: fetching orders since ${since}`);
      const access = await ebayAccessToken(refresh);
      const orders = await ebayFetchOrders(access, since);
      const rows = ordersToRows(orders, acc.name);
      console.log(`${acc.name}: ${orders.length} orders → ${rows.length} rows`);
      allRows.push(...rows);
    } catch (e) {
      console.error(`${acc.name} failed: ${e.message}`);
    }
  }

  if (allRows.length === 0) {
    console.log("Nothing new to write.");
    return;
  }

  const writeToken = await msAccessToken();
  await msAppendRows(writeToken, fileId, TABLE_NAME, allRows);
  console.log(`Wrote ${allRows.length} rows to ${TABLE_NAME}.`);
}

main().catch(e => {
  console.error(e);
  process.exit(1);
});
