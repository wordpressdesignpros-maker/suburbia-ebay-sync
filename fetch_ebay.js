// fetch_ebay.js  (income-form version)
//
// Polls eBay Sell Fulfillment orders across 6 accounts and writes each sold
// line item into a monthly tab of a OneDrive Excel workbook laid out like
// Simon's income/expenditure form:
//   INCOME:  A=DATE (DD.MM.YYYY)  B=PRODUCT (eBay title)  C=QUANTITY  D=£ (unit)
// A hidden "_state" sheet holds an order-line key per row so nothing is written
// twice. Month tabs are named e.g. "Jun 2026" and auto-created if missing.
//
// Env (GitHub secrets):
//   EBAY_APP_ID, EBAY_CERT_ID, EBAY_REFRESH_TOKEN_1..6
//   MS_CLIENT_ID, MS_TENANT_ID, MS_REFRESH_TOKEN (public client, no secret)
//   ONEDRIVE_FILE_ID

const ACCOUNTS = [1, 2, 3, 4, 5, 6].map((i) => ({ idx: i }));
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const STATE = "_state";

// ---------------------------------------------------------------------------
// eBay
// ---------------------------------------------------------------------------
async function ebayAccessToken(refreshToken) {
  const creds = Buffer.from(`${process.env.EBAY_APP_ID}:${process.env.EBAY_CERT_ID}`).toString("base64");
  const r = await fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "POST",
    headers: { Authorization: `Basic ${creds}`, "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({
      grant_type: "refresh_token",
      refresh_token: refreshToken,
      scope: "https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly",
    }),
  });
  if (!r.ok) throw new Error(`eBay token refresh failed: ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}

async function ebayFetchOrders(accessToken, sinceIso) {
  const orders = [];
  let url = `https://api.ebay.com/sell/fulfillment/v1/order?${new URLSearchParams({
    filter: `creationdate:[${sinceIso}..]`, limit: "200",
  })}`;
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } });
    if (!r.ok) throw new Error(`eBay order fetch failed: ${r.status} ${await r.text()}`);
    const j = await r.json();
    if (j.orders) orders.push(...j.orders);
    url = j.next || null;
  }
  return orders;
}

// ---------------------------------------------------------------------------
// Microsoft Graph (public client refresh token)
// ---------------------------------------------------------------------------
async function msToken() {
  const tenant = process.env.MS_TENANT_ID || "consumers";
  const params = {
    client_id: process.env.MS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: process.env.MS_REFRESH_TOKEN,
    scope: "Files.ReadWrite offline_access",
  };
  if (process.env.MS_CLIENT_SECRET) params.client_secret = process.env.MS_CLIENT_SECRET;
  const r = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams(params),
  });
  if (!r.ok) throw new Error(`Microsoft token refresh failed: ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}

const FID = () => process.env.ONEDRIVE_FILE_ID;
const WB = () => `https://graph.microsoft.com/v1.0/me/drive/items/${FID()}/workbook`;
const wsPath = (name) => `${WB()}/worksheets('${encodeURIComponent(name)}')`;

async function gfetch(token, url, opts = {}) {
  const r = await fetch(url, {
    ...opts,
    headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) },
  });
  if (!r.ok) throw new Error(`Graph ${opts.method || "GET"} ${url.split("/workbook")[1] || url} -> ${r.status} ${await r.text()}`);
  return r.status === 204 ? null : r.json();
}

async function listSheets(token) {
  const j = await gfetch(token, `${WB()}/worksheets?$select=name`);
  return (j.value || []).map((w) => w.name);
}

async function ensureMonthTab(token, name, existing) {
  if (existing.includes(name)) return;
  await gfetch(token, `${WB()}/worksheets/add`, { method: "POST", body: JSON.stringify({ name }) });
  // minimal headers so a runtime-created tab is still usable
  await gfetch(token, `${wsPath(name)}/range(address='A4:D4')`, {
    method: "PATCH", body: JSON.stringify({ values: [["DATE", "PRODUCT", "QUANTITY", "£"]] }),
  });
  await gfetch(token, `${wsPath(name)}/range(address='B2')`, {
    method: "PATCH", body: JSON.stringify({ values: [["INCOME"]] }),
  });
  existing.push(name);
}

// count income rows already present (column A from row 6)
async function nextIncomeRow(token, name) {
  const j = await gfetch(token, `${wsPath(name)}/range(address='A6:A5000')?$select=values`);
  const vals = j.values || [];
  let count = 0;
  for (const row of vals) {
    if (row[0] !== null && row[0] !== "") count++;
    else break; // contiguous block
  }
  return 6 + count;
}

async function readState(token) {
  const j = await gfetch(token, `${wsPath(STATE)}/usedRange?$select=values,rowCount`);
  const vals = j.values || [];
  const keys = new Set();
  for (let i = 1; i < vals.length; i++) if (vals[i][0]) keys.add(String(vals[i][0]));
  return { keys, nextRow: (j.rowCount || 1) + 1 };
}

function ddmmyyyy(iso) {
  const d = new Date(iso);
  const p = (n) => String(n).padStart(2, "0");
  return `${p(d.getUTCDate())}.${p(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`;
}
function tabFor(iso) {
  const d = new Date(iso);
  return `${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`;
}

// Sort a month tab's income block (A6:D...) by date ascending. Dates are
// "DD.MM.YYYY" text; within one month tab that sorts chronologically.
async function sortTab(token, tab) {
  const j = await gfetch(token, `${wsPath(tab)}/range(address='A6:D5000')?$select=values`);
  const rows = (j.values || []).filter((r) => r[0] !== null && r[0] !== "");
  if (rows.length < 2) return;
  const k = (s) => { const [d, m, y] = String(s).split("."); return (+y) * 10000 + (+m) * 100 + (+d); };
  rows.sort((a, b) => k(a[0]) - k(b[0]));
  await gfetch(token, `${wsPath(tab)}/range(address='A6:D${5 + rows.length}')`, {
    method: "PATCH", body: JSON.stringify({ values: rows }),
  });
  console.log(`Sorted ${tab} (${rows.length} rows)`);
}

function defaultSince() {
  // From the 1st of the current month, so a full month is always captured.
  // De-dup (the _state sheet) prevents anything being written twice.
  const now = new Date();
  return new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1, 0, 0, 0)).toISOString();
}

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
async function main() {
  if (!FID()) throw new Error("ONEDRIVE_FILE_ID not set");
  const token = await msToken();
  const { keys: seen, nextRow: stateNextRow } = await readState(token);
  let sheets = await listSheets(token);
  const since = defaultSince();

  // gather new line items grouped by month tab
  const byTab = {}; // tab -> array of [date, title, qty, unitPrice]
  const newKeys = [];

  for (const acc of ACCOUNTS) {
    const refresh = process.env[`EBAY_REFRESH_TOKEN_${acc.idx}`];
    if (!refresh) { console.log(`Account ${acc.idx}: no token, skipping`); continue; }
    try {
      const access = await ebayAccessToken(refresh);
      const orders = await ebayFetchOrders(access, since);
      let added = 0;
      for (const o of orders) {
        for (const li of o.lineItems || []) {
          const key = `${o.orderId}|${li.lineItemId}`;
          if (seen.has(key)) continue;
          seen.add(key);
          const qty = li.quantity || 1;
          const lineCost = parseFloat(li.lineItemCost?.value || 0);
          const unit = qty ? Math.round((lineCost / qty) * 100) / 100 : lineCost;
          const tab = tabFor(o.creationDate);
          (byTab[tab] = byTab[tab] || []).push([ddmmyyyy(o.creationDate), li.title || "", qty, unit]);
          newKeys.push(key);
          added++;
        }
      }
      console.log(`Account ${acc.idx}: ${orders.length} orders, ${added} new line items`);
    } catch (e) {
      console.error(`Account ${acc.idx} failed: ${e.message}`);
    }
  }

  const tabs = Object.keys(byTab);

  let written = 0;
  for (const tab of tabs) {
    await ensureMonthTab(token, tab, sheets);
    const rows = byTab[tab];
    const start = await nextIncomeRow(token, tab);
    const end = start + rows.length - 1;
    await gfetch(token, `${wsPath(tab)}/range(address='A${start}:D${end}')`, {
      method: "PATCH", body: JSON.stringify({ values: rows }),
    });
    console.log(`${tab}: wrote ${rows.length} rows (A${start}:D${end})`);
    written += rows.length;
  }

  // record keys in _state
  if (newKeys.length) {
    const sStart = stateNextRow;
    const sEnd = sStart + newKeys.length - 1;
    await gfetch(token, `${wsPath(STATE)}/range(address='A${sStart}:A${sEnd}')`, {
      method: "PATCH", body: JSON.stringify({ values: newKeys.map((k) => [k]) }),
    });
  }

  // keep affected tabs + the current month tab sorted by date
  const toSort = new Set(tabs);
  toSort.add(tabFor(new Date().toISOString()));
  for (const tab of toSort) {
    if (sheets.includes(tab)) {
      try { await sortTab(token, tab); } catch (e) { console.error(`sort ${tab} failed: ${e.message}`); }
    }
  }

  console.log(`Wrote ${written} income rows across ${tabs.length} tab(s); recorded ${newKeys.length} keys.`);
}

main().catch((e) => { console.error(e); process.exit(1); });
