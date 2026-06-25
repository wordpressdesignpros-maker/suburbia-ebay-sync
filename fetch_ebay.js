// fetch_ebay.js  (income + expenditure form version)
//
// INCOME  (cols A-D): eBay sales -> DATE / PRODUCT / QUANTITY / £ (unit)
// EXPENDITURE (cols K-N): eBay final-value fees + postage label costs from the
//   eBay Finances API -> DATE / DESCRIPTION / CATEGORY / £
// Each row keyed in the hidden "_state" sheet so nothing is written twice.
// Cancelled / fully-refunded orders are skipped on the income side.
//
// Env (GitHub secrets): EBAY_APP_ID, EBAY_CERT_ID, EBAY_REFRESH_TOKEN_1..6,
//   MS_CLIENT_ID, MS_TENANT_ID, MS_REFRESH_TOKEN, ONEDRIVE_FILE_ID

const ACCOUNTS = [1, 2, 3, 4, 5, 6].map((i) => ({ idx: i }));
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const STATE = "_state";
const SCOPES = "https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly https://api.ebay.com/oauth/api_scope/sell.finances";

// ---------------------------------------------------------------------------
// eBay auth + APIs
// ---------------------------------------------------------------------------
async function ebayAccessToken(refreshToken) {
  const creds = Buffer.from(`${process.env.EBAY_APP_ID}:${process.env.EBAY_CERT_ID}`).toString("base64");
  const r = await fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "POST",
    headers: { Authorization: `Basic ${creds}`, "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: refreshToken, scope: SCOPES }),
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

async function ebayFetchTransactions(accessToken, sinceIso) {
  const txns = [];
  let url = `https://apiz.ebay.com/sell/finances/v1/transaction?${new URLSearchParams({
    filter: `transactionDate:[${sinceIso}..]`, limit: "200",
  })}`;
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } });
    if (!r.ok) throw new Error(`eBay finances fetch failed: ${r.status} ${await r.text()}`);
    const j = await r.json();
    if (j.transactions) txns.push(...j.transactions);
    url = j.next || null;
  }
  return txns;
}

function feeOfSale(t) {
  if (t.totalFeeAmount && t.totalFeeAmount.value) return Math.round(parseFloat(t.totalFeeAmount.value) * 100) / 100;
  let f = 0;
  for (const li of t.orderLineItems || []) for (const mf of li.marketplaceFees || []) f += parseFloat(mf.amount?.value || 0);
  return Math.round(f * 100) / 100;
}

// ---------------------------------------------------------------------------
// Microsoft Graph
// ---------------------------------------------------------------------------
async function msToken() {
  const tenant = process.env.MS_TENANT_ID || "consumers";
  const params = {
    client_id: process.env.MS_CLIENT_ID, grant_type: "refresh_token",
    refresh_token: process.env.MS_REFRESH_TOKEN, scope: "Files.ReadWrite offline_access",
  };
  if (process.env.MS_CLIENT_SECRET) params.client_secret = process.env.MS_CLIENT_SECRET;
  const r = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
    method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams(params),
  });
  if (!r.ok) throw new Error(`Microsoft token refresh failed: ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}

const FID = () => process.env.ONEDRIVE_FILE_ID;
const WB = () => `https://graph.microsoft.com/v1.0/me/drive/items/${FID()}/workbook`;
const wsPath = (name) => `${WB()}/worksheets('${encodeURIComponent(name)}')`;

async function gfetch(token, url, opts = {}) {
  const r = await fetch(url, { ...opts, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) } });
  if (!r.ok) throw new Error(`Graph ${opts.method || "GET"} ${(url.split("/workbook")[1] || url).slice(0, 80)} -> ${r.status} ${await r.text()}`);
  return r.status === 204 ? null : r.json();
}

async function listSheets(token) {
  const j = await gfetch(token, `${WB()}/worksheets?$select=name`);
  return (j.value || []).map((w) => w.name);
}

async function ensureMonthTab(token, name, existing) {
  if (existing.includes(name)) return;
  await gfetch(token, `${WB()}/worksheets/add`, { method: "POST", body: JSON.stringify({ name }) });
  await gfetch(token, `${wsPath(name)}/range(address='A4:D4')`, { method: "PATCH", body: JSON.stringify({ values: [["DATE", "PRODUCT", "QUANTITY", "£"]] }) });
  await gfetch(token, `${wsPath(name)}/range(address='K4:N4')`, { method: "PATCH", body: JSON.stringify({ values: [["DATE", "DESCRIPTION", "CATEGORY", "£"]] }) });
  await gfetch(token, `${wsPath(name)}/range(address='B2')`, { method: "PATCH", body: JSON.stringify({ values: [["INCOME"]] }) });
  await gfetch(token, `${wsPath(name)}/range(address='K2')`, { method: "PATCH", body: JSON.stringify({ values: [["EXPENDITURE"]] }) });
  existing.push(name);
}

async function nextRowIn(token, name, col) {
  const j = await gfetch(token, `${wsPath(name)}/range(address='${col}6:${col}5000')?$select=values`);
  const vals = j.values || [];
  let count = 0;
  for (const row of vals) { if (row[0] !== null && row[0] !== "") count++; else break; }
  return 6 + count;
}

async function readState(token) {
  const j = await gfetch(token, `${wsPath(STATE)}/usedRange?$select=values,rowCount`);
  const vals = j.values || [];
  const keys = new Set();
  for (let i = 1; i < vals.length; i++) if (vals[i][0]) keys.add(String(vals[i][0]));
  return { keys, nextRow: (j.rowCount || 1) + 1 };
}

function ddmmyyyy(iso) { const d = new Date(iso); const p = (n) => String(n).padStart(2, "0"); return `${p(d.getUTCDate())}.${p(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`; }
function tabFor(iso) { const d = new Date(iso); return `${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`; }
function dateKey(s) { const [d, m, y] = String(s).split("."); return (+y) * 10000 + (+m) * 100 + (+d); }

async function sortBlock(token, name, startCol, endCol) {
  const j = await gfetch(token, `${wsPath(name)}/range(address='${startCol}6:${endCol}5000')?$select=values`);
  const rows = (j.values || []).filter((r) => r[0] !== null && r[0] !== "");
  if (rows.length < 2) return;
  rows.sort((a, b) => dateKey(a[0]) - dateKey(b[0]));
  await gfetch(token, `${wsPath(name)}/range(address='${startCol}6:${endCol}${5 + rows.length}')`, { method: "PATCH", body: JSON.stringify({ values: rows }) });
}

function defaultSince() { const n = new Date(); return new Date(Date.UTC(n.getUTCFullYear(), n.getUTCMonth(), 1, 0, 0, 0)).toISOString(); }

// ---------------------------------------------------------------------------
// Main
// ---------------------------------------------------------------------------
async function main() {
  if (!FID()) throw new Error("ONEDRIVE_FILE_ID not set");
  const token = await msToken();
  const { keys: seen, nextRow: stateNextRow } = await readState(token);
  let sheets = await listSheets(token);
  const since = defaultSince();

  const incByTab = {};   // tab -> [[date,title,qty,unit]]
  const expByTab = {};   // tab -> [[date,desc,category,amount]]
  const newKeys = [];

  for (const acc of ACCOUNTS) {
    const refresh = process.env[`EBAY_REFRESH_TOKEN_${acc.idx}`];
    if (!refresh) { console.log(`Account ${acc.idx}: no token, skipping`); continue; }
    try {
      const access = await ebayAccessToken(refresh);

      // INCOME
      const orders = await ebayFetchOrders(access, since);
      let inc = 0;
      for (const o of orders) {
        const cs = o.cancelStatus?.cancelState;
        if (cs && cs !== "NONE_REQUESTED") continue;
        if (o.orderPaymentStatus === "FULLY_REFUNDED") continue;
        for (const li of o.lineItems || []) {
          const key = `${o.orderId}|${li.lineItemId}`;
          if (seen.has(key)) continue;
          seen.add(key); newKeys.push(key);
          const qty = li.quantity || 1;
          const unit = qty ? Math.round((parseFloat(li.lineItemCost?.value || 0) / qty) * 100) / 100 : parseFloat(li.lineItemCost?.value || 0);
          const tab = tabFor(o.creationDate);
          (incByTab[tab] = incByTab[tab] || []).push([ddmmyyyy(o.creationDate), li.title || "", qty, unit]);
          inc++;
        }
      }

      // EXPENDITURE (fees + shipping labels) from Finances API
      let exp = 0;
      const txns = await ebayFetchTransactions(access, since);
      for (const t of txns) {
        const tab = tabFor(t.transactionDate);
        if (t.transactionType === "SALE") {
          const fee = feeOfSale(t);
          if (fee > 0) {
            const key = `FEE:${t.transactionId}`;
            if (seen.has(key)) continue;
            seen.add(key); newKeys.push(key);
            (expByTab[tab] = expByTab[tab] || []).push([ddmmyyyy(t.transactionDate), `eBay fees ${t.orderId || ""}`.trim(), "eBay Fees", fee]);
            exp++;
          }
        } else if (t.transactionType === "SHIPPING_LABEL") {
          const amt = Math.abs(Math.round(parseFloat(t.amount?.value || 0) * 100) / 100);
          if (amt > 0) {
            const key = `LBL:${t.transactionId}`;
            if (seen.has(key)) continue;
            seen.add(key); newKeys.push(key);
            (expByTab[tab] = expByTab[tab] || []).push([ddmmyyyy(t.transactionDate), "Postage label", "Postage", amt]);
            exp++;
          }
        }
      }
      console.log(`Account ${acc.idx}: ${orders.length} orders (${inc} new income), ${txns.length} txns (${exp} new expenditure)`);
    } catch (e) {
      console.error(`Account ${acc.idx} failed: ${e.message}`);
    }
  }

  const tabsTouched = new Set([...Object.keys(incByTab), ...Object.keys(expByTab)]);

  let wInc = 0, wExp = 0;
  for (const tab of tabsTouched) await ensureMonthTab(token, tab, sheets);

  for (const tab of Object.keys(incByTab)) {
    const rows = incByTab[tab]; const start = await nextRowIn(token, tab, "A");
    await gfetch(token, `${wsPath(tab)}/range(address='A${start}:D${start + rows.length - 1}')`, { method: "PATCH", body: JSON.stringify({ values: rows }) });
    wInc += rows.length;
  }
  for (const tab of Object.keys(expByTab)) {
    const rows = expByTab[tab]; const start = await nextRowIn(token, tab, "K");
    await gfetch(token, `${wsPath(tab)}/range(address='K${start}:N${start + rows.length - 1}')`, { method: "PATCH", body: JSON.stringify({ values: rows }) });
    wExp += rows.length;
  }

  if (newKeys.length) {
    await gfetch(token, `${wsPath(STATE)}/range(address='A${stateNextRow}:A${stateNextRow + newKeys.length - 1}')`, { method: "PATCH", body: JSON.stringify({ values: newKeys.map((k) => [k]) }) });
  }

  // sort income + expenditure of affected tabs and the current month
  const toSort = new Set(tabsTouched);
  toSort.add(tabFor(new Date().toISOString()));
  for (const tab of toSort) {
    if (!sheets.includes(tab)) continue;
    try { await sortBlock(token, tab, "A", "D"); } catch (e) { console.error(`sort income ${tab}: ${e.message}`); }
    try { await sortBlock(token, tab, "K", "N"); } catch (e) { console.error(`sort exp ${tab}: ${e.message}`); }
  }

  console.log(`Wrote ${wInc} income + ${wExp} expenditure rows; recorded ${newKeys.length} keys.`);
}

main().catch((e) => { console.error(e); process.exit(1); });
