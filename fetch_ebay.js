// fetch_ebay.js  (rebuild-from-live, accurate net figures)
//
// Each run REBUILDS the current month tab from live eBay data, so cancelled /
// refunded orders self-correct. Layout per month tab:
//   INCOME      A:G  DATE / CUSTOMER / POSTCODE / PRODUCT / QUANTITY / UNIT £ / TOTAL £
//   MANUAL EXP  I:L  you fill in (never touched by the sync)
//   SUMMARY     N:O  item sales, refunds, net income, eBay fees, ad fees,
//                    postage, other exp, net profit, margin, orders, AOV
//   BY ACCOUNT  Q:T  per-account income + fees + ad fees
//   BEST SELLERS V:X top products by quantity + sales
// eBay fees, ad fees, postage and refunds are running totals from the ledger.

const crypto = require("crypto");

const ACCOUNTS = [1, 2, 3, 4, 5, 6].map((i) => ({ idx: i }));
// Real eBay shop names for the BY ACCOUNT panel (slots 1-6).
const ACCOUNT_NAMES = ["superfly", "aqualightingsolutions", "autolightingsolutions", "lightingdepot", "premiumlightingsolutions", "vividlighting"];
const MONTHS = ["Jan","Feb","Mar","Apr","May","Jun","Jul","Aug","Sep","Oct","Nov","Dec"];
const SCOPES = "https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly https://api.ebay.com/oauth/api_scope/sell.finances";

// ---------------------------------------------------------------- eBay
async function ebayAccessToken(refreshToken) {
  const creds = Buffer.from(`${process.env.EBAY_APP_ID}:${process.env.EBAY_CERT_ID}`).toString("base64");
  const r = await fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "POST", headers: { Authorization: `Basic ${creds}`, "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: refreshToken, scope: SCOPES }),
  });
  if (!r.ok) throw new Error(`eBay token refresh failed: ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}

async function ebayFetchOrders(accessToken, sinceIso) {
  const orders = [];
  let url = `https://api.ebay.com/sell/fulfillment/v1/order?${new URLSearchParams({ filter: `creationdate:[${sinceIso}..]`, limit: "200" })}`;
  while (url) {
    const r = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json" } });
    if (!r.ok) throw new Error(`eBay order fetch failed: ${r.status} ${await r.text()}`);
    const j = await r.json();
    if (j.orders) orders.push(...j.orders);
    url = j.next || null;
  }
  return orders;
}

function signHeaders(method, pathOnly, authority) {
  const jwe = process.env.EBAY_SIGNING_KEY;
  const created = Math.floor(Date.now() / 1000);
  const params = `("x-ebay-signature-key" "@method" "@path" "@authority");created=${created}`;
  const base =
    `"x-ebay-signature-key": ${jwe}\n"@method": ${method}\n"@path": ${pathOnly}\n"@authority": ${authority}\n"@signature-params": ${params}`;
  const keyObj = crypto.createPrivateKey({ key: Buffer.from(process.env.EBAY_SIGNING_PRIVATE, "base64"), format: "der", type: "pkcs8" });
  const sig = crypto.sign(null, Buffer.from(base, "utf8"), keyObj);
  return { "x-ebay-signature-key": jwe, "Signature-Input": `sig1=${params}`, "Signature": `sig1=:${sig.toString("base64")}:` };
}

async function ebayFetchTransactions(accessToken, sinceIso) {
  const txns = [];
  let url = `https://apiz.ebay.com/sell/finances/v1/transaction?${new URLSearchParams({ filter: `transactionDate:[${sinceIso}..]`, limit: "200" })}`;
  while (url) {
    const u = new URL(url);
    const r = await fetch(url, { headers: { Authorization: `Bearer ${accessToken}`, "Content-Type": "application/json", ...signHeaders("GET", u.pathname, u.host) } });
    if (!r.ok) throw new Error(`eBay finances fetch failed: ${r.status} ${await r.text()}`);
    const j = await r.json();
    if (j.transactions) txns.push(...j.transactions);
    url = j.next || null;
  }
  return txns;
}

function feeAmount(t) {
  if (t.totalFeeAmount && t.totalFeeAmount.value) return Math.abs(parseFloat(t.totalFeeAmount.value));
  let f = 0;
  for (const li of t.orderLineItems || []) for (const mf of li.marketplaceFees || []) f += Math.abs(parseFloat(mf.amount?.value || 0));
  return f;
}

// promoted listings / advertising fees vs ordinary selling fees
const isAdFee = (ft) => /AD[_ ]?FEE|ADVERT|PROMOT/i.test(String(ft || ""));
function splitFees(t) {
  let sell = 0, ad = 0, hadDetail = false;
  for (const li of t.orderLineItems || []) for (const mf of li.marketplaceFees || []) {
    hadDetail = true;
    const amt = Math.abs(parseFloat(mf.amount?.value || 0));
    if (isAdFee(mf.feeType)) ad += amt; else sell += amt;
  }
  if (!hadDetail) {
    const amt = Math.abs(parseFloat(t.totalFeeAmount?.value || 0));
    if (isAdFee(t.feeType)) ad += amt; else sell += amt;
  }
  return { sell, ad };
}

// ---------------------------------------------------------------- Graph
async function msToken() {
  const tenant = process.env.MS_TENANT_ID || "consumers";
  const p = { client_id: process.env.MS_CLIENT_ID, grant_type: "refresh_token", refresh_token: process.env.MS_REFRESH_TOKEN, scope: "Files.ReadWrite offline_access" };
  if (process.env.MS_CLIENT_SECRET) p.client_secret = process.env.MS_CLIENT_SECRET;
  const r = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, { method: "POST", headers: { "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams(p) });
  if (!r.ok) throw new Error(`Microsoft token refresh failed: ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}
const FID = () => process.env.ONEDRIVE_FILE_ID;
const WB = () => `https://graph.microsoft.com/v1.0/me/drive/items/${FID()}/workbook`;
const wsPath = (n) => `${WB()}/worksheets('${encodeURIComponent(n)}')`;
async function gfetch(token, url, opts = {}, attempt = 1) {
  const r = await fetch(url, { ...opts, headers: { Authorization: `Bearer ${token}`, "Content-Type": "application/json", ...(opts.headers || {}) } });
  if (!r.ok) {
    const body = (await r.text()).slice(0, 200);
    // Transient Graph errors (503/500/429/locked) — back off and retry so a
    // blip between a clear and a write can't leave a section blank.
    const transient = r.status >= 500 || r.status === 429 || r.status === 423;
    if (transient && attempt < 6) {
      await new Promise((res) => setTimeout(res, 1500 * attempt));
      return gfetch(token, url, opts, attempt + 1);
    }
    throw new Error(`Graph ${opts.method || "GET"} -> ${r.status} ${body}`);
  }
  return r.status === 204 ? null : r.json();
}
const listSheets = async (token) => ((await gfetch(token, `${WB()}/worksheets?$select=name`)).value || []).map((w) => w.name);
const patch = (token, tab, addr, values) => gfetch(token, `${wsPath(tab)}/range(address='${addr}')`, { method: "PATCH", body: JSON.stringify({ values }) });
const clearRange = (token, tab, addr) => gfetch(token, `${wsPath(tab)}/range(address='${addr}')/clear`, { method: "POST", body: JSON.stringify({ applyTo: "Contents" }) });

async function ensureMonthTab(token, name, existing) {
  if (existing.includes(name)) return;
  await gfetch(token, `${WB()}/worksheets/add`, { method: "POST", body: JSON.stringify({ name }) });
  await patch(token, name, "A2", [["INCOME"]]);
  await patch(token, name, "I2", [["OTHER EXPENDITURE (you fill in)"]]);
  await patch(token, name, "N2", [["SUMMARY"]]);
  await patch(token, name, "Q2", [["BY ACCOUNT"]]);
  await patch(token, name, "V2", [["BEST SELLERS"]]);
  await patch(token, name, "A4:G4", [["DATE", "CUSTOMER", "POSTCODE", "PRODUCT", "QUANTITY", "UNIT £", "TOTAL £"]]);
  await patch(token, name, "I4:L4", [["DATE", "DESCRIPTION", "CATEGORY", "£"]]);
  await patch(token, name, "Q4:T4", [["ACCOUNT", "INCOME £", "FEES £", "AD FEES £"]]);
  await patch(token, name, "V4:X4", [["PRODUCT", "QTY", "SALES £"]]);
  await patch(token, name, "N4:O14", [
    ["Item sales", "=SUM($G$6:$G$100000)"], ["Refunds (deducted)", 0], ["Net income", "=O4-O5"],
    ["eBay fees", 0], ["Promoted/ad fees", 0], ["Postage", 0], ["Other expenditure", "=SUM($L$6:$L$100000)"],
    ["Net profit", "=O6-O7-O8-O9-O10"], ["Profit margin", "=IF(O6=0,0,O11/O6)"], ["Orders", 0], ["Avg order value", "=IF(O13=0,0,O6/O13)"],
  ]);
  await patch(token, name, "Q5:Q10", ACCOUNT_NAMES.map((n) => [n]));
  existing.push(name);
}

const ddmmyyyy = (iso) => { const d = new Date(iso); const p = (n) => String(n).padStart(2, "0"); return `${p(d.getUTCDate())}.${p(d.getUTCMonth() + 1)}.${d.getUTCFullYear()}`; };
const tabFor = (iso) => { const d = new Date(iso); return `${MONTHS[d.getUTCMonth()]} ${d.getUTCFullYear()}`; };
const dateKey = (s) => { const [d, m, y] = String(s).split("."); return (+y) * 10000 + (+m) * 100 + (+d); };
const round2 = (n) => Math.round(n * 100) / 100;
const monthStartIso = () => { const n = new Date(); return new Date(Date.UTC(n.getUTCFullYear(), n.getUTCMonth(), 1)).toISOString(); };

// ---------------------------------------------------------------- main
async function main() {
  if (!FID()) throw new Error("ONEDRIVE_FILE_ID not set");
  const token = await msToken();
  const sheets = await listSheets(token);
  const since = monthStartIso();
  const tab = tabFor(new Date().toISOString());

  const income = [];
  const perAcc = {};
  let refunds = 0, fees = 0, ads = 0, feeCredits = 0, postage = 0;
  const orderIds = new Set();

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));
  const failed = [];

  for (const acc of ACCOUNTS) {
    const refresh = process.env[`EBAY_REFRESH_TOKEN_${acc.idx}`];
    perAcc[acc.idx] = { income: 0, fees: 0, ads: 0, feeCredits: 0 };
    if (!refresh) { failed.push(acc.idx); continue; }

    let ok = false, lastErr;
    for (let attempt = 1; attempt <= 3 && !ok; attempt++) {
      try {
        const access = await ebayAccessToken(refresh);
        const orders = await ebayFetchOrders(access, since);
        const txns = await ebayFetchTransactions(access, since);

        // accumulate into locals first so a retry can never double-count
        const localIncome = [], localIds = new Set();
        let aInc = 0, aFees = 0, aAds = 0, aCred = 0, aRef = 0, aPost = 0;
        for (const o of orders) {
          const cs = o.cancelStatus?.cancelState;
          if (cs && cs !== "NONE_REQUESTED") continue;
          localIds.add(o.orderId);
          const ship = o.fulfillmentStartInstructions?.[0]?.shippingStep?.shipTo;
          const customer = ship?.fullName || "";
          const postcode = ship?.contactAddress?.postalCode || "";
          for (const li of o.lineItems || []) {
            const qty = li.quantity || 1;
            const lineCost = parseFloat(li.lineItemCost?.value || 0);
            const unit = qty ? round2(lineCost / qty) : lineCost;
            localIncome.push([ddmmyyyy(o.creationDate), customer, postcode, li.title || "", qty, unit, round2(lineCost)]);
            aInc += lineCost;
          }
        }
        for (const t of txns) {
          if (t.transactionType === "SALE") { const s = splitFees(t); aFees += s.sell; aAds += s.ad; }
          else if (t.transactionType === "NON_SALE_CHARGE") { const amt = Math.abs(parseFloat(t.amount?.value || 0)); if (isAdFee(t.feeType)) aAds += amt; else aFees += amt; }
          else if (t.transactionType === "SHIPPING_LABEL") aPost += Math.abs(parseFloat(t.amount?.value || 0));
          else if (t.transactionType === "REFUND") { aRef += Math.abs(parseFloat(t.amount?.value || 0)); aCred += feeAmount(t); }
        }

        income.push(...localIncome);
        for (const id of localIds) orderIds.add(id);
        perAcc[acc.idx] = { income: aInc, fees: aFees, ads: aAds, feeCredits: aCred };
        fees += aFees; ads += aAds; feeCredits += aCred; refunds += aRef; postage += aPost;
        ok = true;
        console.log(`Account ${acc.idx}: ${orders.length} orders, ${txns.length} txns`);
      } catch (e) {
        lastErr = e;
        if (attempt < 3) await sleep(2500 * attempt);
      }
    }
    if (!ok) { console.error(`Account ${acc.idx} failed after 3 attempts: ${lastErr && lastErr.message}`); failed.push(acc.idx); }
  }

  // Never overwrite good data with a partial result.
  if (failed.length) {
    throw new Error(`Aborting write — accounts [${failed.join(", ")}] could not be fetched this run; keeping the last complete figures.`);
  }

  income.sort((a, b) => dateKey(b[0]) - dateKey(a[0])); // newest first (top), oldest last
  fees = round2(fees - feeCredits); ads = round2(ads); refunds = round2(refunds); postage = round2(postage);

  // best sellers: aggregate by product title across all accounts
  const prodMap = new Map();
  for (const r of income) {
    const title = r[3] || "(no title)";
    const cur = prodMap.get(title) || { qty: 0, sales: 0 };
    cur.qty += Number(r[4]) || 0;
    cur.sales += Number(r[6]) || 0;
    prodMap.set(title, cur);
  }
  const best = [...prodMap.entries()]
    .sort((a, b) => b[1].qty - a[1].qty || b[1].sales - a[1].sales)
    .slice(0, 15)
    .map(([title, v]) => [title, v.qty, round2(v.sales)]);

  await ensureMonthTab(token, tab, sheets);
  // Write rows FIRST, then clear only the rows below the new data. A failure
  // can therefore never leave a populated section blank.
  if (income.length) await patch(token, tab, `A6:G${5 + income.length}`, income);
  await clearRange(token, tab, `A${6 + income.length}:G100000`);
  await patch(token, tab, "O5", [[refunds]]);
  await patch(token, tab, "O7", [[fees]]);
  await patch(token, tab, "O8", [[ads]]);
  await patch(token, tab, "O9", [[postage]]);
  await patch(token, tab, "O13", [[orderIds.size]]);
  // per-account rows: [name, income, fees, adFees] — sorted by highest income
  const accCombined = [];
  for (let i = 1; i <= 6; i++) {
    const p = perAcc[i] || { income: 0, fees: 0, ads: 0, feeCredits: 0 };
    accCombined.push([ACCOUNT_NAMES[i - 1], round2(p.income), round2(p.fees - p.feeCredits), round2(p.ads)]);
  }
  accCombined.sort((a, b) => b[1] - a[1]);
  await patch(token, tab, "Q5:Q10", accCombined.map((r) => [r[0]]));
  await patch(token, tab, "R5:T10", accCombined.map((r) => [r[1], r[2], r[3]]));
  // TOTAL line under the table (income, fees, ad fees)
  const tot = (i) => round2(accCombined.reduce((s, r) => s + r[i], 0));
  await patch(token, tab, "Q11:T11", [["TOTAL", tot(1), tot(2), tot(3)]]);
  try {
    await gfetch(token, `${wsPath(tab)}/range(address='Q1:Q11')/format`, { method: "PATCH", body: JSON.stringify({ columnWidth: 150 }) });
    await gfetch(token, `${wsPath(tab)}/range(address='R11:T11')`, { method: "PATCH", body: JSON.stringify({ numberFormat: [["£#,##0.00", "£#,##0.00", "£#,##0.00"]] }) });
    await gfetch(token, `${wsPath(tab)}/range(address='Q11:T11')/format/font`, { method: "PATCH", body: JSON.stringify({ bold: true }) });
  } catch (_) {}
  if (best.length) await patch(token, tab, `V6:X${5 + best.length}`, best);
  await clearRange(token, tab, `V${6 + best.length}:X100`);

  console.log(`${tab}: ${income.length} income rows, ${orderIds.size} orders, fees £${fees}, ads £${ads}, postage £${postage}, refunds £${refunds}`);
}

main().catch((e) => { console.error(e); process.exit(1); });
