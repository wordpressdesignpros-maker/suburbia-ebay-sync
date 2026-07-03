// diag_postage.js — read-only diagnostic for shipping-label / postage costs.
// Dumps every SHIPPING_LABEL transaction per account and reconciles the two
// summing methods so we can see whether voids/credits inflate the figure.
const crypto = require("crypto");
const ACCOUNT_NAMES = ["superfly", "aqualightingsolutions", "autolightingsolutions", "lightingdepot", "premiumlightingsolutions", "vividlighting"];
const SCOPES = "https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly https://api.ebay.com/oauth/api_scope/sell.finances";
const TZ = "Europe/London";

async function ebayAccessToken(refreshToken) {
  const creds = Buffer.from(`${process.env.EBAY_APP_ID}:${process.env.EBAY_CERT_ID}`).toString("base64");
  const r = await fetch("https://api.ebay.com/identity/v1/oauth2/token", {
    method: "POST", headers: { Authorization: `Basic ${creds}`, "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ grant_type: "refresh_token", refresh_token: refreshToken, scope: SCOPES }),
  });
  if (!r.ok) throw new Error(`token ${r.status} ${await r.text()}`);
  return (await r.json()).access_token;
}
function signHeaders(method, pathOnly, authority) {
  const jwe = process.env.EBAY_SIGNING_KEY;
  const created = Math.floor(Date.now() / 1000);
  const params = `("x-ebay-signature-key" "@method" "@path" "@authority");created=${created}`;
  const base = `"x-ebay-signature-key": ${jwe}\n"@method": ${method}\n"@path": ${pathOnly}\n"@authority": ${authority}\n"@signature-params": ${params}`;
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
    if (!r.ok) throw new Error(`finances ${r.status} ${await r.text()}`);
    const j = await r.json();
    if (j.transactions) txns.push(...j.transactions);
    url = j.next || null;
  }
  return txns;
}
const ukMonth = (iso) => { const f = new Intl.DateTimeFormat("en-GB", { timeZone: TZ, year: "numeric", month: "2-digit" }); const o = {}; for (const p of f.formatToParts(new Date(iso))) if (p.type !== "literal") o[p.type] = p.value; return `${o.year}-${o.month}`; };
const r2 = (n) => Math.round(n * 100) / 100;

async function main() {
  const since = "2026-06-01T00:00:00.000Z";
  const grand = {};
  for (let i = 1; i <= 6; i++) {
    const refresh = process.env[`EBAY_REFRESH_TOKEN_${i}`];
    if (!refresh) { console.log(`\nACC${i} ${ACCOUNT_NAMES[i - 1]}: NO TOKEN`); continue; }
    let txns;
    try { const at = await ebayAccessToken(refresh); txns = await ebayFetchTransactions(at, since); }
    catch (e) { console.log(`\nACC${i} ${ACCOUNT_NAMES[i - 1]}: ERROR ${e.message}`); continue; }

    const byType = {};        // transactionType/bookingEntry -> sum
    const byFeeType = {};     // transactionType/feeType/bookingEntry -> {sum,n} (June only)
    const labels = [];        // shipping-label rows
    for (const t of txns) {
      const amt = parseFloat(t.amount?.value || 0);
      const key = `${t.transactionType}/${t.bookingEntry || "?"}`;
      byType[key] = r2((byType[key] || 0) + amt);
      if (ukMonth(t.transactionDate) === "2026-06") {
        const fk = `${t.transactionType}/${t.feeType || "(none)"}/${t.bookingEntry || "?"}`;
        const e = (byFeeType[fk] ||= { sum: 0, n: 0 });
        e.sum = r2(e.sum + amt); e.n++;
      }
      if (t.transactionType === "SHIPPING_LABEL") {
        labels.push({ m: ukMonth(t.transactionDate), amt, be: t.bookingEntry || "?", ft: t.feeType || "", d: (t.transactionDate || "").slice(0, 10) });
      }
    }
    const sums = {};          // month -> {curAbs, signed, nDebit, nCredit}
    for (const l of labels) {
      const s = (sums[l.m] ||= { curAbs: 0, signed: 0, nDebit: 0, nCredit: 0 });
      s.curAbs = r2(s.curAbs + Math.abs(l.amt));                                  // current sync method
      s.signed = r2(s.signed + (l.be === "CREDIT" ? -Math.abs(l.amt) : Math.abs(l.amt))); // debit - credit
      if (l.be === "CREDIT") s.nCredit++; else s.nDebit++;
    }
    console.log(`\nACC${i} ${ACCOUNT_NAMES[i - 1]}: ${txns.length} txns, ${labels.length} shipping-label rows`);
    console.log(`  byType: ${JSON.stringify(byType)}`);
    console.log(`  June by feeType: ${JSON.stringify(byFeeType)}`);
    console.log(`  postage by month (curAbs=current method, signed=debit-credit): ${JSON.stringify(sums)}`);
    const credits = labels.filter((l) => l.be === "CREDIT");
    if (credits.length) console.log(`  CREDIT (void/refunded) labels: ${JSON.stringify(credits)}`);
    for (const m of Object.keys(sums)) { const g = (grand[m] ||= { curAbs: 0, signed: 0 }); g.curAbs = r2(g.curAbs + sums[m].curAbs); g.signed = r2(g.signed + sums[m].signed); }
  }
  console.log(`\n===== GRAND TOTAL postage by month =====`);
  console.log(JSON.stringify(grand, null, 2));
}
main().catch((e) => { console.error(e); process.exit(1); });
