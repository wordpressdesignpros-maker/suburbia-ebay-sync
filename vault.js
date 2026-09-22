// vault.js — encrypted refresh-token store shared by fetch_ebay.js and rotate_token.js.
//
// Refresh tokens live in tokens/<slot>.enc, encrypted with AES-256-GCM under a
// key derived from the EBAY_SIGNING_PRIVATE secret (only ever present inside
// GitHub Actions). The repo is public, so the ciphertext is useless without
// that secret. A slot with no .enc file falls back to EBAY_REFRESH_TOKEN_<slot>.
const crypto = require("crypto");
const fs = require("fs");
const path = require("path");

const SCOPES = "https://api.ebay.com/oauth/api_scope/sell.fulfillment.readonly https://api.ebay.com/oauth/api_scope/sell.finances";
const RU_NAME = "Simon_Smith-SimonSmi-Suburb-wsoosa";
const TOKENS_DIR = path.join(__dirname, "tokens");
const fileFor = (slot) => path.join(TOKENS_DIR, `${slot}.enc`);

function vaultKey() {
  const secret = process.env.EBAY_SIGNING_PRIVATE;
  if (!secret) throw new Error("EBAY_SIGNING_PRIVATE not set; cannot derive the token vault key");
  return crypto.createHash("sha256").update(`${secret}:ebay-token-vault`).digest();
}

function encrypt(plain) {
  const iv = crypto.randomBytes(12);
  const c = crypto.createCipheriv("aes-256-gcm", vaultKey(), iv);
  const data = Buffer.concat([c.update(plain, "utf8"), c.final()]);
  return { iv: iv.toString("base64"), tag: c.getAuthTag().toString("base64"), data: data.toString("base64") };
}

function decrypt(j) {
  const d = crypto.createDecipheriv("aes-256-gcm", vaultKey(), Buffer.from(j.iv, "base64"));
  d.setAuthTag(Buffer.from(j.tag, "base64"));
  return Buffer.concat([d.update(Buffer.from(j.data, "base64")), d.final()]).toString("utf8");
}

// The refresh token for a slot: the encrypted vault file wins (it is always
// the more recent rotation), otherwise the EBAY_REFRESH_TOKEN_<slot> secret.
function refreshTokenFor(slot) {
  const f = fileFor(slot);
  if (fs.existsSync(f)) {
    try {
      return decrypt(JSON.parse(fs.readFileSync(f, "utf8")));
    } catch (e) {
      console.error(`tokens/${slot}.enc could not be read (${e.message}); falling back to EBAY_REFRESH_TOKEN_${slot}`);
    }
  }
  return process.env[`EBAY_REFRESH_TOKEN_${slot}`];
}

function storeRefreshToken(slot, refreshToken, meta = {}) {
  fs.mkdirSync(TOKENS_DIR, { recursive: true });
  const out = { slot, ...meta, ...encrypt(refreshToken) };
  fs.writeFileSync(fileFor(slot), JSON.stringify(out, null, 2) + "\n");
  return fileFor(slot);
}

module.exports = { SCOPES, RU_NAME, refreshTokenFor, storeRefreshToken, encrypt, decrypt };
