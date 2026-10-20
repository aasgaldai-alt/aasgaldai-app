// @ts-nocheck
// Web Push (RFC 8291 aes128gcm + RFC 8292 VAPID) — WebCrypto дээр, гадны сангүй.
import { Buffer } from "node:buffer";
const subtle = globalThis.crypto.subtle;
const te = new TextEncoder();
const b64u = (b) => Buffer.from(b).toString("base64url");
const fromB64u = (s) => new Uint8Array(Buffer.from(String(s), "base64url"));
function concat(...arrs) {
  const n = arrs.reduce((a, x) => a + x.length, 0), out = new Uint8Array(n);
  let o = 0;
  for (const a of arrs) { out.set(a, o); o += a.length; }
  return out;
}

export async function vapidGenerate() {
  const kp = await subtle.generateKey({ name: "ECDSA", namedCurve: "P-256" }, true, ["sign", "verify"]);
  const jwk = await subtle.exportKey("jwk", kp.privateKey);
  const pub = new Uint8Array(await subtle.exportKey("raw", kp.publicKey));
  return { jwk, pub: b64u(pub) };
}

async function vapidHeader(endpoint, vapid, subject) {
  const aud = new URL(endpoint).origin;
  const h = b64u(te.encode(JSON.stringify({ typ: "JWT", alg: "ES256" })));
  const p = b64u(te.encode(JSON.stringify({ aud, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: subject })));
  const key = await subtle.importKey("jwk", vapid.jwk, { name: "ECDSA", namedCurve: "P-256" }, false, ["sign"]);
  const sig = new Uint8Array(await subtle.sign({ name: "ECDSA", hash: "SHA-256" }, key, te.encode(h + "." + p)));
  return "vapid t=" + h + "." + p + "." + b64u(sig) + ", k=" + vapid.pub;
}

async function hkdf(salt, ikm, info, len) {
  const key = await subtle.importKey("raw", ikm, "HKDF", false, ["deriveBits"]);
  return new Uint8Array(await subtle.deriveBits({ name: "HKDF", hash: "SHA-256", salt, info }, key, len * 8));
}

export async function encryptPayload(sub, plain, saltOpt, asKeysOpt) {
  const uaPub = fromB64u(sub.p256dh), authSecret = fromB64u(sub.auth);
  const as = asKeysOpt || (await subtle.generateKey({ name: "ECDH", namedCurve: "P-256" }, true, ["deriveBits"]));
  const asPub = new Uint8Array(await subtle.exportKey("raw", as.publicKey));
  const uaKey = await subtle.importKey("raw", uaPub, { name: "ECDH", namedCurve: "P-256" }, false, []);
  const shared = new Uint8Array(await subtle.deriveBits({ name: "ECDH", public: uaKey }, as.privateKey, 256));
  const ikm = await hkdf(authSecret, shared, concat(te.encode("WebPush: info\0"), uaPub, asPub), 32);
  const salt = saltOpt || globalThis.crypto.getRandomValues(new Uint8Array(16));
  const cek = await hkdf(salt, ikm, te.encode("Content-Encoding: aes128gcm\0"), 16);
  const nonce = await hkdf(salt, ikm, te.encode("Content-Encoding: nonce\0"), 12);
  const k = await subtle.importKey("raw", cek, "AES-GCM", false, ["encrypt"]);
  const ct = new Uint8Array(await subtle.encrypt({ name: "AES-GCM", iv: nonce }, k, concat(plain, new Uint8Array([2]))));
  const head = new Uint8Array(21 + asPub.length);
  head.set(salt, 0);
  new DataView(head.buffer).setUint32(16, 4096);
  head[20] = asPub.length;
  head.set(asPub, 21);
  return concat(head, ct);
}

// Нэг төхөөрөмж рүү илгээнэ. Буцаах: HTTP статус (201 = амжилттай, 404/410 = захиалга хүчингүй)
export async function sendPush(sub, vapid, data, subject) {
  const body = await encryptPayload(sub, te.encode(JSON.stringify(data)));
  const r = await fetch(sub.endpoint, {
    method: "POST",
    headers: { Authorization: await vapidHeader(sub.endpoint, vapid, subject), "Content-Encoding": "aes128gcm", "Content-Type": "application/octet-stream", TTL: "86400", Urgency: "high" },
    body,
  });
  let text = "";
  try { text = (await r.text()).slice(0, 200); } catch (_) {}
  return { status: r.status, text };
}
