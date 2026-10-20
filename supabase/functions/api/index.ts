// @ts-nocheck
// Аасгалдай API — Supabase Edge Function. Өгөгдөл: Postgres (sheets хүснэгт), зураг: Storage (files bucket).
import postgres from "npm:postgres@3.4.5";
import { runRequest, encRows, decRows } from "./core.js";

const sql = postgres(Deno.env.get("SUPABASE_DB_URL"), { prepare: false, max: 4, idle_timeout: 30, connect_timeout: 10 });
const SB_URL = Deno.env.get("SUPABASE_URL");
function serviceKey() {
  const k = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY");
  if (k) return k;
  try { const j = JSON.parse(Deno.env.get("SUPABASE_SECRET_KEYS") || "{}"); return j.default || Object.values(j)[0]; } catch (_) { return ""; }
}
const SKEY = serviceKey();
const authH = () => (SKEY.startsWith("eyJ") ? { apikey: SKEY, Authorization: "Bearer " + SKEY } : { apikey: SKEY, Authorization: "Bearer " + SKEY });

const LEGACY_URL = "https://script.google.com/macros/s/AKfycbxEbho1pB6UOulQAeMfHWw0AzPe8M6qC3YsiiwglUrc8hzZuwPBBF0SsUtVQs0v8b4omQ/exec";
class Conflict extends Error {}
const cache = new Map(); // name -> {rows, ver}  (instance дотор хадгалагдах кэш)

const store = {
  async snapshot() {
    const vers = await sql`select name, ver from sheets`;
    const live = new Set(), need = [];
    for (const v of vers) {
      live.add(v.name);
      const c = cache.get(v.name);
      if (!c || c.ver !== Number(v.ver)) need.push(v.name);
    }
    if (need.length) {
      const rows = await sql`select name, ver, data::text as data from sheets where name = any(${need})`;
      for (const r of rows) cache.set(r.name, { ver: Number(r.ver), rows: decRows(JSON.parse(r.data)) });
    }
    for (const k of [...cache.keys()]) if (!live.has(k)) cache.delete(k);
    return Object.fromEntries(cache);
  },
  async commit(list) {
    try {
      const vers = await sql.begin(async (tx) => {
        const res = [];
        for (const it of list) {
          const txt = JSON.stringify(encRows(it.rows));
          const r = it.baseVer === 0
            ? await tx`insert into sheets (name, data, ver) values (${it.name}, ${txt}::jsonb, 1) on conflict (name) do nothing returning ver`
            : await tx`update sheets set data = ${txt}::jsonb, ver = ver + 1, updated_at = now() where name = ${it.name} and ver = ${it.baseVer} returning ver`;
          if (!r.length) throw new Conflict("conflict");
          res.push(Number(r[0].ver));
        }
        return res;
      });
      list.forEach((it, i) => cache.set(it.name, { rows: it.rows, ver: vers[i] }));
      return true;
    } catch (e) {
      if (e instanceof Conflict) { for (const it of list) cache.delete(it.name); return false; }
      throw e;
    }
  },
  async replaceAll(sheets) {
    await sql.begin(async (tx) => {
      const names = Object.keys(sheets);
      await tx`delete from sheets where not (name = any(${names}))`;
      for (const n of names) {
        const txt = JSON.stringify(encRows(sheets[n]));
        await tx`insert into sheets (name, data, ver) values (${n}, ${txt}::jsonb, 1)
                 on conflict (name) do update set data = excluded.data, ver = sheets.ver + 1, updated_at = now()`;
      }
    });
    cache.clear();
  },
  async upload(id, mime, b64) {
    const bin = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0));
    const r = await fetch(`${SB_URL}/storage/v1/object/files/${encodeURIComponent(id)}`, {
      method: "POST", headers: { ...authH(), "Content-Type": mime, "x-upsert": "true", "cache-control": "31536000" }, body: bin,
    });
    if (!r.ok) throw new Error("Зураг хадгалж чадсангүй (" + r.status + "): " + (await r.text()).slice(0, 200));
  },
  async download(id) {
    const r = await fetch(`${SB_URL}/storage/v1/object/files/${encodeURIComponent(id)}`, { headers: authH() });
    if (!r.ok) { await r.body?.cancel(); return null; }
    return { mime: r.headers.get("content-type") || "image/jpeg", bytes: new Uint8Array(await r.arrayBuffer()) };
  },
  async legacy(body) {
    const r = await fetch(LEGACY_URL, { method: "POST", headers: { "Content-Type": "text/plain;charset=utf-8" }, body: JSON.stringify(body), redirect: "follow" });
    const t = await r.text();
    try { return JSON.parse(t); } catch (_) { return { error: "server", message: "Хуучин сервер: " + r.status }; }
  },
  async listFiles() {
    const out = [];
    for (let off = 0; ; off += 1000) {
      const r = await fetch(`${SB_URL}/storage/v1/object/list/files`, {
        method: "POST", headers: { ...authH(), "Content-Type": "application/json" },
        body: JSON.stringify({ prefix: "", limit: 1000, offset: off }),
      });
      if (!r.ok) throw new Error("list " + r.status + " " + (await r.text()).slice(0, 200));
      const j = await r.json();
      j.forEach((x) => out.push(x.name));
      if (j.length < 1000) break;
    }
    return out;
  },
};

const CORS = { "Access-Control-Allow-Origin": "*", "Access-Control-Allow-Headers": "*", "Access-Control-Allow-Methods": "POST, GET, OPTIONS" };
const json = (o) => new Response(JSON.stringify(o), { headers: { ...CORS, "Content-Type": "application/json; charset=utf-8" } });

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: CORS });
  if (req.method !== "POST") return json({ ok: true, service: "aasgaldai-api" });
  let body;
  try { body = JSON.parse(await req.text()); } catch (_) { return json({ error: "bad", message: "JSON буруу" }); }
  try { return json(await runRequest(body, store)); }
  catch (e) { console.error(e); return json({ error: "server", message: String(e && e.message || e) }); }
});
