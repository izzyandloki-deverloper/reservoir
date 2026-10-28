// Reservoir-Worker-2026.10.28-w169
globalThis.w169_TOP = "Reservoir-Worker-2026.10.28-w169";
// Reservoir API Worker. Validates the Microsoft Entra token on every request, then reads and writes the D1 fleet table.
const FIELDS = ["name","rego","expiry","serial","make","model","colour","year","body","vin","engine","compliance","fuel","driver","garaged","finance","directDebit","contract","lastService","nextService","serviceProvider","servicePhone","serviceEmail","serviceBooking","financeExpiry","financeDocs"];

let jwksCache = { at: 0, keys: null };

const b64u = (s) => Uint8Array.from(atob(s.replace(/-/g, "+").replace(/_/g, "/").padEnd(Math.ceil(s.length / 4) * 4, "=")), (c) => c.charCodeAt(0));
const parseJson = (s) => JSON.parse(new TextDecoder().decode(b64u(s)));

async function getKeys(env) {
  if (jwksCache.keys && Date.now() - jwksCache.at < 3600000) return jwksCache.keys;
  const res = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/discovery/v2.0/keys`);
  if (!res.ok) throw new Error("jwks");
  jwksCache = { at: Date.now(), keys: (await res.json()).keys };
  return jwksCache.keys;
}

async function verifyToken(request, env) {
  const h = request.headers.get("Authorization") || "";
  if (!h.startsWith("Bearer ")) throw new Error("no token");
  const token = h.slice(7);
  const parts = token.split(".");
  if (parts.length !== 3) throw new Error("bad token");
  const header = parseJson(parts[0]);
  const claims = parseJson(parts[1]);
  if (header.alg !== "RS256") throw new Error("alg");
  let keys = await getKeys(env);
  let jwk = keys.find((k) => k.kid === header.kid);
  if (!jwk) { jwksCache.keys = null; keys = await getKeys(env); jwk = keys.find((k) => k.kid === header.kid); }
  if (!jwk) throw new Error("kid");
  const key = await crypto.subtle.importKey("jwk", jwk, { name: "RSASSA-PKCS1-v1_5", hash: "SHA-256" }, false, ["verify"]);
  const ok = await crypto.subtle.verify("RSASSA-PKCS1-v1_5", key, b64u(parts[2]), new TextEncoder().encode(parts[0] + "." + parts[1]));
  if (!ok) throw new Error("signature");
  const now = Math.floor(Date.now() / 1000);
  if (claims.exp < now - 60 || (claims.nbf && claims.nbf > now + 60)) throw new Error("expired");
  if (claims.tid !== env.TENANT_ID) throw new Error("tenant");
  if (claims.aud !== env.CLIENT_ID && claims.aud !== `api://${env.CLIENT_ID}`) throw new Error("aud");
  const issOk = claims.iss === `https://login.microsoftonline.com/${env.TENANT_ID}/v2.0` || claims.iss === `https://sts.windows.net/${env.TENANT_ID}/`;
  if (!issOk) throw new Error("iss");
  const roles = Array.isArray(claims.roles) ? claims.roles : [];
  const tier = ["T1", "T2", "T3"].find((t) => roles.includes(t));
  if (!tier) throw new Error("no role");
  return { tier, name: claims.name || claims.preferred_username || "", email: String(claims.email || claims.preferred_username || claims.upn || "").toLowerCase() };
}

const cors = (env) => ({
  "Access-Control-Allow-Origin": env.ALLOWED_ORIGIN,
  "Access-Control-Allow-Headers": "Authorization, Content-Type, X-View-As",
  "Access-Control-Allow-Methods": "GET, PUT, POST, DELETE, OPTIONS",
  "Access-Control-Max-Age": "86400",
  Vary: "Origin"
});
const reply = (env, status, body) => new Response(body === null ? null : JSON.stringify(body), { status, headers: { ...cors(env), "Content-Type": "application/json", "Cache-Control": "no-store" } });

// A save sends the whole vehicle, so a phone holding an old copy used to wipe whatever another device had just added. The app now sends the ids it had loaded (known). Anything on the server that the sender never had is kept, anything the sender had and removed is dropped, and anything it sent wins.
const mergeById = (incoming, storedJson, known) => {
  let stored;
  try { stored = JSON.parse(storedJson || "[]"); } catch { stored = []; }
  if (!Array.isArray(known) || !Array.isArray(stored)) return incoming;
  const have = new Set(incoming.map((x) => x && x.id));
  const seen = new Set(known.map(String));
  const keep = stored.filter((x) => x && x.id !== undefined && !have.has(x.id) && !seen.has(String(x.id)));
  return [...incoming, ...keep];
};

const fleetLimit = (f) => f === "financeDocs" ? 20000 : 500;
let fleetColsDone = false;
async function fleetCols(env) {
  if (fleetColsDone) return;
  for (const c of ["financeExpiry", "financeDocs"]) { try { await env.DB.prepare(`ALTER TABLE fleet ADD COLUMN ${c} TEXT`).run(); } catch (e) { /* already there */ } }
  fleetColsDone = true;
}

const rowToItem = (r) => {
  const item = { id: r.id, group: r.grp };
  FIELDS.forEach((f) => { item[f] = r[f] ?? ""; });
  try { item.repairs = JSON.parse(r.repairs || "[]"); } catch { item.repairs = []; }
  try { item.serviceHistory = JSON.parse(r.serviceHistory || "[]"); } catch { item.serviceHistory = []; }
  return item;
};

const LABELS = { name: "Name", rego: "Rego", expiry: "Rego Expiry", serial: "Registration Serial Number", make: "Make", model: "Model", colour: "Colour", year: "Year", body: "Body Type", vin: "VIN", engine: "Engine Number", compliance: "Compliance Plate", fuel: "Fuel", driver: "Driver", garaged: "Garaged Address", finance: "Finance Company", directDebit: "Direct Debit", contract: "Contract Number", serviceProvider: "Service Provider", servicePhone: "Service Phone", serviceEmail: "Service Email", serviceBooking: "Service Booking Link", financeExpiry: "Finance Expiry", financeDocs: "Finance Documents" };


// ---- SharePoint document storage through Microsoft Graph (Reservoir site, app only access) ----
const GRAPH = "https://graph.microsoft.com/v1.0";
const GRAPH_CLIENT_ID_DEFAULT = "9748df41-dbe2-47ce-ba3d-34bafa8670a7";
const GRAPH_SITE_ID_DEFAULT = "directflowplumbing.sharepoint.com,1f11da18-92c7-4490-aa35-79124b6fa696,3857ff5c-4f9e-491a-bd93-959c29dd3fac";
const MAX_UPLOAD = 25 * 1024 * 1024;
const CHUNK = 320 * 1024 * 10;
let graphToken = { at: 0, value: "" };

async function getGraphToken(env) {
  if (graphToken.value && Date.now() < graphToken.at) return graphToken.value;
  if (!env.GRAPH_CLIENT_SECRET) throw new Error("graph secret missing");
  const body = new URLSearchParams({ client_id: env.GRAPH_CLIENT_ID || GRAPH_CLIENT_ID_DEFAULT, client_secret: env.GRAPH_CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" });
  const res = await fetch(`https://login.microsoftonline.com/${env.TENANT_ID}/oauth2/v2.0/token`, { method: "POST", body });
  if (!res.ok) throw new Error("graph token " + res.status);
  const j = await res.json();
  graphToken = { value: j.access_token, at: Date.now() + (j.expires_in - 300) * 1000 };
  return graphToken.value;
}
const siteId = (env) => env.GRAPH_SITE_ID || GRAPH_SITE_ID_DEFAULT;
const cleanName = (s) => String(s || "").replace(/[\\/:*?"<>|#%~&{}]/g, "-").replace(/\s+/g, " ").trim().replace(/^\.+/, "").slice(0, 120);
const pathOf = (segs) => segs.map((x) => encodeURIComponent(x)).join("/");

async function graphFetch(env, url, init = {}) {
  const token = await getGraphToken(env);
  return fetch(url.startsWith("http") ? url : GRAPH + url, { ...init, headers: { Authorization: `Bearer ${token}`, ...(init.headers || {}) } });
}

async function uploadToSharePoint(env, segs, fileName, bytes, contentType, replace) {
  const behaviour = replace ? "replace" : "rename";
  const base = `/sites/${siteId(env)}/drive/root:/${pathOf([...segs, fileName])}:`;
  if (bytes.byteLength <= 4 * 1024 * 1024) {
    const res = await graphFetch(env, `${base}/content?@microsoft.graph.conflictBehavior=${behaviour}`, { method: "PUT", headers: { "Content-Type": contentType || "application/octet-stream" }, body: bytes });
    if (!res.ok) throw new Error("upload " + res.status);
    return res.json();
  }
  const sess = await graphFetch(env, `${base}/createUploadSession`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": behaviour } }) });
  if (!sess.ok) throw new Error("session " + sess.status);
  const { uploadUrl } = await sess.json();
  let last;
  for (let start = 0; start < bytes.byteLength; start += CHUNK) {
    const end = Math.min(start + CHUNK, bytes.byteLength);
    const res = await fetch(uploadUrl, { method: "PUT", headers: { "Content-Length": String(end - start), "Content-Range": `bytes ${start}-${end - 1}/${bytes.byteLength}` }, body: bytes.slice(start, end) });
    if (!res.ok && res.status !== 202) throw new Error("chunk " + res.status);
    last = res;
  }
  return last.json();
}

// ---- Xero bills by email (Microsoft Graph sendMail from a DFP mailbox) ----
const MAIL_SENDER_DEFAULT = "accounts@directflow.com.au";
const MAX_EMAIL = 3.5 * 1024 * 1024;
const toB64 = (buf) => {
  const b = new Uint8Array(buf);
  let out = "";
  for (let i = 0; i < b.length; i += 0x8000) out += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000));
  return btoa(out);
};

// Reads a stored file's details and which folder it sits in. Only files inside Fleet Management or Receipts are ever touched.
async function driveMeta(env, id) {
  const r = await graphFetch(env, `/sites/${siteId(env)}/drive/items/${id}?select=id,name,file,parentReference`);
  if (!r.ok) return null;
  const m = await r.json();
  if (!m.file) return null;
  let path = (m.parentReference && m.parentReference.path) || "";
  try { path = decodeURIComponent(path); } catch (e) { /* keep as is */ }
  const i = path.indexOf("root:/");
  m.folder = i >= 0 ? path.slice(i + 6) : "";
  if (!(m.folder.startsWith("Fleet Management/") || m.folder.startsWith("Receipts/"))) return null;
  return m;
}

async function driveBytes(env, id) {
  const dl = await graphFetch(env, `/sites/${siteId(env)}/drive/items/${id}/content`, { redirect: "manual" });
  const loc = dl.headers.get("Location");
  const file = loc ? await fetch(loc) : dl;
  if (!file.ok) throw new Error("download " + file.status);
  return file.arrayBuffer();
}

async function sendToXero(env, user, meta) {
  if (!env.XERO_BILLS_EMAIL) throw new Error("XERO_BILLS_EMAIL is not set on the Worker");
  const bytes = await driveBytes(env, meta.id);
  if (bytes.byteLength > MAX_EMAIL) throw new Error("too large");
  const sender = env.MAIL_SENDER || MAIL_SENDER_DEFAULT;
  const payload = JSON.stringify({
    message: {
      subject: meta.name,
      body: { contentType: "Text", content: `Sent from Reservoir by ${user.name}.` },
      toRecipients: [{ emailAddress: { address: env.XERO_BILLS_EMAIL } }],
      attachments: [{ "@odata.type": "#microsoft.graph.fileAttachment", name: meta.name, contentType: (meta.file && meta.file.mimeType) || "application/pdf", contentBytes: toB64(bytes) }]
    },
    saveToSentItems: true
  });
  const post = () => graphFetch(env, `/users/${encodeURIComponent(sender)}/sendMail`, { method: "POST", headers: { "Content-Type": "application/json" }, body: payload });
  let res = await post();
  if (res.status === 403 || res.status === 401) {
    // The cached app token may pre-date a newly granted permission. Fetch a fresh one and try once more.
    graphToken = { at: 0, value: "" };
    res = await post();
  }
  if (!res.ok) {
    let code = "";
    try { code = (await res.json()).error.code || ""; } catch (e) { /* no body */ }
    let roles = "";
    try { roles = (parseJson(graphToken.value.split(".")[1]).roles || []).join(" "); } catch (e) { /* ignore */ }
    throw new Error(`sendMail ${res.status} ${code} (app permissions: ${roles || "none"})`.trim());
  }
}

let receiptRecordsReady = false;
async function receiptRecordsTable(env) {
  if (receiptRecordsReady) return;
  receiptRecordsReady = true;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS receipt_records (id TEXT PRIMARY KEY, by_name TEXT, by_email TEXT, type TEXT, amount TEXT, note TEXT, item_id TEXT, file_name TEXT, status TEXT, view TEXT, sent_date TEXT, done_date TEXT, created_at TEXT DEFAULT (datetime('now')))").run();
}
const receiptRecordOut = (r) => ({ id: r.id, kind: "receipt", receiptType: r.type, amount: r.amount || "", text: r.note || "", itemId: r.item_id || "", fileName: r.file_name || "", status: r.status, by: r.by_name || "", date: r.sent_date || String(r.created_at || "").slice(0, 10), doneDate: r.done_date || "", photos: r.view ? [{ name: r.file_name || "", view: r.view }] : [], server: true });

async function logUpdate(env, user, vehicleId, label, summary, detail) {
  await env.DB.prepare("INSERT INTO t2_updates (user_name, tier, vehicle_id, vehicle_label, summary, detail) VALUES (?, ?, ?, ?, ?, ?)").bind(user.name, user.tier, vehicleId, String(label || "").slice(0, 100), String(summary).slice(0, 200), JSON.stringify(detail || {}).slice(0, 4000)).run();
}

const WORKER_BUILD = "Reservoir-Worker-2026.10.28-w169";
// ---- Xero connection (OAuth2 web app). Secrets: XERO_CLIENT_SECRET. Optional vars: XERO_CLIENT_ID, XERO_REDIRECT, XERO_SCOPES ----
const XERO_CLIENT_ID_DEFAULT = "6387204C4C394207A0C825A12FB86975";
const XERO_REDIRECT_DEFAULT = "https://reservoir-api.izzyandloki.workers.dev/api/xero/callback";
const XERO_SCOPES_DEFAULT = "offline_access accounting.contacts accounting.settings accounting.invoices payroll.employees payroll.settings payroll.payruns payroll.payslip payroll.timesheets";
const xeroId = (env) => env.XERO_CLIENT_ID || XERO_CLIENT_ID_DEFAULT;
const xeroRedirect = (env) => env.XERO_REDIRECT || XERO_REDIRECT_DEFAULT;
const b64enc = (buf) => btoa(String.fromCharCode(...new Uint8Array(buf))).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, "");
async function xeroSign(env, text) {
  const key = await crypto.subtle.importKey("raw", new TextEncoder().encode(env.XERO_CLIENT_SECRET), { name: "HMAC", hash: "SHA-256" }, false, ["sign"]);
  return b64enc(await crypto.subtle.sign("HMAC", key, new TextEncoder().encode(text)));
}
async function xeroState(env, name) {
  const payload = b64enc(new TextEncoder().encode(JSON.stringify({ exp: Date.now() + 10 * 60 * 1000, name })));
  return `${payload}.${await xeroSign(env, payload)}`;
}
async function xeroStateOk(env, state) {
  const [payload, sig] = String(state || "").split(".");
  if (!payload || !sig || sig !== await xeroSign(env, payload)) return null;
  try {
    const data = parseJson(payload);
    return data.exp > Date.now() ? data : null;
  } catch { return null; }
}
const xeroBasic = (env) => "Basic " + btoa(`${xeroId(env)}:${env.XERO_CLIENT_SECRET}`);
let xeroTableMade = false;
async function xeroTable(env) {
  if (xeroTableMade) return;
  xeroTableMade = true;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_auth (id INTEGER PRIMARY KEY CHECK (id = 1), refresh_token TEXT, access_token TEXT, expires_at INTEGER, tenant_id TEXT, tenant_name TEXT, connected_by TEXT, updated_at TEXT DEFAULT (datetime('now')))").run();
}
// ---- Staff profile details that are DFP's own, not Xero's: uniform sizes, identification numbers, licences, tickets, qualifications and training ----
// One row per person, keyed by their name with spaces and punctuation removed. The email is stored when the person has signed in.
// Staff ask for changes through Team Requests (profile_requests); approving a request writes the values into dfp_reservoir. Directors and managers can also edit a person directly.
const PROFILE_COLS = { polo: "polo", jumper: "jumper", jacket: "jacket", whiteCard: "white_card", usi: "usi", leavePlus: "leave_plus", wwccNo: "wwcc_no", wwccExpiry: "wwcc_expiry", firstAidExpiry: "first_aid_expiry", cprExpiry: "cpr_expiry", dayHours: "day_hours", driverNo: "driver_no", driverExpiry: "driver_expiry", truckClass: "truck_class", truckNo: "truck_no", truckExpiry: "truck_expiry", forkliftNo: "forklift_no", forkliftExpiry: "forklift_expiry", excavatorNo: "excavator_no", excavatorExpiry: "excavator_expiry", plumbRegNo: "plumb_reg_no", plumbRegExpiry: "plumb_reg_expiry", driverCopy: "driver_copy", truckCopy: "truck_copy", forkliftCopy: "forklift_copy", excavatorCopy: "excavator_copy", plumbRegCopy: "plumb_reg_copy" };
const OWN_COLS = ["preferred_name", "birth_time", "em_name", "em_relationship", "em_mobile", "em_email", "em_saved_at"];
const PROFILE_DATES = new Set(["driverExpiry", "wwccExpiry", "firstAidExpiry", "cprExpiry", "truckExpiry", "forkliftExpiry", "excavatorExpiry", "plumbRegExpiry"]);
const profileKey = (name) => String(name || "").toLowerCase().replace(/[^a-z0-9]/g, "");
// Uniform sizes only. Identification numbers, licences and tickets are not seeded: their real values are entered through Reservoir.
const PROFILE_SEED_SIZES = [["Andrew Wowk", "L", "L", "XL"], ["Mitch Saba", "XL", "XL", "XL"], ["Hamish Burchell", "XXL", "XXL", "XXL"], ["Chad Lyons", "M", "L", "L"], ["Mark Hansen", "M", "L", "L"], ["Tyler Willey", "L", "XL", "XL"], ["Zac Rochow-O\u2019Toole", "L", "XL", "XL"], ["Tayler Reid", "L", "L", ""]];
// ---- DFP Reservoir: the one table for everything DFP keeps about a person that never leaves DFP ----
// Replaces staff_profile, staff_emergency and staff_personal_email. One row per person, keyed by their name with spaces and punctuation removed (person_key), with the work email stored as soon as it is known. Rows are found by email first, then by name.
// Holds: uniform sizes, identification numbers, licences, tickets, training, personal email and contract details. Vehicles stay in the fleet table.
// Preferred name, time of birth and the emergency contact are here too, because they never go to or come from Xero. Bank details go back and forth with Xero, so they live in dfp_xero.
// Nothing here is ever read from Xero or written to Xero by a sync. The old three tables are no longer read or written. They are left in place, untouched, until their data has been checked and Tayler says to drop them.
let reservoirReady = false;
async function reservoirTables(env) {
  if (reservoirReady) return;
  const cols = Object.values(PROFILE_COLS).map((c) => `${c} TEXT`).join(", ");
  await env.DB.prepare(`CREATE TABLE IF NOT EXISTS dfp_reservoir (person_key TEXT PRIMARY KEY, name TEXT, email TEXT, personal_email TEXT, contract_details TEXT, ${cols}, training TEXT, source TEXT, created_at TEXT DEFAULT (datetime('now')), updated_by TEXT, updated_at TEXT DEFAULT (datetime('now')))`).run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS dfp_reservoir_email ON dfp_reservoir (email)").run();
  // Preferred name, time of birth and the emergency contact: DFP's own details, never sent to or read from Xero.
  // Driver licence number and expiry, and the id of the photo of each licence or registration card (profile_photos), were added after the table was first made.
  for (const col of ["day_hours", "driver_no", "driver_expiry", "driver_copy", "truck_copy", "forklift_copy", "excavator_copy", "plumb_reg_copy"]) { try { await env.DB.prepare(`ALTER TABLE dfp_reservoir ADD COLUMN ${col} TEXT`).run(); } catch (e) { /* already there */ } }
  for (const col of OWN_COLS) { try { await env.DB.prepare(`ALTER TABLE dfp_reservoir ADD COLUMN ${col} TEXT`).run(); } catch (e) { /* already there */ } }
  await settingsTable(env);
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS profile_requests (id TEXT PRIMARY KEY, person_key TEXT, name TEXT, email TEXT, section TEXT, section_title TEXT, fields TEXT, vals TEXT, text TEXT, status TEXT NOT NULL DEFAULT 'new', created_at TEXT DEFAULT (datetime('now')), processed_by TEXT, processed_at TEXT)").run();
  // seen_at: when the person first saw the "details updated" banner after the office processed their change. profile_audit: the permanent record of every change request, who sent it, who processed it and when. Rows here are never removed.
  try { await env.DB.prepare("ALTER TABLE profile_requests ADD COLUMN seen_at TEXT").run(); } catch (e) { /* already there */ }
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS profile_audit (id INTEGER PRIMARY KEY AUTOINCREMENT, at TEXT DEFAULT (datetime('now')), event TEXT, request_id TEXT, person TEXT, email TEXT, section TEXT, section_title TEXT, fields TEXT, by_name TEXT, note TEXT)").run();
  // Profile changes now show a one time "updated" banner. Ones processed before that are marked as already seen, once.
  if (!(await migDone(env, "profile_seen_v1"))) {
    await env.DB.prepare("UPDATE profile_requests SET seen_at = datetime('now') WHERE status = 'processed' AND seen_at IS NULL AND section IN ('personal', 'emergency')").run();
    await migMark(env, "profile_seen_v1", "old processed profile requests marked seen");
  }
  await dfpTables(env);
  // Uniform sizes that were in the app's staff list are put in the database so staff see what they saw before. Only a size that has never been saved (NULL) is filled. A size someone has saved or deliberately cleared is never touched.
  for (const [name, polo, jumper, jacket] of PROFILE_SEED_SIZES) {
    const key = profileKey(name);
    await env.DB.prepare("INSERT OR IGNORE INTO dfp_reservoir (person_key, name, source, updated_by) VALUES (?, ?, 'initial sizes', 'initial sizes')").bind(key, name).run();
    await env.DB.prepare("UPDATE dfp_reservoir SET polo = COALESCE(polo, ?), jumper = COALESCE(jumper, ?), jacket = COALESCE(jacket, ?) WHERE person_key = ?").bind(polo, jumper, jacket, key).run();
  }
  await moveOwnFields(env);
  reservoirReady = true;
}
// One time: preferred name, time of birth and the emergency contact move from dfp_xero to dfp_reservoir. Only blanks in dfp_reservoir are filled. Every value is checked after it is copied, and the old dfp_xero columns are dropped only when every one checked out. If anything did not check out nothing is dropped and the whole step runs again on the next start.
async function moveOwnFields(env) {
  if (await migDone(env, "own_fields_to_reservoir_v1")) return;
  const detail = { people: 0, dropped: [], problems: [] };
  const info = ((await env.DB.prepare("PRAGMA table_info(dfp_xero)").all()).results) || [];
  const have = OWN_COLS.filter((c) => info.some((x) => x.name === c));
  const filled = (v) => v !== null && v !== undefined && String(v) !== "";
  if (have.length) {
    const rows = ((await env.DB.prepare(`SELECT email, name, ${have.join(", ")} FROM dfp_xero`).all()).results) || [];
    for (const r of rows) {
      if (!have.some((c) => filled(r[c]))) continue;
      detail.people++;
      const e = String(r.email || "").toLowerCase(), k = profileKey(r.name);
      let d = e ? await env.DB.prepare("SELECT person_key FROM dfp_reservoir WHERE email = ?").bind(e).first() : null;
      if (!d && k) d = await env.DB.prepare("SELECT person_key FROM dfp_reservoir WHERE person_key = ?").bind(k).first();
      // Someone known by a shorter name (Mitch for Mitchell) has a dfp_reservoir row under that name, so look there before making a new one.
      const known = String(r.preferred_name || "").trim(), lastName = String(r.name || "").trim().split(/\s+/).slice(1).join(" ");
      if (!d && known && lastName) d = await env.DB.prepare("SELECT person_key FROM dfp_reservoir WHERE person_key = ?").bind(profileKey(`${known} ${lastName}`)).first();
      if (!d) {
        const key = k || "email:" + e;
        await env.DB.prepare("INSERT INTO dfp_reservoir (person_key, name, email, source, updated_by) VALUES (?, ?, ?, 'moved from dfp_xero', 'move')").bind(key, r.name || null, e || null).run();
        d = { person_key: key };
      }
      if (e) await env.DB.prepare("UPDATE dfp_reservoir SET email = ? WHERE person_key = ? AND (email IS NULL OR email = '')").bind(e, d.person_key).run();
      for (const c of have) if (filled(r[c])) await env.DB.prepare(`UPDATE dfp_reservoir SET ${c} = ? WHERE person_key = ? AND (${c} IS NULL OR ${c} = '')`).bind(r[c], d.person_key).run();
      const back = await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE person_key = ?").bind(d.person_key).first();
      for (const c of have) if (filled(r[c]) && !filled(back && back[c])) detail.problems.push(`${r.email}: ${c}`);
    }
    if (!detail.problems.length) {
      for (const c of have) {
        try { await env.DB.prepare(`ALTER TABLE dfp_xero DROP COLUMN ${c}`).run(); detail.dropped.push(c); }
        catch (e) { detail.problems.push(`could not drop ${c}: ${String(e && e.message || e).slice(0, 80)}`); }
      }
    }
  }
  if (detail.problems.length) { console.log("own fields move", JSON.stringify(detail.problems)); return; }
  // One time, 20 October 2026: the director asked for Zac and Mitch to be set as preferred names. Only fills a blank, never replaces one, and runs once both people have been found.
  if (!(await migDone(env, "preferred_names_zac_mitch_v1"))) {
    const z = await env.DB.prepare("UPDATE dfp_reservoir SET preferred_name = 'Zac' WHERE lower(email) IN (SELECT lower(email) FROM dfp_xero WHERE lower(name) LIKE 'zachary rochow%') AND (preferred_name IS NULL OR TRIM(preferred_name) = '')").run();
    const m = await env.DB.prepare("UPDATE dfp_reservoir SET preferred_name = 'Mitch' WHERE lower(email) IN (SELECT lower(email) FROM dfp_xero WHERE lower(name) = 'mitchell saba') AND (preferred_name IS NULL OR TRIM(preferred_name) = '')").run();
    const both = await env.DB.prepare("SELECT COUNT(*) AS n FROM dfp_xero WHERE (lower(name) LIKE 'zachary rochow%' OR lower(name) = 'mitchell saba')").first();
    if (both && both.n >= 2) await migMark(env, "preferred_names_zac_mitch_v1", JSON.stringify({ zac: z && z.meta && z.meta.changes, mitch: m && m.meta && m.meta.changes }));
  }
  await migMark(env, "own_fields_to_reservoir_v1", JSON.stringify(detail));
}
// Saves preferred name, time of birth or emergency contact fields on a person's dfp_reservoir row, making the row if the person has none yet.
async function reservoirOwnSet(env, email, name, patch, by) {
  const r = await reservoirRowFor(env, email, name, true);
  if (!r) throw new Error("There is no person to save this against");
  const keys = Object.keys(patch).filter((k) => OWN_COLS.includes(k));
  if (!keys.length) return;
  await env.DB.prepare(`UPDATE dfp_reservoir SET ${keys.map((k) => `${k} = ?`).join(", ")}, updated_by = ?, updated_at = datetime('now') WHERE person_key = ?`).bind(...keys.map((k) => patch[k]), by || "unknown", r.person_key).run();
}
async function profileTables(env) { return reservoirTables(env); }
// ---- Reading people from DFP's own tables. These replace the live Xero reads for the Details, Profile and Team Members screens. ----
const NOT_IN_RESERVOIR = "Your details are not in Reservoir yet. Ask the office to run the one time fill.";
// The legal full name as Xero holds it: first name, any middle names, then the surname. The plain name stays First Last because everything else in the app matches people on that.
function fullNameOf(name, middle) {
  const n = String(name || "").trim().replace(/\s+/g, " ");
  const m = String(middle || "").trim().replace(/\s+/g, " ");
  if (!n || !m) return n;
  const i = n.indexOf(" ");
  return i < 0 ? `${n} ${m}` : `${n.slice(0, i)} ${m} ${n.slice(i + 1)}`;
}
// A dfp_xero row with the person's own DFP details (preferred name, time of birth, emergency contact) from dfp_reservoir, joined on work email.
const PERSON_SQL = "SELECT x.*, r.preferred_name, r.birth_time, r.em_name, r.em_relationship, r.em_mobile, r.em_email, r.em_saved_at FROM dfp_xero x LEFT JOIN dfp_reservoir r ON lower(r.email) = lower(x.email)";
// One dfp_xero row as the app's person record (the same shape the live Xero read used to give). Emergency contact only when DFP has one saved.
function dfpPersonOut(r) {
  const saved = !!r.em_saved_at;
  return {
    id: r.xero_employee_id || "", status: r.status || "ACTIVE", name: r.name || "", address: r.address || "", dob: r.dob || "", mobile: r.mobile || "", email: r.email || "",
    accountName: r.bank_account_name || "", bsb: r.bsb || "", accountNumber: r.account_number || "", commenced: r.start_date || "", title: String(r.job_title || "").trim(), preferredName: String(r.preferred_name || "").trim(), birthTime: /^\d{2}:\d{2}$/.test(String(r.birth_time || "")) ? r.birth_time : "", fullName: fullNameOf(r.name, r.middle_names),
    emName: saved ? r.em_name || "" : "", emRelationship: saved ? r.em_relationship || "" : "", emMobile: saved ? r.em_mobile || "" : "", emEmail: saved ? r.em_email || "" : ""
  };
}
// One xero_payroll row as the pay details payload the Details screen reads.
function dfpPayOut(r) {
  const ml = r.leave_loading === null || r.leave_loading === undefined ? null : Number(r.leave_loading) === 1;
  const hasFund = !!(r.super_fund_name || r.super_member_number || r.super_usi);
  return { found: true, rates: { ordinary: r.rate_ordinary === undefined ? null : r.rate_ordinary, ot15: r.rate_ot15 === undefined ? null : r.rate_ot15, ot2: r.rate_ot2 === undefined ? null : r.rate_ot2 }, hasOvertime: Number(r.has_overtime) === 1, classification: r.classification || "", employmentType: r.employment_basis || "", leaveLoading: ml, superMemberships: hasFund ? [{ fundName: r.super_fund_name || "", memberNumber: r.super_member_number || "", usi: r.super_usi || "" }] : [], superWhy: hasFund ? "" : "No super fund on record", missing: r.rates_note ? String(r.rates_note).split("; ").filter(Boolean) : [], used: [], templateLines: [], syncedAt: r.synced_at || "", annualSalary: r.annual_salary === null || r.annual_salary === undefined ? null : Number(r.annual_salary), weekHours: r.week_hours === null || r.week_hours === undefined ? null : Number(r.week_hours), superPct: r.super_pct === null || r.super_pct === undefined ? null : Number(r.super_pct), noAward: !String(r.classification || "").trim() };
}
// Makes sure a dfp_xero row exists for this work email. A new employee is always created in Xero first, so the row normally exists after the one time fill. This only creates a bare row (email and name) for anyone who has not been filled yet.
async function dfpXeroRowFor(env, email, name) {
  const key = String(email || "").trim().toLowerCase();
  await dfpTables(env);
  await env.DB.prepare("INSERT OR IGNORE INTO dfp_xero (email, name, status, source, updated_by) VALUES (?, ?, 'ACTIVE', 'reservoir', ?)").bind(key, name || null, key).run();
  return key;
}
// Finds a person's row by work email, then by name. With create, makes one if there is none. When the email is given and the row has none (or a different one), the row keeps it.
async function reservoirRowFor(env, email, name, create) {
  await reservoirTables(env);
  const e = String(email || "").trim().toLowerCase();
  const k = profileKey(name);
  let r = e ? await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE email = ?").bind(e).first() : null;
  if (!r && k) r = await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE person_key = ?").bind(k).first();
  if (!r && create) {
    const key = k || (e ? "email:" + e : "");
    if (!key) return null;
    await env.DB.prepare("INSERT INTO dfp_reservoir (person_key, name, email, source, updated_by) VALUES (?, ?, ?, 'reservoir', ?)").bind(key, name ? String(name).slice(0, 120) : null, e || null, e || "unknown").run();
    r = await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE person_key = ?").bind(key).first();
  } else if (r && e && r.email !== e) {
    await env.DB.prepare("UPDATE dfp_reservoir SET email = ? WHERE person_key = ?").bind(e, r.person_key).run();
    r.email = e;
  }
  return r;
}
// Short codes and sizes are always saved in capitals, so N/A is never saved as n/a.
const PROFILE_UPPER = new Set(["polo", "jumper", "jacket", "whiteCard", "usi", "leavePlus", "wwccNo", "driverNo", "truckClass", "truckNo", "forkliftNo", "excavatorNo", "plumbRegNo"]);
function profileClean(k, v) {
  let t = String(v === undefined || v === null ? "" : v).replace(/\s+/g, " ").trim().slice(0, 120);
  if (PROFILE_UPPER.has(k)) t = t.toUpperCase();
  if (k === "dayHours" && t) { const n = Number(t); if (!(n >= 1 && n <= 12)) return null; t = String(Math.round(n * 100) / 100); }
  if (PROFILE_DATES.has(k) && t && !/^\d{4}-\d{2}-\d{2}$/.test(t)) return null;
  return t;
}
function profileTraining(list) {
  if (!Array.isArray(list)) return null;
  const out = [];
  for (const x of list.slice(0, 60)) {
    const name = String((x && x.name) || "").replace(/\s+/g, " ").trim().slice(0, 120);
    if (!name) continue;
    const completed = String((x && x.completed) || "").trim(), expiry = String((x && x.expiry) || "").trim();
    if ((completed && !/^\d{4}-\d{2}-\d{2}$/.test(completed)) || (expiry && !/^\d{4}-\d{2}-\d{2}$/.test(expiry))) return null;
    out.push({ name, completed, expiry });
  }
  return out;
}
const profileOut = (r) => {
  const o = { name: r.name || "", email: r.email || "", training: [] };
  for (const [k, c] of Object.entries(PROFILE_COLS)) o[k] = r[c] || "";
  try { o.training = JSON.parse(r.training || "[]"); } catch (e) { /* none */ }
  return o;
};
// ---- Licence and registration cards: self managed ----
// The person saves their own card straight away (number, expiry, photos). The office is told with a notice and can check the photo. If the photo is not good enough the office resets that card: it is cleared and the person gets a job on their punch list to do it again.
const CARD_DEFS = {
  plumbReg: { label: "Plumbing Registration", keys: ["plumbRegNo", "plumbRegExpiry"], copy: "plumbRegCopy" },
  driver: { label: "Driver Licence", keys: ["driverNo", "driverExpiry"], copy: "driverCopy" },
  truck: { label: "Truck Licence", keys: ["truckClass", "truckNo", "truckExpiry"], copy: "truckCopy" },
  forklift: { label: "Forklift Licence", keys: ["forkliftNo", "forkliftExpiry"], copy: "forkliftCopy" },
  excavator: { label: "Excavator Licence", keys: ["excavatorNo", "excavatorExpiry"], copy: "excavatorCopy" }
};
const dmyOf = (iso) => /^\d{4}-\d{2}-\d{2}$/.test(String(iso || "")) ? `${iso.slice(8, 10)}/${iso.slice(5, 7)}/${iso.slice(0, 4)}` : String(iso || "");
async function profileFor(env, user) {
  await reservoirTables(env);
  const email = String(user.email || "").toLowerCase();
  let r = email ? await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE email = ?").bind(email).first() : null;
  if (!r && user.name) r = await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE person_key = ?").bind(profileKey(user.name)).first();
  return r ? profileOut(r) : null;
}
// Writes the allowed fields for one person. Only keys that are present are changed. Returns the saved record.
async function profileWrite(env, person, fields, by) {
  await reservoirTables(env);
  const key = profileKey(person.name);
  if (!key) throw new Error("no name");
  const sets = [], binds = [];
  for (const [k, col] of Object.entries(PROFILE_COLS)) {
    if (!Object.prototype.hasOwnProperty.call(fields, k)) continue;
    const v = profileClean(k, fields[k]);
    if (v === null) throw new Error(`bad value for ${k}`);
    sets.push(col); binds.push(v);
  }
  if (Object.prototype.hasOwnProperty.call(fields, "training")) {
    const t = profileTraining(fields.training);
    if (t === null) throw new Error("bad training list");
    sets.push("training"); binds.push(JSON.stringify(t));
  }
  const row = await reservoirRowFor(env, person.email, person.name, true);
  if (sets.length) {
    await env.DB.prepare(`UPDATE dfp_reservoir SET ${sets.map((c) => `${c} = ?`).join(", ")}, updated_by = ?, updated_at = datetime('now') WHERE person_key = ?`).bind(...binds, by, row.person_key).run();
  }
  const r = await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE person_key = ?").bind(row.person_key).first();
  return profileOut(r);
}

// ---- Bank and super changes. The director presses Processed and the new details are written into Xero first. Only when Xero has taken them (and shows them back) are they copied into Reservoir's own tables and the request marked processed. ----
const digitsOnly = (v) => String(v || "").replace(/\D/g, "");
const maskAcct = (v) => { const d = digitsOnly(v); return d ? `XXXX ${d.slice(-4)}` : String(v || ""); };
// One row in the permanent change record. Account numbers are kept masked here. Never blocks the change it records.
async function profileAudit(env, event, row, by, note) {
  try {
    let fields = [];
    try { fields = JSON.parse(row.fields || "[]"); } catch (e) { /* none */ }
    await env.DB.prepare("INSERT INTO profile_audit (event, request_id, person, email, section, section_title, fields, by_name, note) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)")
      .bind(event, row.id, row.name || "", row.email || "", row.section || "", row.section_title || "", JSON.stringify(fields), String(by || "").slice(0, 120), String(note || "").slice(0, 300)).run();
  } catch (e) { console.log("profile audit error", String(e && e.message || e)); }
}
// Xero's answer as plain text, including the validation messages it sends back on a refusal.
async function xeroRefusal(res) {
  const raw = (await res.text().catch(() => "")).trim();
  let msg = "";
  try {
    const j = JSON.parse(raw);
    const errs = [];
    const walk = (o) => { if (!o || typeof o !== "object") return; if (Array.isArray(o)) { o.forEach(walk); return; } if (o.ValidationErrors) o.ValidationErrors.forEach((x) => x && x.Message && errs.push(x.Message)); if (o.Message) errs.push(o.Message); Object.values(o).forEach(walk); };
    walk(j);
    msg = [...new Set(errs)].join(" ");
  } catch (e) { /* not JSON */ }
  return `Xero answered ${res.status}. ${msg || raw.replace(/\s+/g, " ").slice(0, 240) || "It sent no message."}`;
}
async function xeroEmployeeRecord(env, employeeId) {
  const r = await xeroFetch(env, `payroll.xro/1.0/Employees/${employeeId}`);
  if (!r.ok) throw new Error(await xeroRefusal(r));
  const rec = ((await r.json()).Employees || [])[0];
  if (!rec) throw new Error("Xero returned no employee record");
  return rec;
}
async function xeroPostEmployee(env, rec, patch) {
  const body = [{ EmployeeID: rec.EmployeeID, FirstName: rec.FirstName, LastName: rec.LastName, DateOfBirth: rec.DateOfBirth, ...patch }];
  const r = await xeroFetch(env, `payroll.xro/1.0/Employees/${rec.EmployeeID}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
  if (!r.ok) throw new Error(await xeroRefusal(r));
}
async function xeroPushBank(env, employeeId, v) {
  const rec = await xeroEmployeeRecord(env, employeeId);
  const accts = Array.isArray(rec.BankAccounts) ? rec.BankAccounts : [];
  if (accts.length > 1) throw new Error("This person has more than one bank account in Xero. Change it in Xero by hand.");
  const cur = accts[0] || {};
  const bsb = v.bsb !== undefined ? digitsOnly(v.bsb) : digitsOnly(cur.BSB);
  const num = v.accountNumber !== undefined ? digitsOnly(v.accountNumber) : digitsOnly(cur.AccountNumber);
  const name = v.accountName !== undefined ? String(v.accountName).trim() : String(cur.AccountName || "").trim();
  if (bsb.length !== 6) throw new Error("The BSB must be 6 digits.");
  if (num.length < 6 || num.length > 9) throw new Error("The account number must be 6 to 9 digits.");
  if (!name) throw new Error("The account name is blank.");
  await xeroPostEmployee(env, rec, { BankAccounts: [{ StatementText: cur.StatementText || "Wages", AccountName: name, BSB: bsb, AccountNumber: num, Remainder: cur.Remainder !== false }] });
  const got = ((await xeroEmployeeRecord(env, employeeId)).BankAccounts || [])[0] || {};
  if (digitsOnly(got.BSB) !== bsb || digitsOnly(got.AccountNumber) !== num) throw new Error("Xero accepted the change but still shows the old bank details. Check the employee in Xero.");
  return { accountName: name, bsb, accountNumber: num };
}
// Super is changed in Xero by hand. Processed only reads the person's current fund back from Xero, checks it matches what they asked for, and returns it for the local tables.
async function xeroPullSuper(env, employeeId, v) {
  const rec = await xeroEmployeeRecord(env, employeeId);
  const sup = await xeroSuperFor(env, rec);
  const got = (sup.memberships || [])[0];
  if (!got) throw new Error("Xero has no super fund on this person's record yet. Set it up in Xero first, then press Processed again.");
  const norm = (x) => String(x || "").replace(/\s+/g, "").toUpperCase();
  if (v.superMember !== undefined && norm(got.memberNumber) !== norm(v.superMember)) throw new Error(`Xero still shows member number ${got.memberNumber || "blank"}, not ${v.superMember}. Update it in Xero first, then press Processed again.`);
  if (v.superUsi !== undefined && norm(got.usi) !== norm(v.superUsi)) throw new Error(`Xero still shows USI ${got.usi || "blank"}, not ${v.superUsi}. Update it in Xero first, then press Processed again.`);
  return { fundName: got.fundName, usi: got.usi, member: got.memberNumber };
}
async function xeroPushSuper(env, employeeId, v) {
  const rec = await xeroEmployeeRecord(env, employeeId);
  const ms = (Array.isArray(rec.SuperMemberships) ? rec.SuperMemberships : []).map((m) => ({ ...m }));
  const cur = ms[0] || null;
  let fundId = cur ? cur.SuperFundID : "";
  let fundName = "", usiOut = "";
  const wantUsi = v.superUsi !== undefined ? String(v.superUsi || "").replace(/\s+/g, "").toUpperCase() : "";
  if (v.superUsi !== undefined || v.superFund !== undefined) {
    const fr = await xeroFetch(env, "payroll.xro/1.0/Superfunds");
    if (!fr.ok) throw new Error(await xeroRefusal(fr));
    const funds = (await fr.json()).SuperFunds || [];
    let f = wantUsi ? funds.find((x) => String(x.USI || "").replace(/\s+/g, "").toUpperCase() === wantUsi) : null;
    if (!f && !wantUsi && v.superFund) { const n = String(v.superFund).trim().toLowerCase(); f = funds.find((x) => String(x.Name || "").trim().toLowerCase() === n) || null; }
    if (!f) throw new Error("That super fund is not set up in Xero yet. Add the fund in Xero, then press Processed again.");
    fundId = f.SuperFundID;
    fundName = String(f.Name || "").trim();
    usiOut = String(f.USI || "").trim();
  }
  if (!fundId) throw new Error("There is no super fund on this person's Xero record to change.");
  const member = v.superMember !== undefined ? String(v.superMember || "").trim() : String(cur ? cur.EmployeeNumber || "" : "").trim();
  if (!member) throw new Error("The member number is blank.");
  if (cur) ms[0] = { ...cur, SuperFundID: fundId, EmployeeNumber: member };
  else ms.push({ SuperFundID: fundId, EmployeeNumber: member });
  await xeroPostEmployee(env, rec, { SuperMemberships: ms });
  const got = ((await xeroEmployeeRecord(env, employeeId)).SuperMemberships || [])[0] || {};
  if (got.SuperFundID !== fundId || String(got.EmployeeNumber || "").trim() !== member) throw new Error("Xero accepted the change but still shows the old super details. Check the employee in Xero.");
  return { fundName, usi: usiOut, member };
}
// Address as the app saves it ("12 Smith Street, Geelong VIC 3220") into the parts Xero keeps.
function xeroAddressParts(addr, current) {
  let raw = String(addr || "").replace(/\s+/g, " ").trim();
  if (raw && raw === raw.toUpperCase()) raw = raw.toLowerCase().replace(/(^|[\s,-])([a-z])/g, (x, p, c) => p + c.toUpperCase()).replace(/\b(Vic|Nsw|Qld|Sa|Wa|Tas|Nt|Act)\b/g, (s) => s.toUpperCase());
  const m = raw.match(/^(.+),\s*(.+?)\s+(VIC|NSW|QLD|SA|WA|TAS|NT|ACT)\s+(\d{4})$/i);
  if (!m) throw new Error("The address needs a street, suburb, state and postcode so Xero can hold it. Ask them to pick it from the address suggestions.");
  return { AddressLine1: m[1].trim(), City: m[2].trim(), Region: m[3].toUpperCase(), PostalCode: m[4], Country: String(current && current.Country || "AUSTRALIA") };
}
// Address and mobile go to Xero, then are read straight back from Xero into the local table, so what staff see is what Xero holds.
async function xeroPushPersonal(env, employeeId, v, email) {
  const rec = await xeroEmployeeRecord(env, employeeId);
  const patch = {};
  if (v.address !== undefined) patch.HomeAddress = xeroAddressParts(v.address, rec.HomeAddress);
  if (v.mobile !== undefined) {
    const d = digitsOnly(normPhone(v.mobile));
    if (d.length < 8 || d.length > 12) throw new Error("The mobile number does not look right.");
    patch.Mobile = d;
  }
  await xeroPostEmployee(env, rec, patch);
  const after = await xeroEmployeeRecord(env, employeeId);
  const staff = xeroStaffRow(after);
  if (patch.Mobile && digitsOnly(after.Mobile) !== patch.Mobile) throw new Error("Xero accepted the change but still shows the old mobile number. Check the employee in Xero.");
  if (patch.HomeAddress && String((after.HomeAddress || {}).AddressLine1 || "").trim().toLowerCase() !== patch.HomeAddress.AddressLine1.toLowerCase()) throw new Error("Xero accepted the change but still shows the old address. Check the employee in Xero.");
  const sets = [], binds = [];
  if (patch.HomeAddress) { sets.push("address = ?"); binds.push(staff.address); }
  if (patch.Mobile) { sets.push("mobile = ?"); binds.push(staff.mobile); }
  if (sets.length) await env.DB.prepare(`UPDATE dfp_xero SET ${sets.join(", ")}, updated_by = ?, updated_at = datetime('now') WHERE lower(email) = ?`).bind(...binds, "Xero", String(email || "").toLowerCase()).run();
}
const profileRequestOut = (r) => {
  let fields = [], values = {};
  try { fields = JSON.parse(r.fields || "[]"); } catch (e) { /* none */ }
  try { values = JSON.parse(r.vals || "{}"); } catch (e) { /* none */ }
  return { id: r.id, kind: "detailschange", type: "Details Change", section: r.section, sectionTitle: r.section_title, fields, values, text: r.text || "", status: r.status, by: String(r.name || "").split(" ")[0], byName: r.name || "", date: String(r.created_at || "").slice(0, 10), doneDate: r.processed_at ? String(r.processed_at).slice(0, 10) : "", seen: !!r.seen_at, server: true };
};
// The personal email is read from dfp_reservoir. Keyed by work email.
async function personalEmailFor(env, email) {
  const key = String(email || "").trim().toLowerCase();
  if (!key) return null;
  await reservoirTables(env);
  const r = await env.DB.prepare("SELECT personal_email FROM dfp_reservoir WHERE email = ?").bind(key).first();
  return r && r.personal_email ? { personalEmail: r.personal_email } : null;
}
async function xeroTokenCall(env, params) {
  const res = await fetch("https://identity.xero.com/connect/token", { method: "POST", headers: { Authorization: xeroBasic(env), "Content-Type": "application/x-www-form-urlencoded" }, body: new URLSearchParams(params) });
  if (!res.ok) throw new Error("xero token " + res.status + " " + (await res.text()).slice(0, 200));
  return res.json();
}
// Returns a valid access token and tenant id, refreshing (and saving the rotated refresh token) when needed.
// Only one refresh runs at a time, because Xero refresh tokens rotate and parallel refreshes can leave a dead token saved.
let xeroRefreshing = null;
// ---- THE XERO GATE. Every call to Xero, and every use of the stored Xero login, goes through xeroAccess or xeroFetch, and both start here. A request that arrives from a signed in person carries their access level on env (set once, in fetch, for the whole request). Anyone who is not T1 or T2, T3 included, is stopped before anything is sent to Xero, whatever route they came in on. Viewing as someone is judged by the level of the person being viewed. Work started by the Worker itself (the scheduled leave refresh, the connect screen) carries no level and is allowed. ----
function xeroGate(env) {
  const tier = env && env.__callerTier;
  if (tier === undefined) return;
  if (tier !== "T1" && tier !== "T2") {
    const e = new Error("Xero is closed to your access level");
    e.xeroBlocked = true;
    throw e;
  }
}
async function xeroAccess(env, force) {
  xeroGate(env);
  await xeroTable(env);
  const row = await env.DB.prepare("SELECT * FROM xero_auth WHERE id = 1").first();
  if (!row || !row.refresh_token) throw new Error("xero not connected");
  if (!force && row.access_token && row.expires_at > Date.now() + 60000) return { token: row.access_token, tenant: row.tenant_id };
  if (xeroRefreshing) return xeroRefreshing;
  xeroRefreshing = (async () => {
    try {
      const t = await xeroTokenCall(env, { grant_type: "refresh_token", refresh_token: row.refresh_token });
      await env.DB.prepare("UPDATE xero_auth SET refresh_token = ?, access_token = ?, expires_at = ?, updated_at = datetime('now') WHERE id = 1").bind(t.refresh_token, t.access_token, Date.now() + t.expires_in * 1000).run();
      return { token: t.access_token, tenant: row.tenant_id };
    } catch (e) {
      // Another request may already have rotated the token. Use theirs if it is saved and still good.
      const again = await env.DB.prepare("SELECT * FROM xero_auth WHERE id = 1").first();
      if (again && again.access_token && again.expires_at > Date.now() + 60000 && again.access_token !== row.access_token) return { token: again.access_token, tenant: again.tenant_id };
      throw e;
    } finally { xeroRefreshing = null; }
  })();
  return xeroRefreshing;
}
const xeroSleep = (ms) => new Promise((r) => setTimeout(r, ms));
// ---- Go live lock. While testing, a director (T1) can remove closed requests. On go live day the director presses Lock Request Removal on the Development screen. From then on nobody can remove a request, so every request stays on record. There is no unlock in the app. ----
// ---- Files kept in the database for the app to use (Development tab, Files card) ----
// Only slots named here can be stored. readAll is true when any signed in person may read the file (the pay slip logo is printed on everyone's slip).
const FILE_SLOTS = { "payslip-logo": { label: "Pay Slip Logo", readAll: true, mimes: ["image/png", "image/jpeg"], maxChars: 900000 } };
let filesTableMade = false;
async function filesTable(env) {
  if (filesTableMade) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS dfp_files (slot TEXT PRIMARY KEY, name TEXT, mime TEXT, size INTEGER, data TEXT, updated_at TEXT DEFAULT (datetime('now')), updated_by TEXT)").run();
  filesTableMade = true;
}
// One settings table for everything that is a name and a value: go live switches (launch.*), company passwords and IDs (secret.*) and the flags for one time migrations already run (migration.*). It replaces launch_settings, secrets and dfp_migrations. The first start after this deploy copies those three across once. They are then no longer read or written, and are left in place until Tayler says to drop them.
let settingsTableMade = false;
async function settingsTable(env) {
  if (settingsTableMade) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS app_settings (key TEXT PRIMARY KEY, value TEXT, detail TEXT, by_name TEXT, at TEXT DEFAULT (datetime('now')))").run();
  const merged = await env.DB.prepare("SELECT key FROM app_settings WHERE key = 'migration.settings_merge_v1'").first();
  if (!merged) {
    const has = async (t) => !!(await env.DB.prepare("SELECT name FROM sqlite_master WHERE type = 'table' AND name = ?").bind(t).first());
    if (await has("launch_settings")) await env.DB.prepare("INSERT OR IGNORE INTO app_settings (key, value, by_name, at) SELECT 'launch.' || key, value, set_by, set_at FROM launch_settings").run();
    if (await has("secrets")) {
      const sc = new Set((((await env.DB.prepare("PRAGMA table_info(secrets)").all()).results) || []).map((c) => c.name));
      await env.DB.prepare(`INSERT OR IGNORE INTO app_settings (key, value, by_name, at) SELECT 'secret.' || key, value, ${sc.has("updated_by") ? "updated_by" : "NULL"}, ${sc.has("updated_at") ? "updated_at" : "datetime('now')"} FROM secrets`).run();
    }
    if (await has("dfp_migrations")) await env.DB.prepare("INSERT OR IGNORE INTO app_settings (key, detail, at) SELECT 'migration.' || name, detail, done_at FROM dfp_migrations").run();
    await env.DB.prepare("INSERT OR IGNORE INTO app_settings (key, detail) VALUES ('migration.settings_merge_v1', 'launch_settings, secrets and dfp_migrations copied in')").run();
  }
  settingsTableMade = true;
}
const launchTable = settingsTable;
async function migDone(env, name) {
  await settingsTable(env);
  return !!(await env.DB.prepare("SELECT key FROM app_settings WHERE key = ?").bind("migration." + name).first());
}
async function migMark(env, name, detail) {
  await settingsTable(env);
  await env.DB.prepare("INSERT OR REPLACE INTO app_settings (key, detail, at) VALUES (?, ?, datetime('now'))").bind("migration." + name, detail || "").run();
}
async function requestsLock(env) {
  await settingsTable(env);
  const r = await env.DB.prepare("SELECT value, by_name, at FROM app_settings WHERE key = 'launch.requests_locked'").first();
  return r && r.value === "1" ? { locked: true, by: r.by_name || "", at: r.at || "" } : { locked: false, by: "", at: "" };
}
const REMOVAL_LOCKED = "Removing requests is locked. Every request is kept for the record.";
// ---- Xero call counter. One row per Melbourne day. Every call made to api.xero.com is counted (retries after a 401 or a rate limit count too, as Xero counts them), and the "calls left" figures Xero sends back are kept. Counting never stops a call from working. ----
let xeroCallsTableMade = false;
async function xeroCallsTable(env) {
  if (xeroCallsTableMade) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_calls (day TEXT PRIMARY KEY, calls INTEGER NOT NULL DEFAULT 0, day_remaining INTEGER, min_remaining INTEGER, remaining_at TEXT)").run();
  xeroCallsTableMade = true;
}
function melbourneDay(ms) {
  const parts = new Intl.DateTimeFormat("en-CA", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit" }).formatToParts(new Date(ms === undefined ? Date.now() : ms));
  const g = (t) => (parts.find((x) => x.type === t) || {}).value;
  return `${g("year")}-${g("month")}-${g("day")}`;
}
async function xeroCountCall(env, res) {
  try {
    await xeroCallsTable(env);
    const num = (name) => { const v = res && res.headers ? res.headers.get(name) : null; const n = v === null || v === undefined || v === "" ? NaN : Number(v); return Number.isFinite(n) ? Math.trunc(n) : null; };
    const dayLeft = num("X-DayLimit-Remaining");
    const minLeft = num("X-MinLimit-Remaining");
    await env.DB.prepare("INSERT INTO xero_calls (day, calls, day_remaining, min_remaining, remaining_at) VALUES (?, 1, ?, ?, CASE WHEN ? IS NULL AND ? IS NULL THEN NULL ELSE datetime('now') END) ON CONFLICT(day) DO UPDATE SET calls = calls + 1, day_remaining = COALESCE(?, day_remaining), min_remaining = COALESCE(?, min_remaining), remaining_at = CASE WHEN ? IS NULL AND ? IS NULL THEN remaining_at ELSE datetime('now') END").bind(melbourneDay(), dayLeft, minLeft, dayLeft, minLeft, dayLeft, minLeft, dayLeft, minLeft).run();
  } catch (e) { /* the count is best effort */ }
}
async function xeroCallsToday(env) {
  await xeroCallsTable(env);
  const day = melbourneDay();
  const r = await env.DB.prepare("SELECT calls, day_remaining, min_remaining, remaining_at FROM xero_calls WHERE day = ?").bind(day).first();
  return { day, today: r ? Number(r.calls) : 0, dayRemaining: r && r.day_remaining !== null && r.day_remaining !== undefined ? Number(r.day_remaining) : null, minuteRemaining: r && r.min_remaining !== null && r.min_remaining !== undefined ? Number(r.min_remaining) : null, remainingAt: r && r.remaining_at ? r.remaining_at : null };
}
// Calls Xero. A 401 gets one forced token refresh and a retry. A 429 (rate limit) waits for Xero's Retry-After and tries again, up to twice.
async function xeroFetch(env, path, init = {}) {
  xeroGate(env);
  let refreshed = false;
  let force = false;
  let waits = 0;
  for (;;) {
    const { token, tenant } = await xeroAccess(env, force);
    force = false;
    const res = await fetch(`https://api.xero.com/${path}`, { ...init, headers: { Authorization: `Bearer ${token}`, "Xero-tenant-id": tenant, Accept: "application/json", ...(init.headers || {}) } });
    await xeroCountCall(env, res);
    if (res.status === 401 && !refreshed) { refreshed = true; force = true; continue; }
    if (res.status === 429 && waits < 2) {
      waits++;
      const secs = Math.min(Math.max(Number(res.headers.get("Retry-After")) || 2, 1), 8);
      console.log("xero rate limited, waiting", secs);
      await xeroSleep(secs * 1000);
      continue;
    }
    return res;
  }
}


// Xero payroll dates arrive as /Date(ms+0000)/ or an ISO string. Returns yyyy-mm-dd or "".
function xeroDate(v) {
  if (!v) return "";
  if (/^\d{4}-\d{2}-\d{2}/.test(String(v))) return String(v).slice(0, 10);
  const m = /\/Date\((-?\d+)/.exec(String(v));
  const d = m ? new Date(Number(m[1])) : new Date(String(v));
  return isNaN(d) ? "" : d.toISOString().slice(0, 10);
}
// Emergency contact. Looks through every field on the record whose name mentions "emergency", whatever shape Xero uses
// (a list of contacts, one contact object, or flat fields such as EmergencyContactName), and sorts the values by field name.
function xeroEmergencyPairs(e) {
  const pairs = [];
  const take = (prefix, v) => {
    if (v === null || v === undefined || v === "") return;
    if (Array.isArray(v)) { if (v.length) take(prefix, v[0]); return; }
    if (typeof v === "object") { Object.keys(v).forEach((k) => take(`${prefix}.${k}`, v[k])); return; }
    pairs.push([prefix, String(v)]);
  };
  Object.keys(e || {}).forEach((k) => { if (/emergency/i.test(k)) take(k, e[k]); });
  return pairs;
}
function xeroEmergency(e) {
  const out = { emName: "", emRelationship: "", emMobile: "", emEmail: "" };
  const leaf = (key) => key.split(".").pop().replace(/^emergency(contact)?/i, "");
  for (const [key, val] of xeroEmergencyPairs(e)) {
    const k = leaf(key) + " " + key;
    const own = leaf(key);
    if (/relation/i.test(k)) out.emRelationship = out.emRelationship || val;
    else if (/mail/i.test(own) || /mail/i.test(k) && !/name/i.test(own)) out.emEmail = out.emEmail || val;
    else if (/phone|mobile|number|contactno/i.test(own) || /phone|mobile/i.test(k)) out.emMobile = out.emMobile || val;
    else if (/^(first)?name$|^$|^full/i.test(own) || /name/i.test(own)) out.emName = out.emName ? (/last|sur/i.test(own) ? `${out.emName} ${val}` : out.emName) : val;
  }
  return out;
}
// Xero's own answer, as plain text for the app to show: the status, the headers that explain a refusal, and the start of the body.
async function xeroWhy(res) {
  const h = [];
  for (const k of ["www-authenticate", "x-rate-limit-problem", "retry-after"]) { const v = res.headers.get(k); if (v) h.push(`${k}: ${v}`); }
  const body = (await res.text().catch(() => "")).replace(/\s+/g, " ").trim().slice(0, 240);
  return `Xero answered ${res.status}${h.length ? ` (${h.join("; ")})` : ""}${body ? `. ${body}` : ". It sent no message."}`;
}
// Runs fn over the items a few at a time. Xero allows only 5 calls in flight at once, so the employee records are fetched 3 at a time.
async function xeroEach(items, fn, limit = 3) {
  const out = new Array(items.length);
  let next = 0;
  await Promise.all(Array.from({ length: Math.min(limit, items.length) }, async () => {
    while (next < items.length) { const i = next++; out[i] = await fn(items[i], i); }
  }));
  return out;
}
// The Xero employee list, kept for 3 minutes so one screen opening does not ask Xero again and again. Only a good answer is kept.
let xeroListKept = null;
async function xeroEmployeeList(env) {
  if (xeroListKept && Date.now() - xeroListKept.at < 3 * 60000) return { list: xeroListKept.list };
  const res = await xeroFetch(env, "payroll.xro/1.0/Employees");
  if (!res.ok) return { fail: { status: res.status, detail: await xeroWhy(res) } };
  const list = (await res.json()).Employees || [];
  xeroListKept = { at: Date.now(), list };
  return { list };
}
function xeroStaffRow(e) {
  const a = e.HomeAddress || {};
  const street = [a.AddressLine1, a.AddressLine2].filter(Boolean).join(" ");
  const place = [a.City, a.Region, a.PostalCode].filter(Boolean).join(" ");
  const bank = (e.BankAccounts || [])[0] || {};
  return {
    id: e.EmployeeID,
    status: e.Status || "",
    name: [e.FirstName, e.LastName].filter(Boolean).join(" "),
    address: [street, place].filter(Boolean).join(", "),
    dob: xeroDate(e.DateOfBirth),
    mobile: e.Mobile || e.Phone || "",
    email: e.Email || "",
    accountName: bank.AccountName || "",
    bsb: bank.BSB || "",
    accountNumber: bank.AccountNumber || "",
    commenced: xeroDate(e.StartDate),
    title: String(e.JobTitle || "").trim(),
    ...xeroEmergency(e)
  };
}

// Employment basis (FULLTIME, PARTTIME, CASUAL, LABOURHIRE, SUPERINCOMESTREAM) from the Employment tab, held on the tax declaration. EmploymentType is only Employee or Contractor, so it is never used here.
function xeroBasis(record) {
  const td = record && record.TaxDeclaration || {};
  const v = td.EmploymentBasis || record.EmploymentBasis || "";
  return String(v).replace(/[\s_-]/g, "").toUpperCase();
}
// Hourly rates for ordinary time, time and a half and double time, worked from the employee's pay template and DFP's earnings rates.
// Which earnings rates are ordinary time, overtime under 2 hours and overtime over 2 hours. Shared by the pay figures and the timesheet push, so both pick the same rates.
function xeroPickRates(record, earningsRates) {
  const byId = new Map(earningsRates.map((r) => [r.EarningsRateID, r]));
  const lines = (record.PayTemplate && record.PayTemplate.EarningsLines) || [];
  const lineFor = (id) => lines.find((l) => l.EarningsRateID === id);
  const named = (test) => earningsRates.find((r) => test(String(r.Name || "")));
  const overtime = (n) => /overtime/i.test(n);
  const ordinaryItem = byId.get(record.OrdinaryEarningsRateID) || named((n) => /ordinary/i.test(n));
  const ot15Item = named((n) => overtime(n) && (/<\s*2/.test(n) || /(less than|under)\s*2/i.test(n) || /1\.5|time and a half/i.test(n)));
  const ot2Item = named((n) => overtime(n) && (/>\s*2/.test(n) || /(more than|over)\s*2/i.test(n) || /double/i.test(n)));
  return { byId, lines, lineFor, ordinaryItem, ot15Item, ot2Item };
}
function xeroRates(record, earningsRates) {
  const { byId, lines, lineFor, ordinaryItem, ot15Item, ot2Item } = xeroPickRates(record, earningsRates);
  const out = { ordinary: null, ot15: null, ot2: null, annual: null, weekHours: null };
  const why = [];
  const money = (n) => Math.round(Number(n) * 100) / 100;
  if (!ordinaryItem) why.push("No ordinary earnings rate found in Xero");
  else {
    const ln = lineFor(ordinaryItem.EarningsRateID);
    if (ln && ln.CalculationType === "ANNUALSALARY") {
      // An annual salary on the pay template: the yearly amount and the hours a week it is based on.
      const a = xnum(ln.AnnualSalary);
      const h = xnum(ln.NumberOfUnitsPerWeek) || xunits(ln.NormalNumberOfUnits);
      if (a > 0) { out.annual = money(a); if (h > 0) out.weekHours = Math.round(h * 100) / 100; }
      else why.push("Ordinary time is an annual salary in Xero but Xero gave no amount");
    } else if (ln && ln.RatePerUnit) out.ordinary = money(ln.RatePerUnit);
    else if (ordinaryItem.RateType === "RATEPERUNIT" && ordinaryItem.RatePerUnit) out.ordinary = money(ordinaryItem.RatePerUnit);
    else why.push("No hourly ordinary rate on the pay template");
  }
  const derive = (item, label) => {
    if (!item) { why.push(`No earnings rate found in Xero for ${label}`); return null; }
    const ln = lineFor(item.EarningsRateID);
    if (ln && ln.RatePerUnit) return money(ln.RatePerUnit);
    if (item.RateType === "RATEPERUNIT" && item.RatePerUnit) return money(item.RatePerUnit);
    const base = out.ordinary || (out.annual && out.weekHours ? out.annual / 52 / out.weekHours : null);
    if (item.RateType === "MULTIPLE" && item.Multiplier && base) return money(base * Number(item.Multiplier));
    why.push(`Could not work out ${label} from Xero`);
    return null;
  };
  // Overtime entitlement: the pay template carries a line for an overtime earnings rate.
  const hasOvertime = !!((ot15Item && lineFor(ot15Item.EarningsRateID)) || (ot2Item && lineFor(ot2Item.EarningsRateID)));
  if (hasOvertime) {
    out.ot15 = derive(ot15Item, "Overtime < 2 Hours");
    out.ot2 = derive(ot2Item, "Overtime > 2 Hours");
  }
  return { rates: out, why, hasOvertime, templateLines: lines.map((l) => (byId.get(l.EarningsRateID) || {}).Name || l.EarningsRateID), used: { ordinary: ordinaryItem && ordinaryItem.Name, ot15: ot15Item && ot15Item.Name, ot2: ot2Item && ot2Item.Name } };
}

// Super fund details for one Xero employee record: fund name, member number and USI. Never returns fund bank details.
// "kept" is optional. When a caller doing many people at once passes the same object each time, the super funds list and each fund product are read from Xero once instead of once per person.
async function xeroSuperFor(env, record, kept) {
  const ms = Array.isArray(record.SuperMemberships) ? record.SuperMemberships : [];
  if (!ms.length) return { memberships: [], why: "No super fund on the Xero employee record" };
  let funds = [];
  let why = "";
  if (kept && kept.funds) { funds = kept.funds; why = kept.why || ""; }
  else {
    try {
      const r = await xeroFetch(env, "payroll.xro/1.0/Superfunds");
      if (r.ok) funds = (await r.json()).SuperFunds || [];
      else { why = `Xero refused the super funds list (${r.status})`; console.log("xero superfunds failed", r.status, (await r.text()).slice(0, 200)); }
    } catch (e) { why = "Could not read the super funds list from Xero"; }
    if (kept) { kept.funds = funds; kept.why = why; kept.products = {}; }
  }
  const out = [];
  for (const m of ms) {
    const f = funds.find((x) => x.SuperFundID === m.SuperFundID) || {};
    const usi = String(f.USI || "").trim();
    let name = String(f.Name || "").trim();
    if (!name && usi) {
      if (kept && kept.products && usi in kept.products) name = kept.products[usi];
      else {
        try {
          const r = await xeroFetch(env, `payroll.xro/1.0/SuperfundProducts?USI=${encodeURIComponent(usi)}`);
          if (r.ok) { const prod = ((await r.json()).SuperFundProducts || [])[0]; if (prod && prod.ProductName) name = String(prod.ProductName).trim(); }
        } catch (e) { /* leave the name blank */ }
        if (kept && kept.products) kept.products[usi] = name;
      }
    }
    if (!name && f.Type === "SMSF") name = "Self managed super fund";
    out.push({ fundName: name, memberNumber: String(m.EmployeeNumber || "").trim(), usi });
  }
  return { memberships: out, why };
}

// ---- Pay slips built from Xero pay runs by the Pull Pay Slips action, and stored. Opening a pay slip never calls Xero. ----
const xnum = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };
const xunits = (v) => Array.isArray(v) ? v.reduce((a, x) => a + xnum(x), 0) : xnum(v);
// Every POSTED pay run, newest payment date first.
let xeroRunsKept = null;
async function xeroPostedRuns(env) {
  if (xeroRunsKept && Date.now() - xeroRunsKept.at < 60000) return { runs: xeroRunsKept.runs };
  let runs = [];
  for (let page = 1; page <= 5; page++) {
    const r = await xeroFetch(env, `payroll.xro/1.0/PayRuns?page=${page}`);
    if (!r.ok) {
      const detail = (await r.text()).slice(0, 300);
      console.log("xero payruns failed", r.status, detail);
      return { fail: { error: r.status === 401 || r.status === 403 ? "Xero has not allowed pay run access yet. Reconnect Xero on the Xero card." : "Xero refused the pay runs", status: r.status, detail } };
    }
    const batch = (await r.json()).PayRuns || [];
    runs = runs.concat(batch);
    if (batch.length < 100) break;
  }
  runs = runs.filter((x) => (x.PayRunStatus || "") === "POSTED").sort((a, b) => xeroDate(b.PaymentDate).localeCompare(xeroDate(a.PaymentDate)));
  xeroRunsKept = { at: Date.now(), runs };
  return { runs };
}
// A figure Xero gave, as a number, or null when Xero gave none at all. A zero is only taken when no other source holds a real figure,
// because a pay run summary may carry "not given" as zero. Nothing is worked out here: no sums, no substitutes.
const xeroFigure = (...vals) => {
  let sawZero = false;
  for (const v of vals) {
    if (v === null || v === undefined || v === "") continue;
    const n = Number(v);
    if (!Number.isFinite(n)) continue;
    if (n !== 0) return money2(n);
    sawZero = true;
  }
  return sawZero ? 0 : null;
};
// The totals for one payslip exactly as Xero gives them: from the full payslip first, then the pay run's summary row. A figure Xero does not give stays null.
function slipTotals(summary, slip) {
  const sm = summary || {};
  const sl = slip || {};
  return {
    gross: xeroFigure(sl.Wages, sm.Wages),
    tax: xeroFigure(sl.Tax, sm.Tax),
    super: xeroFigure(sl.Super, sm.Super),
    deductions: xeroFigure(sl.Deductions, sm.Deductions),
    reimbursements: xeroFigure(sl.Reimbursements, sm.Reimbursements),
    net: xeroFigure(sl.NetPay, sm.NetPay)
  };
}
// Turns one Xero payslip into the shape the app prints. Every figure comes from Xero; nothing is worked out here except merging identical lines.
function xeroSlipShape(slip, run, record, items, supMemberships, fundOf, tot) {
  const byId = (list, key) => new Map((list || []).map((x) => [x[key], x]));
  const earn = byId(items.EarningsRates, "EarningsRateID");
  const ded = byId(items.DeductionTypes, "DeductionTypeID");
  const reimb = byId(items.ReimbursementTypes, "ReimbursementTypeID");
  const merged = new Map();
  const addLine = (l, fallback) => {
    const hours = xunits(l.NumberOfUnits);
    const rate = xnum(l.RatePerUnit);
    const fixed = xnum(l.FixedAmount) || xnum(l.Amount);
    const amount = money2(fixed || hours * rate);
    const label = txt((earn.get(l.EarningsRateID) || {}).Name, 80) || fallback;
    const key = `${label}|${rate}|${fixed ? "f" : "h"}`;
    const prev = merged.get(key);
    if (prev) { prev.hours = money2(prev.hours + hours); prev.amount = money2(prev.amount + amount); }
    else merged.set(key, { label, hours: fixed ? 0 : money2(hours), rate: fixed ? 0 : money2(rate), amount });
  };
  (slip.EarningsLines || []).forEach((l) => addLine(l, "Earnings"));
  (slip.TimesheetEarningsLines || []).forEach((l) => addLine(l, "Timesheet earnings"));
  (slip.LeaveEarningsLines || []).forEach((l) => addLine(l, "Leave"));
  const lines = [...merged.values()].filter((l) => l.amount !== 0);
  const taxLines = (slip.TaxLines || []).map((l) => ({ label: (/tax.?free|leave loading|threshold/i.test(String(l.Description || "")) ? "" : txt(l.Description, 80)) || txt(l.TaxTypeName, 80) || "Tax", amount: money2(l.Amount) })).filter((l) => l.amount !== 0);
  const deductions = (slip.DeductionLines || []).map((l) => ({ label: txt((ded.get(l.DeductionTypeID) || {}).Name, 80) || "Deduction", amount: money2(l.Amount) })).filter((l) => l.amount !== 0);
  const reimbursements = (slip.ReimbursementLines || []).map((l) => ({ label: txt(l.Description || (reimb.get(l.ReimbursementTypeID) || {}).Name, 80) || "Reimbursement", amount: money2(l.Amount) })).filter((l) => l.amount !== 0);
  const SUPER_LABEL = { SGC: "Superannuation Guarantee (SGC)", SALARYSACRIFICE: "Salary Sacrifice", EMPLOYERADDITIONAL: "Employer Additional", EMPLOYEE: "Employee Contribution" };
  const superOf = (l) => {
    // A line is matched to its fund by the membership Xero names. With no match the fund is only taken when the person has just one, otherwise it is left blank, never guessed.
    const ms = supMemberships || [];
    const m = ms.find((x) => x.SuperMembershipID === l.SuperMembershipID) || (ms.length === 1 ? ms[0] : null);
    return (m && fundOf(m)) || {};
  };
  const superLines = (slip.SuperannuationLines || []).map((l) => {
    const f = superOf(l);
    const key = String(l.ContributionType || "").replace(/[\s_-]/g, "").toUpperCase();
    return { label: SUPER_LABEL[key] || txt(l.ContributionType, 60).replace(/([a-z])([A-Z])/g, "$1 $2") || "Superannuation", amount: money2(l.Amount), fund: f.fundName || "", member: f.memberNumber || "", usi: f.usi || "" };
  }).filter((l) => l.amount !== 0);
  const gross = tot.gross;
  const warnings = [];
  const lineTotal = money2(lines.reduce((a, l) => a + l.amount, 0));
  if (gross !== null && Math.abs(lineTotal - gross) > 0.02) warnings.push(`The earnings lines Xero returned add to ${lineTotal.toFixed(2)} but Xero reports total earnings of ${gross.toFixed(2)}.`);
  const taxTotal = money2(taxLines.reduce((a, l) => a + l.amount, 0));
  if (tot.tax !== null && Math.abs(taxTotal - tot.tax) > 0.02) warnings.push(`The tax lines Xero returned add to ${taxTotal.toFixed(2)} but Xero reports tax of ${tot.tax.toFixed(2)}.`);
  const start = xeroDate(run.PayRunPeriodStartDate);
  const end = xeroDate(run.PayRunPeriodEndDate);
  const days = start && end ? Math.round((new Date(end) - new Date(start)) / 864e5) + 1 : 0;
  const ha = record.HomeAddress || {};
  const addr = [ha.AddressLine1, ha.AddressLine2, [ha.City, ha.Region, ha.PostalCode].filter(Boolean).join(" ")].map((x) => txt(x, 80)).filter(Boolean).join(", ");
  const banks = Array.isArray(record.BankAccounts) ? record.BankAccounts : [];
  const bank = banks.find((b) => b.Remainder) || banks[0] || {};
  const acct = txt(bank.AccountNumber, 20).replace(/\s+/g, "");
  const supTotal = tot.super;
  // The fund shown on the slip is the one Xero names on the super lines. If the lines name different funds, none is shown at the top (each line keeps its own). With no lines it is taken only when the person has a single membership.
  const named = superLines.filter((l) => l.fund || l.member);
  const sameFund = named.length && named.every((l) => l.fund === named[0].fund && l.member === named[0].member);
  const firstFund = sameFund ? named[0] : named.length ? { fund: "", member: "", usi: "" } : (() => { const f = superOf({}); return { fund: f.fundName || "", member: f.memberNumber || "", usi: f.usi || "" }; })();
  if (!supTotal && (slip.SuperannuationLines || []).length) warnings.push("Xero returned a superannuation line for this pay but no superannuation amount.");
  // Anything Xero did not give is listed here and shown as "Not supplied" on the pay slip. Nothing is worked out in its place.
  const missing = [["total earnings", tot.gross], ["tax", tot.tax], ["superannuation", tot.super], ["net pay", tot.net]].filter((x) => x[1] === null).map((x) => x[0]);
  missing.forEach((x) => warnings.push(`Xero did not supply ${x} for this pay.`));
  return {
    employee: { name: txt(`${record.FirstName || ""} ${record.LastName || ""}`, 120), address: addr, frequency: days === 7 ? "Weekly" : days === 14 ? "Fortnightly" : days >= 28 ? "Monthly" : "", employmentType: xeroBasis(record) },
    period: { start, end }, paymentDate: xeroDate(run.PaymentDate),
    lines, gross, taxLines, tax: tot.tax, deductions, deductionsTotal: tot.deductions, reimbursements, reimbursementsTotal: tot.reimbursements, net: tot.net,
    super: { amount: supTotal, lines: superLines.map((l) => ({ label: l.label, amount: l.amount })), fund: firstFund.fund || "", member: firstFund.member || "", usi: firstFund.usi || "" },
    bank: { accountName: txt(bank.AccountName, 120), bsb: txt(bank.BSB, 10), masked: acct ? `XXXX ${acct.slice(-4)}` : "" },
    reference: txt(bank.StatementText, 80) || `Wages ${end.slice(8, 10)}.${end.slice(5, 7)}.${end.slice(0, 4)}`,
    warnings, missing
  };
}

// One pay run with every person's summary row. Posted runs are kept in memory for five minutes, so a run is read from Xero once however many people are filed.
const xeroRunCache = new Map();
async function xeroRunDetail(env, runId) {
  const kept = xeroRunCache.get(runId);
  if (kept && Date.now() - kept.at < 5 * 60000) return kept.res;
  const r = await xeroFetch(env, `payroll.xro/1.0/PayRuns/${runId}`);
  if (!r.ok) {
    const detail = (await r.text()).slice(0, 300);
    console.log("xero payrun failed", r.status, detail);
    return { fail: { status: 502, body: { error: r.status === 401 || r.status === 403 ? "Xero has not allowed pay run access yet. Reconnect Xero on the Xero card." : "Xero refused the pay run", status: r.status, detail } } };
  }
  const run = ((await r.json()).PayRuns || [])[0];
  const res = { run };
  if (run && (run.PayRunStatus || "") === "POSTED") xeroRunCache.set(runId, { at: Date.now(), res });
  return res;
}
let xeroItemsKept = null;
async function xeroItems(env) {
  if (xeroItemsKept && Date.now() - xeroItemsKept.at < 10 * 60000) return xeroItemsKept.items;
  try {
    const pi = await xeroFetch(env, "payroll.xro/1.0/PayItems");
    if (pi.ok) { const items = (await pi.json()).PayItems || {}; xeroItemsKept = { at: Date.now(), items }; return items; }
  } catch (e) { /* lines fall back to generic names */ }
  return {};
}
// ---- Filing pay slip PDFs into each person's staff folder in SharePoint ----
// Reservoir > Workplace Culture > Employee Records > Current Employees > <person> > Payslips
const STAFF_PATHS = ["Workplace Culture/Employee Records/Current Employees", "Reservoir/Workplace Culture/Employee Records/Current Employees"];
let staffRootKept = null;
async function staffRoot(env) {
  if (staffRootKept && Date.now() - staffRootKept.at < 10 * 60000) return { root: staffRootKept.root };
  const bases = [`/sites/${siteId(env)}/drive`];
  try {
    const dl = await graphFetch(env, `/sites/${siteId(env)}/drives`);
    if (dl.ok) for (const d of (await dl.json()).value || []) bases.push(`/drives/${d.id}`);
  } catch (e) { /* the default library only */ }
  for (const base of bases) {
    for (const path of STAFF_PATHS) {
      const r = await graphFetch(env, `${base}/root:/${pathOf(path.split("/"))}`);
      if (!r.ok) continue;
      const it = await r.json();
      if (!it.folder) continue;
      const root = { driveBase: `/drives/${it.parentReference && it.parentReference.driveId}`, id: it.id, path };
      staffRootKept = { at: Date.now(), root };
      return { root };
    }
  }
  let top = [];
  try {
    const t = await graphFetch(env, `/sites/${siteId(env)}/drive/root/children?$select=name&$top=60`);
    if (t.ok) top = ((await t.json()).value || []).map((x) => x.name);
  } catch (e) { /* none to show */ }
  return { why: `Could not find ${STAFF_PATHS[0]} in the Reservoir site.${top.length ? ` The top level folders there are: ${top.join(", ")}.` : ""}` };
}
async function driveChildren(env, driveBase, id) {
  let url = `${driveBase}/items/${id}/children?$select=id,name,folder&$top=200`;
  const out = [];
  for (let i = 0; i < 10 && url; i++) {
    const r = await graphFetch(env, url);
    if (!r.ok) throw new Error("folder list " + r.status);
    const j = await r.json();
    (j.value || []).forEach((x) => out.push(x));
    url = j["@odata.nextLink"] || "";
  }
  return out;
}
const nameTokens = (v) => String(v || "").toLowerCase().replace(/['\u2018\u2019`\u00b4]/g, "").replace(/[^a-z0-9 ]/g, " ").split(/\s+/).filter(Boolean);
// Finds one person's folder by name. Order and punctuation do not matter. Returns { folder } or { why }.
function matchStaffFolder(folders, first, last) {
  const f = nameTokens(first), l = nameTokens(last);
  const want = new Set([...f, ...l]);
  const dirs = folders.filter((x) => x.folder);
  const exact = dirs.filter((x) => { const t = nameTokens(x.name); return t.length === want.size && t.every((y) => want.has(y)); });
  if (exact.length === 1) return { folder: exact[0] };
  if (exact.length > 1) return { why: "More than one folder matches that name" };
  const loose = dirs.filter((x) => { const t = new Set(nameTokens(x.name)); return f.concat(l).length > 0 && f.every((y) => t.has(y)) && l.every((y) => t.has(y)); });
  if (loose.length === 1) return { folder: loose[0] };
  return { why: loose.length ? "More than one folder matches that name" : "No folder with that name under Current Employees" };
}
// ---- PDF scan start. Reads the text of an employment contract PDF so the office does not have to type the dates. Best effort: scanned images and unusual fonts come back unreadable and the office enters the dates by hand. ----
const pdfBin = (bytes) => { let s = ""; const n = bytes.length; for (let i = 0; i < n; i += 8192) s += String.fromCharCode.apply(null, bytes.subarray(i, Math.min(n, i + 8192))); return s; };
const pdfBytes = (s) => { const b = new Uint8Array(s.length); for (let i = 0; i < s.length; i++) b[i] = s.charCodeAt(i) & 255; return b; };
async function pdfInflate(bytes) {
  for (const cut of [0, 1, 2]) {
    try {
      const ds = new DecompressionStream("deflate");
      const w = ds.writable.getWriter();
      w.write(bytes.subarray(0, bytes.length - cut)).catch(() => {});
      w.close().catch(() => {});
      const out = await new Response(ds.readable).arrayBuffer();
      return new Uint8Array(out);
    } catch (e) { /* try with the trailing line break removed */ }
  }
  return null;
}
// The text between a matched pair of << >> that starts at index i of s (s[i] is the first "<").
function pdfDictAt(s, i) {
  let depth = 0;
  for (let j = i; j < s.length - 1; j++) {
    if (s[j] === "<" && s[j + 1] === "<") { depth++; j++; }
    else if (s[j] === ">" && s[j + 1] === ">") { depth--; j++; if (depth === 0) return s.slice(i, j + 1); }
  }
  return "";
}
// The value that follows /Key in a dictionary text: a dictionary, an array, a reference, or a bare token.
function pdfGet(dict, key) {
  const m = new RegExp(`/${key}(?![A-Za-z0-9])\\s*`).exec(dict);
  if (!m) return "";
  const rest = dict.slice(m.index + m[0].length);
  if (rest.startsWith("<<")) return pdfDictAt(rest, 0);
  if (rest[0] === "[") { const e = rest.indexOf("]"); return e < 0 ? "" : rest.slice(0, e + 1); }
  const r = /^(\d+)\s+(\d+)\s+R/.exec(rest);
  if (r) return `${r[1]} ${r[2]} R`;
  const t = /^[^\s/<>\[\]]+/.exec(rest);
  return t ? t[0] : "";
}
const pdfRef = (v) => { const m = /^(\d+)\s+\d+\s+R/.exec(String(v || "").trim()); return m ? Number(m[1]) : 0; };
function pdfAscii85(bytes) {
  const out = [];
  let tuple = [], i = 0;
  const s = pdfBin(bytes);
  while (i < s.length) {
    const ch = s[i++];
    if (ch === "~") break;
    if (/\s/.test(ch)) continue;
    if (ch === "z" && tuple.length === 0) { out.push(0, 0, 0, 0); continue; }
    const v = ch.charCodeAt(0) - 33;
    if (v < 0 || v > 84) continue;
    tuple.push(v);
    if (tuple.length === 5) {
      let n = 0;
      for (const t of tuple) n = n * 85 + t;
      out.push((n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255);
      tuple = [];
    }
  }
  if (tuple.length > 1) {
    const keep = tuple.length - 1;
    while (tuple.length < 5) tuple.push(84);
    let n = 0;
    for (const t of tuple) n = n * 85 + t;
    const bs = [(n >>> 24) & 255, (n >>> 16) & 255, (n >>> 8) & 255, n & 255];
    out.push(...bs.slice(0, keep));
  }
  return new Uint8Array(out);
}
function pdfAsciiHex(bytes) {
  const s = pdfBin(bytes).split(">")[0].replace(/\s+/g, "");
  const out = [];
  for (let i = 0; i < s.length; i += 2) out.push(parseInt(s.slice(i, i + 2).padEnd(2, "0"), 16));
  return new Uint8Array(out);
}
async function pdfObjects(raw) {
  const objs = new Map();
  const re = /(\d+)\s+(\d+)\s+obj\b/g;
  let m;
  while ((m = re.exec(raw))) {
    const num = Number(m[1]);
    const bodyStart = m.index + m[0].length;
    const sIdx = raw.indexOf("stream", bodyStart);
    const eIdx = raw.indexOf("endobj", bodyStart);
    let dictText, stream = null;
    if (sIdx >= 0 && (eIdx < 0 || sIdx < eIdx) && /^stream\r?\n/.test(raw.slice(sIdx, sIdx + 8)) && /^\s*<</.test(raw.slice(bodyStart, bodyStart + 6))) {
      dictText = raw.slice(bodyStart, sIdx);
      const dataStart = sIdx + (raw[sIdx + 6] === "\r" ? 8 : 7);
      let len = Number(pdfGet(dictText, "Length"));
      if (!Number.isFinite(len)) {
        const lr = /^(\d+)\s+\d+\s+R/.exec(pdfGet(dictText, "Length"));
        const lm = lr ? new RegExp(`(?:^|[\\r\\n])${lr[1]}\\s+\\d+\\s+obj\\s+(\\d+)\\s+endobj`).exec(raw) : null;
        len = lm ? Number(lm[1]) : NaN;
      }
      let dataEnd = Number.isFinite(len) && len > 0 && raw.slice(dataStart + len, dataStart + len + 12).trim().startsWith("endstream") ? dataStart + len : raw.indexOf("endstream", dataStart);
      if (!(dataEnd > dataStart)) dataEnd = raw.indexOf("endstream", dataStart);
      if (dataEnd < 0) dataEnd = dataStart;
      stream = { start: dataStart, end: dataEnd };
      re.lastIndex = Math.max(re.lastIndex, dataEnd);
    } else {
      dictText = eIdx >= 0 ? raw.slice(bodyStart, eIdx) : raw.slice(bodyStart, bodyStart + 2000);
    }
    objs.set(num, { dict: dictText, stream });
  }
  const dataOf = async (o) => {
    if (!o || !o.stream) return null;
    let bytes = pdfBytes(raw.slice(o.stream.start, o.stream.end));
    const fm = /\/Filter\s*(\[[^\]]*\]|\/[A-Za-z0-9]+)/.exec(o.dict);
    const filt = fm ? fm[1] : "";
    const names = [...filt.matchAll(/\/([A-Za-z0-9]+)/g)].map((x) => x[1]);
    for (const nm of names) {
      if (!bytes) return null;
      if (nm === "FlateDecode" || nm === "Fl") bytes = await pdfInflate(bytes);
      else if (nm === "ASCII85Decode" || nm === "A85") bytes = pdfAscii85(bytes);
      else if (nm === "ASCIIHexDecode" || nm === "AHx") bytes = pdfAsciiHex(bytes);
      else return null;
    }
    return bytes;
  };
  // Objects packed inside object streams (PDF 1.5 and later).
  for (const [num, o] of [...objs]) {
    if (!/\/Type\s*\/ObjStm/.test(o.dict)) continue;
    const bytes = await dataOf(o);
    if (!bytes) continue;
    const txt = pdfBin(bytes);
    const n = Number(pdfGet(o.dict, "N")), first = Number(pdfGet(o.dict, "First"));
    if (!(n > 0) || !(first > 0)) continue;
    const head = txt.slice(0, first).trim().split(/\s+/).map(Number);
    for (let k = 0; k < n; k++) {
      const id = head[k * 2], off = head[k * 2 + 1];
      const end = k + 1 < n ? head[(k + 1) * 2 + 1] : txt.length - first;
      if (Number.isFinite(id) && Number.isFinite(off)) objs.set(id, { dict: txt.slice(first + off, first + end), stream: null, packed: true });
    }
  }
  return { objs, dataOf };
}
function pdfHexToStr(hex) {
  const h = hex.replace(/\s+/g, "");
  let out = "";
  for (let i = 0; i + 3 < h.length + 1; i += 4) { const c = parseInt(h.slice(i, i + 4), 16); if (!Number.isNaN(c)) out += String.fromCharCode(c); }
  return out;
}
function pdfToUnicode(text) {
  const map = new Map();
  let width = 1;
  const cs = /begincodespacerange([\s\S]*?)endcodespacerange/.exec(text);
  if (cs) { const h = /<([0-9a-fA-F]+)>/.exec(cs[1]); if (h) width = Math.max(1, Math.round(h[1].length / 2)); }
  let m;
  const chars = /beginbfchar([\s\S]*?)endbfchar/g;
  while ((m = chars.exec(text))) {
    const pairs = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>/g;
    let p;
    while ((p = pairs.exec(m[1]))) map.set(parseInt(p[1], 16), pdfHexToStr(p[2]));
  }
  const ranges = /beginbfrange([\s\S]*?)endbfrange/g;
  while ((m = ranges.exec(text))) {
    const items = /<([0-9a-fA-F]+)>\s*<([0-9a-fA-F]+)>\s*(\[[^\]]*\]|<[0-9a-fA-F]+>)/g;
    let p;
    while ((p = items.exec(m[1]))) {
      const lo = parseInt(p[1], 16), hi = parseInt(p[2], 16);
      if (hi - lo > 70000) continue;
      if (p[3][0] === "[") {
        const arr = [...p[3].matchAll(/<([0-9a-fA-F]+)>/g)].map((x) => pdfHexToStr(x[1]));
        for (let c = lo; c <= hi && c - lo < arr.length; c++) map.set(c, arr[c - lo]);
      } else {
        const base = pdfHexToStr(p[3].slice(1, -1));
        const last = base.length ? base.charCodeAt(base.length - 1) : 0;
        for (let c = lo; c <= hi; c++) map.set(c, base.slice(0, -1) + String.fromCharCode(last + (c - lo)));
      }
    }
  }
  return { map, width };
}
// Splits a content stream into tokens: strings (as { s } or { h }), names, numbers, arrays and operators.
function pdfTokens(c) {
  const out = [];
  let i = 0;
  const n = c.length;
  const isWs = (ch) => ch === " " || ch === "\n" || ch === "\r" || ch === "\t" || ch === "\f" || ch === "\0";
  const isDelim = (ch) => "()<>[]{}/%".includes(ch);
  const readArray = () => {
    const arr = [];
    i++;
    while (i < n && c[i] !== "]") {
      const t = next();
      if (t === undefined) break;
      arr.push(t);
    }
    i++;
    return { a: arr };
  };
  const next = () => {
    while (i < n && isWs(c[i])) i++;
    if (i >= n) return undefined;
    const ch = c[i];
    if (ch === "%") { while (i < n && c[i] !== "\n" && c[i] !== "\r") i++; return next(); }
    if (ch === "(") {
      let depth = 1, s = "";
      i++;
      while (i < n && depth > 0) {
        const d = c[i];
        if (d === "\\") {
          const e = c[i + 1];
          if (e >= "0" && e <= "7") { let o = e; i += 2; while (o.length < 3 && c[i] >= "0" && c[i] <= "7") { o += c[i]; i++; } s += String.fromCharCode(parseInt(o, 8) & 255); continue; }
          if (e === "n") s += "\n"; else if (e === "r") s += "\r"; else if (e === "t") s += "\t"; else if (e === "b") s += "\b"; else if (e === "f") s += "\f";
          else if (e === "\r") { if (c[i + 2] === "\n") i++; } else if (e === "\n") { /* line continuation */ } else s += e;
          i += 2;
          continue;
        }
        if (d === "(") depth++;
        else if (d === ")") { depth--; if (depth === 0) { i++; break; } }
        s += d;
        i++;
      }
      return { s };
    }
    if (ch === "<" && c[i + 1] === "<") { const d = pdfDictAt(c, i); i += Math.max(2, d.length); return { d: true }; }
    if (ch === "<") { const e = c.indexOf(">", i); const h = c.slice(i + 1, e < 0 ? n : e).replace(/\s+/g, ""); i = e < 0 ? n : e + 1; return { h }; }
    if (ch === "[") return readArray();
    if (ch === "/") { let j = i + 1; while (j < n && !isWs(c[j]) && !isDelim(c[j])) j++; const name = c.slice(i + 1, j); i = j; return { n: name }; }
    if (ch === "]" || ch === ">" || ch === ")" || ch === "{" || ch === "}") { i++; return next(); }
    let j = i;
    while (j < n && !isWs(c[j]) && !isDelim(c[j])) j++;
    const word = c.slice(i, j);
    i = j > i ? j : i + 1;
    const num = Number(word);
    return word !== "" && Number.isFinite(num) ? { v: num } : { op: word };
  };
  let t;
  while ((t = next()) !== undefined) out.push(t);
  return out;
}
function pdfShow(tok, font) {
  const bytes = tok.h !== undefined ? (() => { const h = tok.h.length % 2 ? tok.h + "0" : tok.h; const b = []; for (let k = 0; k < h.length; k += 2) b.push(parseInt(h.slice(k, k + 2), 16)); return b; })() : [...tok.s].map((x) => x.charCodeAt(0) & 255);
  if (font && font.map && font.map.size) {
    let out = "";
    const w = font.width;
    for (let k = 0; k + w <= bytes.length; k += w) {
      let code = 0;
      for (let q = 0; q < w; q++) code = code * 256 + bytes[k + q];
      const got = font.map.get(code);
      out += got !== undefined ? got : w === 1 ? String.fromCharCode(code) : "";
    }
    return out;
  }
  if (font && font.twoByte) return "";
  return bytes.map((b) => String.fromCharCode(b)).join("");
}
function pdfPageText(tokens, fonts) {
  let out = "", font = null, y = null, lineY = null, x = 0;
  const stack = [];
  const newline = () => { if (!out.endsWith("\n")) out += "\n"; };
  for (let k = 0; k < tokens.length; k++) {
    const t = tokens[k];
    if (t.op === undefined) { stack.push(t); continue; }
    const op = t.op;
    const nums = stack.filter((s) => s.v !== undefined).map((s) => s.v);
    if (op === "Tf") { const nm = stack.find((s) => s.n !== undefined); font = nm ? fonts[nm.n] || null : null; }
    else if (op === "Tj" || op === "'" || op === '"') {
      if (op !== "Tj") newline();
      const s = stack.filter((q) => q.s !== undefined || q.h !== undefined).pop();
      if (s) out += pdfShow(s, font);
    } else if (op === "TJ") {
      const arr = stack.find((q) => q.a);
      if (arr) for (const item of arr.a) { if (item.v !== undefined) { if (item.v < -180) out += " "; } else if (item.s !== undefined || item.h !== undefined) out += pdfShow(item, font); }
    } else if (op === "Td" || op === "TD") { if (nums.length >= 2 && Math.abs(nums[1]) > 0.5) newline(); else if (nums.length >= 2 && nums[0] > 0.5) out += " "; }
    else if (op === "T*") newline();
    else if (op === "Tm") { if (nums.length >= 6) { if (lineY !== null && Math.abs(nums[5] - lineY) > 0.5) newline(); else if (lineY !== null && nums[4] - x > 0.5) out += " "; lineY = nums[5]; x = nums[4]; } }
    else if (op === "BT") { /* keep position between blocks */ }
    else if (op === "ET") { out += " "; }
    stack.length = 0;
  }
  return out;
}
async function pdfText(bytes) {
  if (!bytes || bytes.length < 20) return { text: "", pages: 0 };
  const raw = pdfBin(bytes);
  if (!raw.startsWith("%PDF")) return { text: "", pages: 0 };
  if (/\/Encrypt\b/.test(raw.slice(-2048)) || /\/Encrypt\s*\d/.test(raw.slice(-4096))) return { text: "", pages: 0, encrypted: true };
  const { objs, dataOf } = await pdfObjects(raw);
  const fontCache = new Map();
  const loadFont = async (num) => {
    if (fontCache.has(num)) return fontCache.get(num);
    const o = objs.get(num);
    let f = { map: null, width: 1, twoByte: false };
    if (o) {
      const tu = pdfRef(pdfGet(o.dict, "ToUnicode"));
      f.twoByte = /\/Identity-H|\/Type0/.test(o.dict);
      if (tu && objs.get(tu)) {
        const b = await dataOf(objs.get(tu));
        if (b) { const parsed = pdfToUnicode(pdfBin(b)); f.map = parsed.map; f.width = parsed.width; }
      }
    }
    fontCache.set(num, f);
    return f;
  };
  const resolveDict = (v) => { const r = pdfRef(v); return r && objs.get(r) ? objs.get(r).dict : String(v || ""); };
  const pages = [...objs.entries()].filter(([, o]) => /\/Type\s*\/Page(?![A-Za-z])/.test(o.dict)).sort((a, b) => a[0] - b[0]);
  const inherited = (o, key) => {
    let cur = o, guard = 0;
    while (cur && guard++ < 12) { const v = pdfGet(cur.dict, key); if (v) return v; const p = pdfRef(pdfGet(cur.dict, "Parent")); cur = p ? objs.get(p) : null; }
    return "";
  };
  let text = "";
  for (const [, page] of pages) {
    const res = resolveDict(inherited(page, "Resources"));
    const fontDict = resolveDict(pdfGet(res, "Font"));
    const fonts = {};
    for (const m of fontDict.matchAll(/\/([^\s/<>\[\]()]+)\s+(\d+)\s+\d+\s+R/g)) fonts[m[1]] = await loadFont(Number(m[2]));
    const contents = pdfGet(page.dict, "Contents");
    const refs = contents.startsWith("[") ? [...contents.matchAll(/(\d+)\s+\d+\s+R/g)].map((x) => Number(x[1])) : [pdfRef(contents)];
    let content = "";
    for (const r of refs) { const o = objs.get(r); const b = o && await dataOf(o); if (b) content += pdfBin(b) + "\n"; }
    text += pdfPageText(pdfTokens(content), fonts) + "\n\f\n";
  }
  return { text, pages: pages.length };
}
const PDF_MONTHS = { jan: 1, january: 1, feb: 2, february: 2, mar: 3, march: 3, apr: 4, april: 4, may: 5, jun: 6, june: 6, jul: 7, july: 7, aug: 8, august: 8, sep: 9, sept: 9, september: 9, oct: 10, october: 10, nov: 11, november: 11, dec: 12, december: 12 };
function pdfDates(text) {
  const found = [];
  const add = (index, length, y, m, d) => {
    if (y < 100) y += 2000;
    if (y < 1990 || y > 2100 || m < 1 || m > 12 || d < 1 || d > 31) return;
    const dt = new Date(Date.UTC(y, m - 1, d));
    if (dt.getUTCMonth() !== m - 1) return;
    found.push({ index, end: index + length, iso: `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}` });
  };
  let m;
  const dmyText = /\b(\d{1,2})(?:st|nd|rd|th)?\s+(?:of\s+)?([A-Za-z]{3,9})\.?,?\s+(\d{4})\b/g;
  while ((m = dmyText.exec(text))) { const mo = PDF_MONTHS[m[2].toLowerCase()]; if (mo) add(m.index, m[0].length, Number(m[3]), mo, Number(m[1])); }
  const mdyText = /\b([A-Za-z]{3,9})\.?\s+(\d{1,2})(?:st|nd|rd|th)?,?\s+(\d{4})\b/g;
  while ((m = mdyText.exec(text))) { const mo = PDF_MONTHS[m[1].toLowerCase()]; if (mo) add(m.index, m[0].length, Number(m[3]), mo, Number(m[2])); }
  const num = /\b(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2}|\d{4})\b/g;
  while ((m = num.exec(text))) add(m.index, m[0].length, Number(m[3]), Number(m[2]), Number(m[1]));
  const iso = /\b(\d{4})-(\d{2})-(\d{2})\b/g;
  while ((m = iso.exec(text))) add(m.index, m[0].length, Number(m[1]), Number(m[2]), Number(m[3]));
  found.sort((a, b) => a.index - b.index);
  return found.filter((d, i) => !found.slice(0, i).some((p) => p.index < d.end && d.index < p.end));
}
// Pulls the employee name check, the commencement date and the signed date out of the contract text. A hand signed date is usually blank, so the signed date is often empty.
function pdfContractFacts(text, personName) {
  const flat = String(text || "").replace(/[\f\r]+/g, "\n").replace(/[ \t]+/g, " ");
  const out = { readable: flat.replace(/\s+/g, "").length > 80, nameFound: false, commenced: "", signed: "" };
  if (!out.readable) return out;
  const lower = flat.toLowerCase();
  const parts = String(personName || "").toLowerCase().split(/\s+/).filter((x) => x.length > 1);
  out.nameFound = parts.length > 0 && parts.every((p) => lower.includes(p));
  const dates = pdfDates(flat);
  const key = /(commenc|start(?:ing)?\s+date|date\s+of\s+(?:start|commencement)|begin(?:s|ning)?|effective\s+(?:from|date)|employment\s+will\s+start)/gi;
  let hit = null, k;
  while ((k = key.exec(flat))) {
    const after = dates.find((d) => d.index >= k.index && d.index - k.index < 160);
    if (after) { hit = after; break; }
  }
  if (hit) out.commenced = hit.iso;
  const sign = /(signed|signature|date\s+signed|dated)/gi;
  let lastSign = -1;
  while ((k = sign.exec(flat))) lastSign = k.index;
  if (lastSign > flat.length * 0.5) {
    const tail = dates.filter((d) => d.index > lastSign - 40 && (!hit || d.index !== hit.index));
    if (tail.length) out.signed = tail[tail.length - 1].iso;
  }
  return out;
}
// ---- PDF scan end ----

// ---- Employment documents: the contract, TFN declaration, super form and any extras, filed by the director into each person's staff folder ----
// Reservoir > Workplace Culture > Employee Records > Current Employees > <person> > Employment Documents
const STAFF_DOC_SLOTS = { contract: "Original Employment Contract", tfn: "TFN Declaration", super: "Superannuation Form", additional: "Additional Documents" };
let staffDocsMade = false;
async function staffDocsTable(env) {
  if (staffDocsMade) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS staff_docs (id TEXT PRIMARY KEY, person_email TEXT NOT NULL, slot TEXT NOT NULL, drive_base TEXT, item_id TEXT, file_name TEXT, mime TEXT, uploaded_by TEXT, uploaded_at TEXT DEFAULT (datetime('now')), removed INTEGER NOT NULL DEFAULT 0)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS staff_docs_person ON staff_docs (person_email, removed)").run();
  for (const col of ["doc_date TEXT", "label TEXT"]) { try { await env.DB.prepare(`ALTER TABLE staff_docs ADD COLUMN ${col}`).run(); } catch (e) { /* column already there */ } }
  staffDocsMade = true;
}
const STAFF_DOC_COLS = "id, slot, file_name, uploaded_at, doc_date, label";
const staffDocOut = (r) => ({ id: r.id, slot: r.slot, name: r.file_name || "", uploadedAt: r.uploaded_at || "", docDate: r.doc_date || "", label: r.label || "" });
// Titles for the Additional Documents types the office can pick.
const STAFF_EXTRA_TYPES = { cv: "CV", details: "Employee Details", tfn: "Tax File Number Declaration", super: "Superannuation Form", template: "Contract Template" };
const isoDay = (s) => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(String(s || "").trim()); if (!m) return ""; const d = new Date(Date.UTC(+m[1], +m[2] - 1, +m[3])); return d.getUTCFullYear() === +m[1] && d.getUTCMonth() === +m[2] - 1 && d.getUTCDate() === +m[3] ? m[0] : ""; };
const fileExt = (n) => { const m = /(\.[A-Za-z0-9]{1,8})$/.exec(String(n || "").trim()); return m ? m[1].toLowerCase() : ""; };
// Puts a file under a folder item, creating the Employment Documents subfolder on the way. Returns the new item.
async function uploadUnderItem(env, driveBase, folderId, subSegs, fileName, bytes, contentType) {
  const base = `${driveBase}/items/${folderId}:/${pathOf([...subSegs, fileName])}:`;
  if (bytes.byteLength <= 4 * 1024 * 1024) {
    const res = await graphFetch(env, `${base}/content?@microsoft.graph.conflictBehavior=rename`, { method: "PUT", headers: { "Content-Type": contentType || "application/octet-stream" }, body: bytes });
    if (!res.ok) throw new Error("upload " + res.status);
    return res.json();
  }
  const sess = await graphFetch(env, `${base}/createUploadSession`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }) });
  if (!sess.ok) throw new Error("session " + sess.status);
  const { uploadUrl } = await sess.json();
  let last;
  for (let start = 0; start < bytes.byteLength; start += CHUNK) {
    const end = Math.min(start + CHUNK, bytes.byteLength);
    const res = await fetch(uploadUrl, { method: "PUT", headers: { "Content-Length": String(end - start), "Content-Range": `bytes ${start}-${end - 1}/${bytes.byteLength}` }, body: bytes.slice(start, end) });
    if (!res.ok && res.status !== 202) throw new Error("chunk " + res.status);
    last = res;
  }
  return last.json();
}
// ---- Policies, stored in D1 with the PDFs in SharePoint ----
let policyAckMade = false;
async function policyAckTable(env) {
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS policy_ack_log (id INTEGER PRIMARY KEY AUTOINCREMENT, person_key TEXT NOT NULL, name TEXT, email TEXT, policy_id TEXT NOT NULL, policy_title TEXT, version INTEGER, opened_at TEXT, read_at TEXT, acked_at TEXT NOT NULL)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS policy_ack_person ON policy_ack_log (person_key, policy_id)").run();
}
async function policyTable(env) {
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS policies (id TEXT PRIMARY KEY, title TEXT NOT NULL, item_id TEXT, file_name TEXT, needs_ack INTEGER NOT NULL DEFAULT 0, removed INTEGER NOT NULL DEFAULT 0, updated_by TEXT, updated_at TEXT NOT NULL DEFAULT (datetime('now')), version INTEGER NOT NULL DEFAULT 1, folder TEXT)").run();
  if (!policyAckMade) { try { await env.DB.prepare("UPDATE policies SET needs_ack = 1 WHERE needs_ack <> 1").run(); policyAckMade = true; } catch (e) { /* tried again next time */ } }
  for (const col of ["version INTEGER NOT NULL DEFAULT 1", "folder TEXT", "docx_item_id TEXT", "docx_name TEXT", "body TEXT", "effective TEXT", "updated_on TEXT", "updated_src TEXT", "kind TEXT NOT NULL DEFAULT 'policy'"]) { try { await env.DB.prepare(`ALTER TABLE policies ADD COLUMN ${col}`).run(); } catch (e) { /* column already there */ } }
}
// Plain text out of a .docx (a zip holding word/document.xml), for searching. No library: the zip index is read by hand and the one file inflated.
async function docxText(bytes) {
  const u = new Uint8Array(bytes);
  const dv = new DataView(u.buffer, u.byteOffset, u.byteLength);
  let eocd = -1;
  for (let i = u.length - 22; i >= Math.max(0, u.length - 65557); i--) { if (dv.getUint32(i, true) === 0x06054b50) { eocd = i; break; } }
  if (eocd < 0) return "";
  const count = dv.getUint16(eocd + 10, true);
  let p = dv.getUint32(eocd + 16, true);
  let xml = "", extra = "";
  for (let n = 0; n < count; n++) {
    if (dv.getUint32(p, true) !== 0x02014b50) break;
    const method = dv.getUint16(p + 10, true), csize = dv.getUint32(p + 20, true);
    const nlen = dv.getUint16(p + 28, true), elen = dv.getUint16(p + 30, true), clen = dv.getUint16(p + 32, true), lho = dv.getUint32(p + 42, true);
    const entry = new TextDecoder().decode(u.subarray(p + 46, p + 46 + nlen));
    // The body is word/document.xml. Page headers and footers (word/header1.xml, word/footer1.xml) are read too, since a "last updated" line is often kept in a footer.
    if (entry === "word/document.xml" || /^word\/(header|footer)\d*\.xml$/.test(entry)) {
      const start = lho + 30 + dv.getUint16(lho + 26, true) + dv.getUint16(lho + 28, true);
      const data = u.subarray(start, start + csize);
      const part = method === 0 ? new TextDecoder().decode(data) : await new Response(new Blob([data]).stream().pipeThrough(new DecompressionStream("deflate-raw"))).text();
      if (entry === "word/document.xml") xml = part; else extra += part + "</w:p>";
    }
    p += 46 + nlen + elen + clen;
  }
  xml = extra + xml; // headers and footers first, so the end of the text is the end of the document
  return xml.replace(/<w:tab\/>/g, " ").replace(/<\/w:p>/g, "\n").replace(/<w:br[^>]*\/>/g, "\n").replace(/<[^>]+>/g, "")
    .replace(/&#x([0-9a-fA-F]+);/g, (m, h) => String.fromCodePoint(parseInt(h, 16))).replace(/&#(\d+);/g, (m, d) => String.fromCodePoint(Number(d)))
    .replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&")
    .replace(/[ \t\u00a0]+/g, " ").replace(/\s*\n\s*/g, "\n").trim().slice(0, 300000);
}
// ---- Dates read out of a policy's own text ----
const MONTH_NUM = { jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6, jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12 };
const MONTH_RE = "(jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|june?|july?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)";
// Every full date in a piece of text, in order of appearance, as { iso, at, end }. Handles 12 August 2024, 12th Aug 2024, August 12, 2024,
// 12/08/2024, 12.08.24, 12-8-2024 and 2024-08-12. Day first (Australian) unless the second number can only be a day. A month and year on their own are skipped.
function datesIn(text) {
  const t = String(text || "");
  const out = [];
  const add = (y, m, d, at, end) => {
    if (y < 100) y += 2000;
    if (m >= 1 && m <= 12 && d >= 1 && d <= 31 && y >= 1990 && y <= 2100) {
      const dt = new Date(Date.UTC(y, m - 1, d));
      if (dt.getUTCMonth() === m - 1 && dt.getUTCDate() === d) out.push({ iso: `${y}-${String(m).padStart(2, "0")}-${String(d).padStart(2, "0")}`, at, end });
    }
  };
  let m;
  const a = new RegExp(`(?<![\\d])(\\d{1,2})(?:st|nd|rd|th)?\\s+(?:of\\s+)?${MONTH_RE}\\b\\.?,?\\s+(\\d{4})(?!\\d)`, "gi");
  while ((m = a.exec(t))) add(Number(m[3]), MONTH_NUM[m[2].slice(0, 3).toLowerCase()], Number(m[1]), m.index, m.index + m[0].length);
  const b = new RegExp(`\\b${MONTH_RE}\\b\\.?\\s+(\\d{1,2})(?:st|nd|rd|th)?,?\\s+(\\d{4})(?!\\d)`, "gi");
  while ((m = b.exec(t))) add(Number(m[3]), MONTH_NUM[m[1].slice(0, 3).toLowerCase()], Number(m[2]), m.index, m.index + m[0].length);
  const c = /(?<![\d\/.\-])(\d{4})-(\d{1,2})-(\d{1,2})(?![\d\-])/g;
  while ((m = c.exec(t))) add(Number(m[1]), Number(m[2]), Number(m[3]), m.index, m.index + m[0].length);
  const d = /(?<![\d\/.\-])(\d{1,2})([\/.\-])(\d{1,2})\2(\d{4}|\d{2})(?![\d\/]|\.\d)/g;
  while ((m = d.exec(t))) {
    if (m[4].length === 2 && m[2] !== "/") continue; // 1.2.03 is a version number, not a date
    let day = Number(m[1]), mon = Number(m[3]);
    if (mon > 12 && day <= 12) { const x = day; day = mon; mon = x; }
    add(Number(m[4]), mon, day, m.index, m.index + m[0].length);
  }
  out.sort((x, y) => x.at - y.at);
  // The same text can match two shapes (12/08/2024 inside another); keep the first of any overlap.
  return out.filter((x, i) => i === 0 || x.at >= out[i - 1].end);
}
// The "effective from" date a policy states about itself, as yyyy-mm-dd. Returns "" when the text gives none.
function findEffective(body) {
  const text = String(body || "");
  const re = /effective/gi;
  let hit;
  while ((hit = re.exec(text))) {
    const win = text.slice(hit.index, hit.index + 120);
    const d = datesIn(win)[0];
    if (d) return d.iso;
  }
  return "";
}
// The date a policy says it was last updated, as { iso, how } or null. Looks first for a labelled date (Last updated, Date last reviewed, Revision date, Version date and so on)
// and takes the most recent one. Failing that, takes the most recent date in the last stretch of the document, where revision tables and sign off lines sit.
// Dates in the future, dates that follow "next review" or "effective", and dates before 2000 are never taken.
// Strong labels are clear enough to look up to 80 characters on for the date. Weak labels ("Version date", "Issue date") are also column headings in revision tables,
// so for those the date has to follow straight on, with nothing but punctuation between.
const UPDATED_STRONG = new RegExp("(?:last\\s+(?:updated|reviewed|revised|amended|modified|edited|update|review|revision|amendment)(?:\\s+(?:date|on))?|date\\s+(?:last\\s+)?(?:updated|reviewed|revised|amended|modified)|date\\s+of\\s+(?:the\\s+)?(?:last\\s+)?(?:update|review|revision|amendment|issue)|(?:updated|revised|amended|reviewed|modified|issued)\\s+on)", "gi");
const UPDATED_WEAK = new RegExp("(?:(?:revision|version|amendment|update|issue|issued|reviewed|revised|updated)\\s+date|date\\s+issued)", "gi");
function findLastUpdated(body, nowMs) {
  const text = String(body || "");
  const limit = new Date((nowMs || Date.now()) + 36 * 3600 * 1000).toISOString().slice(0, 10);
  const fine = (iso) => policyDateOk(iso) && iso <= limit;
  let best = "";
  for (const [re, tight] of [[UPDATED_STRONG, false], [UPDATED_WEAK, true]]) {
    re.lastIndex = 0;
    let hit;
    while ((hit = re.exec(text))) {
      if (/(next|due|upcoming)\s*$/i.test(text.slice(Math.max(0, hit.index - 12), hit.index))) continue;
      const from = hit.index + hit[0].length;
      const win = text.slice(from, from + 80);
      const d = datesIn(win).find((x) => fine(x.iso));
      if (!d) continue;
      if (tight && !/^[\s:|\u2013\u2014\-.,]*(?:(?:is|was)\s*)?[\s:|\u2013\u2014\-.,]*$/i.test(win.slice(0, d.at))) continue;
      if (d.iso > best) best = d.iso;
    }
  }
  if (best) return { iso: best, how: "label" };
  const tailAt = Math.max(0, text.length - 1500);
  let tailBest = "";
  for (const d of datesIn(text.slice(tailAt))) {
    if (!fine(d.iso)) continue;
    const before = text.slice(Math.max(0, tailAt + d.at - 50), tailAt + d.at);
    if (/(next\s+review|review\s+due|due\s+for\s+review|due\s+date|effective|expires?|expiry|printed|signed)[^\n]{0,40}$/i.test(before)) continue;
    if (d.iso > tailBest) tailBest = d.iso;
  }
  return tailBest ? { iso: tailBest, how: "end" } : null;
}
// A policy name out of a file name: no folder, no extension, no trailing version tag or "(1)", underscores as spaces.
function titleFromName(name) {
  let t = String(name || "").split(/[\\/]/).pop().replace(/\.(pdf|docx|doc)$/i, "").replace(/_/g, " ");
  t = t.replace(/\s*\(\d+\)\s*$/, "").replace(/[\s\-]+v\s?\d+(?:\.\d+)*\s*$/i, "").replace(/[\s\-]+(?:version|ver)\s?\d+(?:\.\d+)*\s*$/i, "");
  return t.replace(/\s+/g, " ").trim().slice(0, 120);
}
const DOCX_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const MAX_CONVERT = 2 * 1024 * 1024;
const isWordFile = (name, type) => /\.docx$/i.test(String(name || "")) || /wordprocessingml/i.test(String(type || ""));
const policyFail = (status, say, e) => Object.assign(new Error(say + (e ? ` (${String(e.message || e).slice(0, 120)})` : "")), { status, say });
// Asks Microsoft Graph to turn a stored Word file into a PDF. Needs Sites.ReadWrite.All or Files.ReadWrite.All on the Graph app.
async function graphToPdf(env, itemId) {
  const call = () => graphFetch(env, `/sites/${siteId(env)}/drive/items/${itemId}/content?format=pdf`, { redirect: "manual" });
  let res = await call();
  if (res.status === 401 || res.status === 403) { graphToken = { at: 0, value: "" }; res = await call(); }
  const loc = res.headers.get("Location");
  const file = loc ? await fetch(loc) : res;
  if (!file.ok) throw new Error(`convert ${res.status} ${file.status}`);
  const bytes = await file.arrayBuffer();
  const head = new TextDecoder().decode(new Uint8Array(bytes.slice(0, 5)));
  if (!head.startsWith("%PDF")) throw new Error("convert gave something that is not a PDF");
  return bytes;
}
// Where policy files live. Policies belong under Workplace Culture. Until the folder has been moved there, the old top level Policies folder keeps working,
// so the move can be done in SharePoint at any time without breaking uploads. The result is kept for one minute.
const POLICY_HOMES = [["Workplace Culture", "Policies"], ["Reservoir", "Workplace Culture", "Policies"], ["Policies"]];
let policyHomeKept = null;
async function policyHome(env) {
  if (policyHomeKept && Date.now() - policyHomeKept.at < 60000) return policyHomeKept.segs;
  let found = null;
  for (const segs of POLICY_HOMES) {
    try {
      const r = await graphFetch(env, `/sites/${siteId(env)}/drive/root:/${pathOf(segs)}?select=id,folder`);
      if (!r.ok) continue;
      const it = await r.json();
      if (it.folder) { found = segs; break; }
    } catch (e) { /* try the next place */ }
  }
  if (!found) return POLICY_HOMES[0];
  policyHomeKept = { at: Date.now(), segs: found };
  return found;
}
// ---- Document libraries ----
// Policies and the Handbook are one thing kept in two places: the same table, routes, versions, Archive folder, text reading and date reading.
// The only difference is the SharePoint folder: Workplace Culture / Policies and Workplace Culture / Handbook.
const docKind = (v) => (v === "handbook" ? "handbook" : "policy");
async function docHome(env, kind) {
  const home = await policyHome(env);
  return kind === "handbook" ? [...home.slice(0, -1), "Handbook"] : home;
}
// A file only counts as saved once SharePoint is asked for it again: the folder must exist, the file must be in that folder, and a PDF must be the size that was sent.
async function confirmStored(env, segs, item, expectedSize) {
  const lookup = async (path) => { try { const r = await graphFetch(env, path); return r.ok ? await r.json() : null; } catch (e) { return null; } };
  const dir = await lookup(`/sites/${siteId(env)}/drive/root:/${pathOf(segs)}?select=id,folder`);
  const file = item && item.id ? await lookup(`/sites/${siteId(env)}/drive/items/${item.id}?select=id,name,size,file,parentReference`) : null;
  const inFolder = !!(dir && dir.folder && file && file.file && file.parentReference && file.parentReference.id === dir.id);
  const sized = !!file && file.size > 0 && (expectedSize === undefined || file.size === expectedSize);
  if (!inFolder || !sized) throw policyFail(502, `The file could not be found in SharePoint after saving (${segs.join(" / ")}). Nothing was added.`);
  return file;
}
// Publishes the next version of a document. A Word file is saved as "Title vn.docx", converted by Graph, and the PDF saved as "Title vn.pdf". The v is always lowercase.
// Every file saved is confirmed in SharePoint before anything else happens. Only once the new files are safely stored does the previous pair move into the Archive folder,
// so a failed conversion or a missing file leaves the current version untouched.
async function policyPublish(env, prev, folder, next, bytes, ctype, origName, kind) {
  const home = await docHome(env, kind);
  const dest = [...home, folder];
  const word = isWordFile(origName, ctype);
  let docxItem = null, pdfItem, pdfSize;
  if (word) {
    if (bytes.byteLength > MAX_CONVERT) throw policyFail(413, "That Word file is over 2 MB, the most Microsoft will convert. Shrink any pictures and try again.");
    docxItem = await uploadToSharePoint(env, dest, `${folder} v${next}.docx`, bytes, DOCX_TYPE, false);
    try {
      const pdfBytes = await graphToPdf(env, docxItem.id);
      pdfSize = pdfBytes.byteLength;
      pdfItem = await uploadToSharePoint(env, dest, `${folder} v${next}.pdf`, pdfBytes, "application/pdf", false);
    } catch (e) {
      try { await graphFetch(env, `/sites/${siteId(env)}/drive/items/${docxItem.id}`, { method: "DELETE" }); } catch (x) { /* leave it */ }
      throw policyFail(502, "Microsoft could not convert that Word file to a PDF. The current version has not been changed.", e);
    }
  } else {
    pdfSize = bytes.byteLength;
    pdfItem = await uploadToSharePoint(env, dest, `${folder} v${next}.pdf`, bytes, "application/pdf", false);
  }
  try {
    if (docxItem) await confirmStored(env, dest, docxItem);
    await confirmStored(env, dest, pdfItem, pdfSize);
  } catch (e) {
    for (const it of [docxItem, pdfItem]) { if (it && it.id) { try { await graphFetch(env, `/sites/${siteId(env)}/drive/items/${it.id}`, { method: "DELETE" }); } catch (x) { /* leave it */ } } }
    throw e;
  }
  // The previous PDF goes into Archive. The previous Word file is deleted, since the new Word file is the master.
  let archived = true;
  if (prev && prev.item_id) {
    try {
      const old = await driveBytes(env, prev.item_id);
      await uploadToSharePoint(env, [...dest, "Archive"], prev.file_name || `${folder} v${next - 1}.pdf`, old, "application/pdf", false);
      await graphFetch(env, `/sites/${siteId(env)}/drive/items/${prev.item_id}`, { method: "DELETE" });
    } catch (e) { archived = false; console.log("policy archive error", String(e && e.message || e)); }
  }
  if (prev && prev.docx_item_id) {
    try { await graphFetch(env, `/sites/${siteId(env)}/drive/items/${prev.docx_item_id}`, { method: "DELETE" }); } catch (e) { console.log("policy old word delete error", String(e && e.message || e)); }
  }
  let body = "";
  if (word) { try { body = await docxText(bytes); } catch (e) { console.log("policy text error", String(e && e.message || e)); } }
  return { body, effective: body ? findEffective(body) : "", itemId: pdfItem.id, fileName: pdfItem.name || `${folder} v${next}.pdf`, docxId: docxItem ? docxItem.id : "", docxName: docxItem ? docxItem.name || `${folder} v${next}.docx` : "", archived };
}
// A title that was removed earlier can leave its files behind in SharePoint. Before the title starts again at v1 they are moved into its Archive folder, so the new v1 never clashes with them.
async function archiveLeftovers(env, dest) {
  let list;
  try {
    const r = await graphFetch(env, `/sites/${siteId(env)}/drive/root:/${pathOf(dest)}:/children?$select=id,name,file&$top=200`);
    if (!r.ok) return 0;
    list = (await r.json()).value || [];
  } catch (e) { return 0; }
  let moved = 0;
  for (const it of list) {
    if (!it.file) continue;
    // Only files this app made ("Title v1.pdf", "Title v1.docx", and a renamed copy) are ever moved. Anything else in the folder is left alone.
    if (!/ v\d+(?: \d+)?\.(pdf|docx)$/i.test(it.name)) continue;
    try {
      const old = await driveBytes(env, it.id);
      await uploadToSharePoint(env, [...dest, "Archive"], it.name, old, "application/octet-stream", false);
      await graphFetch(env, `/sites/${siteId(env)}/drive/items/${it.id}`, { method: "DELETE" });
      moved++;
    } catch (e) {
      throw policyFail(502, `The earlier files for this title could not be moved to its Archive (${dest.join(" / ")}). Nothing was added.`, e);
    }
  }
  return moved;
}
const policyOut = (r) => ({ id: r.id, title: r.title, version: r.version || 1, hasWord: !!r.docx_item_id, effective: r.effective || "", fileName: r.file_name || "", hasFile: !!r.item_id, needsAck: true, kind: r.kind === "handbook" ? "handbook" : "policy", updatedOn: r.updated_on || "", updatedSrc: r.updated_on ? (policyDateManual(r) ? "manual" : r.updated_src) : "", updated: r.updated_on ? r.updated_on : String(r.updated_at || "").replace(" ", "T") + "Z" });
// The "last updated" date shown for a policy is the date the director sets from the document itself. Until one is set, the upload date is shown. A plain yyyy-mm-dd, or "" to clear it.
function policyDateOk(v) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(v || ""))) return false;
  const d = new Date(v + "T00:00:00Z");
  return !isNaN(d) && d.toISOString().slice(0, 10) === v && Number(v.slice(0, 4)) >= 2000;
}

// A date set by hand (or saved before dates were read automatically) is never overwritten by reading the document.
// Nobody can set a last updated date by hand any more: it only ever comes from reading the document. A date left from before that is replaced by the next read of the document.
const policyDateManual = (r) => false;
const policySrc = (f) => (f ? (f.how === "label" ? "doc" : "doc-end") : null);
// Stores the text of a policy and reads its dates out of it. The effective date always follows the text. The last updated date follows the text unless it was set by hand.
async function policyApplyText(env, r, text) {
  const effective = findEffective(text);
  const last = findLastUpdated(text);
  const manual = policyDateManual(r);
  if (manual) await env.DB.prepare("UPDATE policies SET body = ?, effective = ? WHERE id = ?").bind(text, effective, r.id).run();
  else await env.DB.prepare("UPDATE policies SET body = ?, effective = ?, updated_on = ?, updated_src = ? WHERE id = ?").bind(text, effective, last ? last.iso : null, policySrc(last), r.id).run();
  return { effective, last, manual };
}
// PDF text read by the app (the Worker cannot read PDF text itself) comes in here. A policy that has a Word file keeps the text from the Word file.
async function policyTextSave(env, row, request) {
  let j;
  try { j = await request.json(); } catch { throw policyFail(400, "bad json"); }
  const text = String(j && j.text || "").replace(/[ \t\u00a0]+/g, " ").replace(/\s*\n\s*/g, "\n").trim().slice(0, 300000);
  if (!text) throw policyFail(400, "No text came with that request");
  const cur = await env.DB.prepare("SELECT body FROM policies WHERE id = ?").bind(row.id).first();
  let skipped = false;
  if (row.docx_item_id && cur && cur.body) skipped = true;
  else await policyApplyText(env, row, text);
  const out = await env.DB.prepare("SELECT id, title, item_id, file_name, needs_ack, updated_at, updated_on, updated_src, version, folder, docx_item_id, docx_name, effective, kind FROM policies WHERE id = ?").bind(row.id).first();
  return { policy: policyOut(out), skipped };
}

const money2 = (v) => { const n = Number(v); return Number.isFinite(n) ? Math.round(n * 100) / 100 : 0; };
const txt = (v, n) => String(v === undefined || v === null ? "" : v).replace(/\s+/g, " ").trim().slice(0, n);
// Australian financial year: starts 1 July. Returns { label: "2026", start: "2026-07-01" } for the year a date falls in.
function financialYear(iso) {
  const y = Number(iso.slice(0, 4));
  const m = Number(iso.slice(5, 7));
  const start = m >= 7 ? y : y - 1;
  return { label: String(start), start: `${start}-07-01` };
}


// ---- PPE: the bag acknowledgement, replenishment orders and the four stage trail (requested, approved, provided, received) ----
// Everything is kept in DFP's own database and nothing here touches Xero. Items are fixed on the Worker so a request can only name a real item.
// "stock" items are supplied by DFP from the shed. "buy" items are bought by the person after approval, with the receipt uploaded against the order to close it.
const PPE_BOOT_SIZES = ["AU 9", "AU 10", "AU 11", "AU 12", "AU 13"];
const PPE_ITEMS = [
  { id: "earmuffs", label: "Ear Muffs", kind: "stock" },
  { id: "earplugs", label: "Ear Plugs", kind: "stock" },
  { id: "masks", label: "Face Masks", kind: "stock" },
  { id: "firstaid", label: "First Aid Kit", kind: "stock", expiry: true },
  { id: "cutgloves", label: "Gloves: Cut Resistant", kind: "stock", sizes: ["S", "M", "L", "XL"] },
  { id: "silicone", label: "Gloves: Silicone", kind: "stock", sizes: ["S", "M", "L", "XL"] },
  { id: "hardhat", label: "Hard Hat", kind: "stock" },
  { id: "respirator", label: "Respirator: P2", kind: "stock" },
  { id: "glassesclear", label: "Safety Glasses: Clear", kind: "stock" },
  { id: "glassesuv", label: "Safety Glasses: UV", kind: "stock" },
  { id: "steelcap", label: "Steel Cap Boots", kind: "stock", sizes: PPE_BOOT_SIZES },
  { id: "sunscreen", label: "Sunscreen", kind: "stock" }
];
const ppeItem = (id) => PPE_ITEMS.find((x) => x.id === id);
let ppeReady = false;
async function ppeTables(env) {
  if (ppeReady) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS ppe_state (person_key TEXT PRIMARY KEY, name TEXT, email TEXT, have TEXT, first_aid_expiry TEXT, acknowledged_at TEXT, updated_at TEXT DEFAULT (datetime('now')))").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS ppe_orders (id TEXT PRIMARY KEY, person_key TEXT, name TEXT, email TEXT, items TEXT, source TEXT, status TEXT NOT NULL DEFAULT 'requested', po_number TEXT, receipt_item TEXT, created_at TEXT DEFAULT (datetime('now')), approved_by TEXT, approved_at TEXT, provided_by TEXT, provided_at TEXT, received_at TEXT, declined_by TEXT, declined_at TEXT)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS ppe_orders_person ON ppe_orders (person_key)").run();
  ppeReady = true;
}
const ppePerson = (user) => ({ key: profileKey(user.name) || (user.email ? "email:" + String(user.email).toLowerCase() : ""), name: String(user.name || user.email || "").slice(0, 120), email: String(user.email || "").toLowerCase() });
const ppeJson = (t, d) => { try { const v = JSON.parse(t || ""); return v === null || v === undefined ? d : v; } catch (e) { return d; } };
const ppeOrderOut = (r) => {
  const items = ppeJson(r.items, []);
  return { id: r.id, kind: "ppe", type: "PPE order", by: String(r.name || "").split(" ")[0], byName: r.name || "", items, source: r.source || "order", status: r.status, poNumber: r.po_number || "", hasReceipt: !!r.receipt_item, date: String(r.created_at || "").slice(0, 10), approvedDate: r.approved_at ? String(r.approved_at).slice(0, 10) : "", providedDate: r.provided_at ? String(r.provided_at).slice(0, 10) : "", receivedDate: r.received_at ? String(r.received_at).slice(0, 10) : "", text: items.map((x) => x.label + (x.size ? " " + x.size : "")).join(", "), server: true };
};
// Where the person's bag stands. none: never acknowledged. partial: something is outstanding. full: every stock item is in the bag and the bought item has its receipt.
function ppeBagStatus(state, orders) {
  if (!state || !state.acknowledged_at) return "none";
  const have = ppeJson(state.have, {});
  const bought = (id) => orders.some((o) => o.status === "received" && ppeJson(o.items, []).some((x) => x.id === id));
  const missing = PPE_ITEMS.filter((x) => x.kind === "stock" ? !have[x.id] : !bought(x.id));
  return missing.length ? "partial" : "full";
}
async function ppeMe(env, user) {
  await ppeTables(env);
  const who = ppePerson(user);
  const state = await env.DB.prepare("SELECT * FROM ppe_state WHERE person_key = ?").bind(who.key).first();
  const rows = ((await env.DB.prepare("SELECT * FROM ppe_orders WHERE person_key = ? ORDER BY created_at DESC, rowid DESC LIMIT 100").bind(who.key).all()).results) || [];
  const status = ppeBagStatus(state, rows);
  const open = (id) => rows.some((o) => ["requested", "approved", "provided"].includes(o.status) && ppeJson(o.items, []).some((x) => x.id === id));
  return { status, acknowledged: !!(state && state.acknowledged_at), have: ppeJson(state && state.have, {}), firstAidExpiry: (state && state.first_aid_expiry) || "", outstanding: PPE_ITEMS.filter((x) => open(x.id)).map((x) => x.id), orders: rows.map(ppeOrderOut) };
}
function ppeCleanItems(list) {
  if (!Array.isArray(list) || !list.length || list.length > 12) return null;
  const seen = new Set(), out = [];
  for (const x of list) {
    const it = ppeItem(x && x.id);
    if (!it || seen.has(it.id)) return null;
    seen.add(it.id);
    const size = String((x && x.size) || "").toUpperCase();
    if (it.sizes) { if (size && !it.sizes.includes(size)) return null; }
    else if (size) return null;
    out.push({ id: it.id, label: it.label, kind: it.kind, ...(it.sizes ? { size } : {}) });
  }
  return out;
}
// Makes one order for the stock items and one for the bought items, so each can move through its own stages.
async function ppeCreateOrders(env, who, items, source) {
  const made = [];
  for (const group of [items.filter((x) => x.kind === "stock"), items.filter((x) => x.kind === "buy")]) {
    if (!group.length) continue;
    const id = crypto.randomUUID();
    await env.DB.prepare("INSERT INTO ppe_orders (id, person_key, name, email, items, source, status) VALUES (?, ?, ?, ?, ?, ?, 'requested')").bind(id, who.key, who.name, who.email, JSON.stringify(group), source).run();
    made.push(ppeOrderOut(await env.DB.prepare("SELECT * FROM ppe_orders WHERE id = ?").bind(id).first()));
  }
  return made;
}

// ---- Stored pay slips: pay slips for current employees, kept in Reservoir's own database ----
// The pull below is the only thing that reads pay slips from Xero. It is run by a director from the Development screen, and is safe to run again: a pay run already stored is skipped, so the same action is the weekly top up.
// People and screens read the stored copy, so nobody else ever causes a Xero call for a pay slip.
// Only current employees are stored (people with a row in xero_payroll). Anyone else on a pay run is skipped. No tax file number and no bank number is stored: the bank block is blanked before saving and filled in again when the person views their own slip, from dfp_xero.
// The first payment date held in Reservoir: the first pay day of the 2025 to 2026 financial year. Anyone needing an earlier pay slip asks the office.
const PAYSLIP_FROM = "2025-07-03";
// The first financial year whose every pay day is held, so its year to date can be added up in full.
const PAYSLIP_FULL_FY = "2025-07-01";
// Bumped when the rules for building a stored pay slip change. A run stored under older rules is not counted as done, so the next pull rebuilds it from Xero.
const PAYSLIP_RULES = 2;
let payslipReady = false;
async function payslipTables(env) {
  if (payslipReady) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS payslip_store (run_id TEXT NOT NULL, employee_id TEXT NOT NULL, email TEXT, name TEXT, period_start TEXT, period_end TEXT, payment_date TEXT NOT NULL, gross REAL NOT NULL DEFAULT 0, tax REAL NOT NULL DEFAULT 0, super_amt REAL NOT NULL DEFAULT 0, net REAL NOT NULL DEFAULT 0, slip TEXT NOT NULL, built_at TEXT DEFAULT (datetime('now')), PRIMARY KEY (run_id, employee_id))").run();
  // missing = 1 when Xero left a main total out of this slip. rules says which version of the pay slip rules built the row, so older rows are rebuilt by the next pull.
  for (const col of ["missing INTEGER NOT NULL DEFAULT 0", "rules INTEGER NOT NULL DEFAULT 0"]) { try { await env.DB.prepare(`ALTER TABLE payslip_store ADD COLUMN ${col}`).run(); } catch (e) { /* already there */ } }
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS payslip_store_email ON payslip_store (email, payment_date)").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS payslip_runs (run_id TEXT PRIMARY KEY, payment_date TEXT, period_start TEXT, period_end TEXT, people INTEGER NOT NULL DEFAULT 0, skipped INTEGER NOT NULL DEFAULT 0, built_at TEXT DEFAULT (datetime('now')))").run();
  try { await env.DB.prepare("ALTER TABLE payslip_runs ADD COLUMN rules INTEGER NOT NULL DEFAULT 0").run(); } catch (e) { /* already there */ }
  // filed = 1 once the director has paid that pay run and tapped Pay Run Filed. Reservoir only, never a Xero figure. When the column is first added, every run already paid (payment date before today, Melbourne time) is marked filed so only upcoming runs show as Next Pay.
  let filedAdded = false;
  try { await env.DB.prepare("ALTER TABLE payslip_runs ADD COLUMN filed INTEGER NOT NULL DEFAULT 0").run(); filedAdded = true; } catch (e) { /* already there */ }
  try { await env.DB.prepare("ALTER TABLE payslip_runs ADD COLUMN filed_at TEXT").run(); } catch (e) { /* already there */ }
  try { await env.DB.prepare("ALTER TABLE payslip_runs ADD COLUMN filed_by TEXT").run(); } catch (e) { /* already there */ }
  if (filedAdded) await env.DB.prepare("UPDATE payslip_runs SET filed = 1, filed_at = datetime('now'), filed_by = 'Existing run' WHERE payment_date < date('now', '+11 hours')").run();
  payslipReady = true;
}
// Employee records kept between batches so a person's record is not fetched again for every batch.
const payslipRecords = new Map();
async function payslipRecordFor(env, employeeId) {
  const kept = payslipRecords.get(employeeId);
  if (kept && Date.now() - kept.at < 15 * 60000) return kept.record;
  const r = await xeroFetch(env, `payroll.xro/1.0/Employees/${employeeId}`);
  if (!r.ok) throw new Error(`Xero refused the employee record (${r.status})`);
  const record = ((await r.json()).Employees || [])[0];
  if (!record) throw new Error("Xero returned no employee record");
  payslipRecords.set(employeeId, { at: Date.now(), record });
  return record;
}
// Builds and stores the next few pay runs that are not stored yet, oldest first. Returns progress, never any pay figure.
async function payslipPull(env, opts) {
  const maxRuns = Math.max(1, Math.min(8, Number(opts && opts.runs) || 3));
  await dfpTables(env);
  await payslipTables(env);
  const people = ((await env.DB.prepare("SELECT employee_id, email, name FROM xero_payroll WHERE employee_id IS NOT NULL AND email IS NOT NULL AND email <> '' AND COALESCE(status, 'ACTIVE') = 'ACTIVE'").all()).results) || [];
  if (!people.length) return { fail: { status: 409, error: "No current employees are in the database yet. Run Populate DFP Internal first." } };
  const byId = new Map(people.map((x) => [x.employee_id, x]));
  // Always read the pay run list fresh, so a run posted in Xero a moment ago is seen. With replace, runs posted here but not yet filed are rebuilt from Xero.
  xeroRunsKept = null;
  if (opts && opts.replace) await env.DB.prepare("UPDATE payslip_runs SET rules = 0 WHERE filed = 0").run();
  const posted = await xeroPostedRuns(env);
  if (posted.fail) return { fail: { status: 502, error: "Xero refused the pay run list", detail: posted.fail.detail || "" } };
  const runs = posted.runs.filter((x) => xeroDate(x.PaymentDate) >= PAYSLIP_FROM).sort((a, b) => xeroDate(a.PaymentDate).localeCompare(xeroDate(b.PaymentDate)));
  const doneRows = ((await env.DB.prepare("SELECT run_id FROM payslip_runs WHERE rules >= ?").bind(PAYSLIP_RULES).all()).results) || [];
  const done = new Set(doneRows.map((x) => x.run_id));
  const pending = runs.filter((x) => !done.has(x.PayRunID));
  const out = { from: PAYSLIP_FROM, runsSince: runs.length, runsDoneBefore: runs.length - pending.length, runsBuilt: 0, slipsStored: 0, skippedNotCurrent: 0, failed: [], flagged: [], remaining: pending.length, people: people.length };
  if (!pending.length) return out;
  const items = await xeroItems(env);
  const kept = {};
  const supFor = new Map();
  for (const run of pending.slice(0, maxRuns)) {
    const rd = await xeroRunDetail(env, run.PayRunID);
    if (rd.fail) { out.failed.push(`pay run ${xeroDate(run.PaymentDate)}: ${(rd.fail.body && rd.fail.body.error) || "Xero refused"}`); break; }
    const detail = rd.run;
    if (!detail || (detail.PayRunStatus || "") !== "POSTED") { out.failed.push(`pay run ${xeroDate(run.PaymentDate)} is not posted`); continue; }
    let stored = 0, skipped = 0, bad = 0;
    const storedIds = [];
    for (const summary of detail.Payslips || []) {
      const person = byId.get(summary.EmployeeID);
      if (!person) { skipped++; continue; }
      if (!summary.PayslipID) { bad++; out.failed.push(`${person.name}: no payslip on ${xeroDate(run.PaymentDate)}`); continue; }
      try {
        const record = await payslipRecordFor(env, person.employee_id);
        if (!supFor.has(person.employee_id)) supFor.set(person.employee_id, await xeroSuperFor(env, record, kept));
        const sup = supFor.get(person.employee_id);
        const sr = await xeroFetch(env, `payroll.xro/1.0/Payslip/${summary.PayslipID}`);
        if (!sr.ok) throw new Error(`Xero refused the payslip (${sr.status})`);
        const sj = await sr.json();
        const slip = Array.isArray(sj.Payslip) ? sj.Payslip[0] : sj.Payslip || (sj.Payslips || [])[0];
        if (!slip || (slip.EmployeeID && slip.EmployeeID !== person.employee_id)) throw new Error("Payslip not found");
        const ms = Array.isArray(record.SuperMemberships) ? record.SuperMemberships : [];
        const fundOf = (m) => { const i = ms.findIndex((x) => x.SuperMembershipID === m.SuperMembershipID); return i >= 0 ? (sup.memberships || [])[i] : null; };
        const tot = slipTotals(summary, slip);
        const shaped = xeroSlipShape(slip, detail, record, items, ms, fundOf, tot);
        // Nothing about the bank goes into the stored copy.
        shaped.bank = { accountName: "", bsb: "", masked: "" };
        shaped.source = "reservoir";
        const gap = shaped.missing && shaped.missing.length ? 1 : 0;
        if (gap) out.flagged.push(`${person.name}, ${xeroDate(run.PaymentDate)}: Xero did not supply ${shaped.missing.join(", ")}`);
        await env.DB.prepare("INSERT OR REPLACE INTO payslip_store (run_id, employee_id, email, name, period_start, period_end, payment_date, gross, tax, super_amt, net, slip, missing, rules) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").bind(detail.PayRunID, person.employee_id, String(person.email).toLowerCase(), person.name || "", shaped.period.start, shaped.period.end, shaped.paymentDate, tot.gross || 0, tot.tax || 0, tot.super || 0, tot.net || 0, JSON.stringify(shaped), gap, PAYSLIP_RULES).run();
        stored++;
        storedIds.push(person.employee_id);
      } catch (e) {
        bad++;
        out.failed.push(`${person.name}, ${xeroDate(run.PaymentDate)}: ${String(e && e.message || e).slice(0, 120)}`);
      }
    }
    out.slipsStored += stored;
    out.skippedNotCurrent += skipped;
    // A run is only marked done when every current person on it was stored, so a failed one is tried again next time.
    if (!bad) {
      // A rebuilt run drops any slip for someone who is no longer on it in Xero.
      if (storedIds.length) await env.DB.prepare(`DELETE FROM payslip_store WHERE run_id = ? AND employee_id NOT IN (${storedIds.map(() => "?").join(",")})`).bind(detail.PayRunID, ...storedIds).run();
      await env.DB.prepare("INSERT INTO payslip_runs (run_id, payment_date, period_start, period_end, people, skipped, rules) VALUES (?, ?, ?, ?, ?, ?, ?) ON CONFLICT(run_id) DO UPDATE SET payment_date = excluded.payment_date, period_start = excluded.period_start, period_end = excluded.period_end, people = excluded.people, skipped = excluded.skipped, rules = excluded.rules, built_at = datetime('now')").bind(detail.PayRunID, xeroDate(detail.PaymentDate || run.PaymentDate), xeroDate(detail.PayRunPeriodStartDate || run.PayRunPeriodStartDate), xeroDate(detail.PayRunPeriodEndDate || run.PayRunPeriodEndDate), stored, skipped, PAYSLIP_RULES).run();
      out.runsBuilt++;
      out.remaining--;
    }
  }
  return out;
}
// One person's stored pay slip list, from the database alone.
async function payslipListFor(env, email) {
  if (!email) return { found: false, reason: "No email on your sign in" };
  await dfpTables(env);
  await payslipTables(env);
  const rows = ((await env.DB.prepare("SELECT s.run_id, s.period_start, s.period_end, s.payment_date, json_extract(s.slip, '$.net') AS net, json_extract(s.slip, '$.lines') AS lines, s.gross AS g, s.tax AS t, s.super_amt AS sp, s.net AS nn, COALESCE(r.filed, 1) AS filed FROM payslip_store s LEFT JOIN payslip_runs r ON r.run_id = s.run_id WHERE s.email = ? ORDER BY s.payment_date DESC").bind(email).all()).results) || [];
  // Hours per financial year (1 July to 30 June, by payment date), added up from the stored slip lines. Leave and public holiday lines are counted apart from hours worked.
  const fyMap = new Map();
  for (const r of rows) {
    let lines = [];
    try { lines = JSON.parse(r.lines || "[]") || []; } catch (e) { lines = []; }
    const fy = financialYear(String(r.payment_date));
    const row = fyMap.get(fy.label) || { label: `${fy.label}/${Number(fy.label) + 1}`, worked: 0, leave: 0, runs: 0 };
    for (const l of lines) {
      const h = Number(l.hours) || 0;
      if (!h) continue;
      if (/leave|holiday|sick|carer|personal|bereave|compassion|jury|long service/i.test(String(l.label || ""))) row.leave += h; else row.worked += h;
    }
    row.runs++;
    fyMap.set(fy.label, row);
  }
  // Financial year stats from 2025/2026 on: pay totals, hours worked, and days off. Pay and hours come from the filed payslips. Days off come from the same approved leave records the Leave screen lists (split pieces left out, as there), counted day by day inside the year, so the two screens always agree. Days after today are Booked Ahead, not taken.
  let fyStats = [];
  try {
  await leaveTable(env);
  const dayLen = await leaveDayHours(env, email);
  const todayIso = new Date().toISOString().slice(0, 10);
  const leaveRows = ((await env.DB.prepare("SELECT * FROM xero_leave WHERE email = ? AND status = 'approved' AND split_of IS NULL").bind(email).all()).results) || [];
  const stat = new Map();
  for (const r of rows) {
    if (r.filed === 0) continue;
    const fy = financialYear(String(r.payment_date));
    if (Number(fy.label) < 2025) continue;
    const o = stat.get(fy.label) || { fy: fy.label, label: `${fy.label}/${Number(fy.label) + 1}`, closed: `${Number(fy.label) + 1}-06-30` < todayIso, runs: 0, firstPaid: String(r.payment_date), lastPaid: String(r.payment_date), gross: 0, tax: 0, super: 0, net: 0, worked: 0 };
    let lines = [];
    try { lines = JSON.parse(r.lines || "[]") || []; } catch (e) { lines = []; }
    for (const l of lines) {
      const h = Number(l.hours) || 0;
      if (!h) continue;
      if (!/leave|holiday|sick|carer|personal|bereave|compassion|jury|long service/i.test(String(l.label || ""))) o.worked += h;
    }
    o.runs++;
    o.gross += Number(r.g) || 0; o.tax += Number(r.t) || 0; o.super += Number(r.sp) || 0; o.net += Number(r.nn) || 0;
    if (String(r.payment_date) < o.firstPaid) o.firstPaid = String(r.payment_date);
    if (String(r.payment_date) > o.lastPaid) o.lastPaid = String(r.payment_date);
    stat.set(fy.label, o);
  }
  const r1 = (n) => Math.round(n * 10) / 10;
  fyStats = [...stat.values()].map((o) => {
    const from = `${o.fy}-07-01`, to = `${Number(o.fy) + 1}-06-30`;
    const kinds = new Map();
    let hoursTaken = 0, booked = 0;
    for (const lr of leaveRows) {
      if (lr.to_date < from || lr.from_date > to) continue;
      const label = lr.type === "other" ? (lr.type_name || "Other Leave") : (LEAVE_LABELS[lr.type] || lr.type_name || "Leave");
      for (const d of leavePlanOf(lr, dayLen)) {
        if (d.date < from || d.date > to || !(d.hours > 0)) continue;
        const frac = Math.min(1, d.hours / dayLen);
        if (d.date > todayIso) { booked += frac; continue; }
        const k = kinds.get(label) || { label, days: 0, hours: 0 };
        k.days += frac; k.hours += d.hours; hoursTaken += d.hours;
        kinds.set(label, k);
      }
    }
    return { fy: o.fy, label: o.label, closed: o.closed, runs: o.runs, firstPaid: o.firstPaid, lastPaid: o.lastPaid, gross: money2(o.gross), tax: money2(o.tax), super: money2(o.super), net: money2(o.net), worked: money2(o.worked), leave: [...kinds.values()].map((k) => ({ label: k.label, days: r1(k.days), hours: money2(k.hours) })).sort((a, b) => a.label.localeCompare(b.label)), leaveHours: money2(hoursTaken), bookedDays: r1(booked), dayHours: dayLen };
  }).sort((a, b) => b.fy.localeCompare(a.fy));
  } catch (e) { console.log("fy stats error", String(e && e.message || e).slice(0, 200)); fyStats = []; }
  const fyHours = [...fyMap.values()].map((x) => ({ label: x.label, worked: money2(x.worked), leave: money2(x.leave), runs: x.runs })).sort((a, b) => b.label.localeCompare(a.label));
  if (!rows.length) {
    const known = await env.DB.prepare("SELECT employee_id FROM xero_payroll WHERE email = ?").bind(email).first();
    if (!known) return { found: false, reason: NOT_IN_RESERVOIR };
  }
  return { found: true, slips: rows.map((r) => ({ id: r.run_id, start: r.period_start, end: r.period_end, paid: r.payment_date, net: typeof r.net === "number" ? r.net : null, filed: r.filed !== 0 })), partial: false, fyHours, fyStats };
}
// One stored pay slip for a person, with year to date added up from their stored slips and their own bank details and job title from dfp_xero. Null when there is none.
async function payslipOneFor(env, email, runId) {
  await dfpTables(env);
  await payslipTables(env);
  const row = await env.DB.prepare("SELECT slip FROM payslip_store WHERE email = ? AND run_id = ?").bind(email, runId).first();
  if (!row) return null;
  const out = JSON.parse(row.slip);
  out.ytd = await payslipYtd(env, email, out.paymentDate);
  const me = await env.DB.prepare("SELECT job_title, bank_account_name, bsb, account_number FROM dfp_xero WHERE email = ?").bind(email).first();
  if (me) {
    const acct = String(me.account_number || "").replace(/\s+/g, "");
    out.bank = { accountName: me.bank_account_name || "", bsb: me.bsb || "", masked: acct ? `XXXX ${acct.slice(-4)}` : "" };
    if (me.job_title) out.employee = { ...(out.employee || {}), jobTitle: me.job_title };
  }
  return out;
}
// Year to date is added up from the stored pay slips, so it needs no Xero call. Xero gives no year to date figure through its API.
// The rule matches Xero's own pay slips: a whole pay run counts by its payment date, and a payment date on or after 1 July starts the new financial year. No pay run is split by days.
// It is complete only when every pay day of the financial year is held and Xero gave every main total on each slip.
async function payslipYtd(env, email, paymentDate) {
  const fy = financialYear(paymentDate);
  const r = await env.DB.prepare("SELECT COALESCE(SUM(gross),0) g, COALESCE(SUM(tax),0) t, COALESCE(SUM(super_amt),0) s, COALESCE(SUM(net),0) n, COUNT(*) c, COALESCE(SUM(missing),0) m FROM payslip_store WHERE email = ? AND payment_date >= ? AND payment_date <= ?").bind(email, fy.start, paymentDate).first();
  const gaps = Number(r.m) || 0;
  return { gross: money2(r.g), tax: money2(r.t), super: money2(r.s), net: money2(r.n), complete: fy.start >= PAYSLIP_FULL_FY && gaps === 0, runs: Number(r.c) || 0, gaps, from: PAYSLIP_FROM };
}

// ---- DFP Xero and Xero Payroll ----
// dfp_xero is every field that goes back and forth between DFP and Xero: name, address, date of birth, mobile, start date, job title and bank details. Preferred name, time of birth and the emergency contact never go to Xero, so they live in dfp_reservoir. It is keyed by work email, and a new employee is always created in Xero first. Bank details are never read back from Xero and never returned by any route.
// It is DFP's own data that DFP feeds to Xero. Xero only needs these for payroll and compliance. DFP is the controlling source, and any sync (the Saturday sweep, not built yet) pushes DFP's values out to Xero rather than reading Xero's copy back in as the authority. The table was seeded once from what was already sitting in Xero's employee records, because that data existed there before this design. That seeding uses INSERT OR IGNORE, so running it again never overwrites a saved or approved change.
// Everything else DFP holds about a person lives in dfp_reservoir, which Xero never reads or writes, and which is never read by any other program.
// xero_payroll mirrors Xero's payroll side: classification, rates, overtime entitlement, leave loading, super and leave balances. Only the payroll complete trigger and the Saturday sweep should write to it (neither is built yet). Each write stamps synced_at and source, and xero_payroll_log is where the sweep records what changed.
// Tax file numbers are never read or stored. Bank details are never read from Xero.
// Per period figures (that week's hours, earnings lines, wages, tax, net) are for the pay run posted trigger to fill and clear after the PDF is filed. The standing values stay as the current snapshot: rates, leave balances, super fund and the latest super amount.
let dfpTablesMade = false;
async function dfpTables(env) {
  if (dfpTablesMade) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS dfp_xero (email TEXT PRIMARY KEY, xero_employee_id TEXT, name TEXT, address TEXT, dob TEXT, mobile TEXT, start_date TEXT, job_title TEXT, bank_account_name TEXT, bsb TEXT, account_number TEXT, status TEXT NOT NULL DEFAULT 'ACTIVE', source TEXT, created_at TEXT DEFAULT (datetime('now')), updated_at TEXT DEFAULT (datetime('now')), updated_by TEXT)").run();
  // A dfp_xero made by an earlier build has no bank columns. Add them if they are missing.
  for (const col of ["bank_account_name TEXT", "bsb TEXT", "account_number TEXT", "middle_names TEXT"]) { try { await env.DB.prepare(`ALTER TABLE dfp_xero ADD COLUMN ${col}`).run(); } catch (e) { /* already there */ } }
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS dfp_xero_employee ON dfp_xero (xero_employee_id)").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_payroll (employee_id TEXT PRIMARY KEY, email TEXT, name TEXT, status TEXT, classification TEXT, employment_basis TEXT, rate_ordinary REAL, rate_ot15 REAL, rate_ot2 REAL, has_overtime INTEGER NOT NULL DEFAULT 0, rates_note TEXT, leave_loading INTEGER, super_fund_name TEXT, super_member_number TEXT, super_usi TEXT, leave_balances TEXT, allowances TEXT, super_amount_period REAL, last_pay_run_id TEXT, synced_at TEXT, source TEXT)").run();
  // A table made by the first w31 has no allowances or super amount columns. Add them if they are missing.
  for (const col of ["allowances TEXT", "super_amount_period REAL", "annual_salary REAL", "week_hours REAL", "super_pct REAL", "super_run_date TEXT"]) { try { await env.DB.prepare(`ALTER TABLE xero_payroll ADD COLUMN ${col}`).run(); } catch (e) { /* already there */ } }
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS xero_payroll_email ON xero_payroll (email)").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_payroll_log (id INTEGER PRIMARY KEY AUTOINCREMENT, employee_id TEXT, name TEXT, field TEXT, old_value TEXT, new_value TEXT, changed_at TEXT DEFAULT (datetime('now')), trigger_name TEXT)").run();
  dfpTablesMade = true;
}
// The super percentage each person is really being paid, worked out from Xero's own figures: the super guarantee Xero paid divided by the earnings that super is payable on, taken from the person's latest posted payslip. Super is not paid on overtime (or on any earnings rate Xero marks exempt from super), so those lines are left out of the base. Dividing by total wages, as an earlier build did, showed 9% or 11.1% for anyone who worked overtime that week. Nothing is stored as a rate, so if the legislated rate moves it shows up on the next pay run. Looks back up to six posted runs for anyone not on the newest. Returns Map(employeeId -> { pct, date }).
async function xeroSuperPcts(env, ids) {
  const want = new Set(ids);
  const found = new Map();
  let calls = 0;
  if (!want.size) return { found, calls };
  const posted = await xeroPostedRuns(env);
  calls++;
  if (posted.fail) return { found, calls, why: posted.fail.error || "Xero refused the pay runs" };
  const items = await xeroItems(env);
  calls++;
  const rateOf = new Map(((items && items.EarningsRates) || []).map((x) => [x.EarningsRateID, x]));
  const exempt = (id) => {
    const r = rateOf.get(id);
    if (!r) return false;
    if (r.IsExemptFromSuper === true) return true;
    if (r.IsExemptFromSuper === false) return false;
    return /OVERTIME/i.test(String(r.EarningsType || ""));
  };
  for (const run of posted.runs.slice(0, 6)) {
    if (found.size >= want.size) break;
    const rd = await xeroRunDetail(env, run.PayRunID);
    calls++;
    if (rd.fail || !rd.run) continue;
    const todo = (rd.run.Payslips || []).filter((sm) => want.has(sm.EmployeeID) && !found.has(sm.EmployeeID) && sm.PayslipID);
    for (let i = 0; i < todo.length; i += 3) {
      await Promise.all(todo.slice(i, i + 3).map(async (sm) => {
        calls++;
        try {
          const r = await xeroFetch(env, `payroll.xro/1.0/Payslip/${sm.PayslipID}`);
          if (!r.ok) return;
          const j = await r.json();
          const slip = Array.isArray(j.Payslip) ? j.Payslip[0] : j.Payslip || (j.Payslips || [])[0];
          if (!slip) return;
          const lines = [...(slip.EarningsLines || []), ...(slip.TimesheetEarningsLines || []), ...(slip.LeaveEarningsLines || [])];
          let base = 0;
          for (const l of lines) {
            if (exempt(l.EarningsRateID)) continue;
            base += xnum(l.FixedAmount) || xnum(l.Amount) || xunits(l.NumberOfUnits) * xnum(l.RatePerUnit);
          }
          const sgcLines = (slip.SuperannuationLines || []).filter((l) => String(l.ContributionType || "").replace(/[\s_-]/g, "").toUpperCase() === "SGC");
          const sup = sgcLines.length ? sgcLines.reduce((a, l) => a + xnum(l.Amount), 0) : xnum(slip.Super) || xnum(sm.Super);
          if (base > 0 && sup > 0) found.set(sm.EmployeeID, { pct: Math.round((sup / base) * 1000) / 10, date: xeroDate(run.PaymentDate) });
        } catch (e) { /* this person is tried again on the next older run */ }
      }));
    }
  }
  return { found, calls };
}
// Seeding only: reads every ACTIVE employee's record as it already sits in Xero (3 at a time) and splits it into one dfp_xero row and one xero_payroll row each. Used for the one time fill. It is not how DFP's data is kept up to date, DFP pushes to Xero.
// Returns { fail } or { people, calls }. Nothing is written here.
async function dfpReadFromXero(env) {
  const got = await xeroEmployeeList(env);
  if (got.fail) return { fail: got.fail };
  const active = got.list.filter((x) => String(x.Status || "ACTIVE").toUpperCase() === "ACTIVE");
  let calls = 1;
  const pr = await xeroFetch(env, "payroll.xro/1.0/PayItems");
  calls++;
  if (!pr.ok) return { fail: { status: pr.status, detail: await xeroWhy(pr) } };
  const earnings = ((await pr.json()).PayItems || {}).EarningsRates || [];
  // The super funds list, and the product name of each fund that has no name, are read once here so the people fetched in parallel below do not each ask again.
  const kept = { funds: [], why: "", products: {} };
  try {
    const sr = await xeroFetch(env, "payroll.xro/1.0/Superfunds");
    calls++;
    if (sr.ok) kept.funds = (await sr.json()).SuperFunds || [];
    else kept.why = `Xero refused the super funds list (${sr.status})`;
  } catch (e) { kept.why = "Could not read the super funds list from Xero"; }
  for (const f of kept.funds) {
    const usi = String(f.USI || "").trim();
    if (!usi || String(f.Name || "").trim() || usi in kept.products) continue;
    kept.products[usi] = "";
    try {
      const r = await xeroFetch(env, `payroll.xro/1.0/SuperfundProducts?USI=${encodeURIComponent(usi)}`);
      calls++;
      if (r.ok) { const prod = ((await r.json()).SuperFundProducts || [])[0]; if (prod && prod.ProductName) kept.products[usi] = String(prod.ProductName).trim(); }
    } catch (e) { /* leave the name blank */ }
  }
  const people = await xeroEach(active, async (item) => {
    let full = null;
    try {
      const r = await xeroFetch(env, `payroll.xro/1.0/Employees/${item.EmployeeID}`);
      calls++;
      if (r.ok) full = ((await r.json()).Employees || [])[0] || null;
    } catch (e) { /* use the list entry below */ }
    const record = full || item;
    const row = xeroStaffRow(record);
    const email = String(row.email || "").trim().toLowerCase();
    const rates = xeroRates(record, earnings);
    const sup = await xeroSuperFor(env, record, kept);
    const td = record.TaxDeclaration || {};
    const balances = (Array.isArray(record.LeaveBalances) ? record.LeaveBalances : []).map((b) => ({ name: String(b.LeaveName || ""), typeId: String(b.LeaveTypeID || ""), units: xnum(b.NumberOfUnits), unitType: String(b.TypeOfUnits || "") }));
    const fund = (sup.memberships || [])[0] || {};
    // Allowances on the pay template: earnings lines whose earnings rate is of type ALLOWANCE.
    const rateById = new Map(earnings.map((x) => [x.EarningsRateID, x]));
    const allowances = ((record.PayTemplate && record.PayTemplate.EarningsLines) || []).filter((l) => { const er = rateById.get(l.EarningsRateID); return er && String(er.EarningsType || "").toUpperCase() === "ALLOWANCE"; }).map((l) => ({ name: String((rateById.get(l.EarningsRateID) || {}).Name || ""), rate: l.RatePerUnit === undefined || l.RatePerUnit === null ? null : xnum(l.RatePerUnit), amount: l.FixedAmount === undefined || l.FixedAmount === null ? null : xnum(l.FixedAmount), calc: String(l.CalculationType || "") }));
    return {
      xero: { email, xeroId: row.id, name: row.name, address: row.address, dob: row.dob, mobile: row.mobile, startDate: row.commenced, jobTitle: row.title, bank: { name: row.accountName, bsb: row.bsb, number: row.accountNumber }, em: { name: row.emName, relationship: row.emRelationship, mobile: row.emMobile, email: row.emEmail } },
      payroll: { employeeId: row.id, email, name: row.name, status: row.status || "ACTIVE", classification: String(record.Classification || "").trim(), basis: xeroBasis(record), ordinary: rates.rates.ordinary, annual: rates.rates.annual, weekHours: rates.rates.weekHours, ot15: rates.rates.ot15, ot2: rates.rates.ot2, hasOvertime: rates.hasOvertime ? 1 : 0, ratesNote: rates.why.join("; "), leaveLoading: typeof td.EligibleToReceiveLeaveLoading === "boolean" ? (td.EligibleToReceiveLeaveLoading ? 1 : 0) : null, fundName: fund.fundName || "", memberNumber: fund.memberNumber || "", usi: fund.usi || "", balances: JSON.stringify(balances), balanceCount: balances.length, allowances: JSON.stringify(allowances), allowanceCount: allowances.length }
    };
  });
  let superWhy = "";
  try {
    const sp = await xeroSuperPcts(env, people.map((x) => x.payroll.employeeId).filter(Boolean));
    calls += sp.calls;
    superWhy = sp.why || "";
    for (const x of people) { const hit = sp.found.get(x.payroll.employeeId); if (hit) { x.payroll.superPct = hit.pct; x.payroll.superRunDate = hit.date; } }
  } catch (e) { superWhy = String(e && e.message || e).slice(0, 160); }
  return { people, calls, superWhy };
}
// Names of the fields that came back empty, never their values.
function dfpGaps(p) {
  const i = p.xero, y = p.payroll;
  const gap = [];
  [["email", i.email], ["address", i.address], ["dob", i.dob], ["mobile", i.mobile], ["start date", i.startDate], ["job title", i.jobTitle], ["bank account name", (i.bank || {}).name], ["bsb", (i.bank || {}).bsb], ["bank account number", (i.bank || {}).number], ["classification", y.classification], ["ordinary rate", y.ordinary], ["super fund", y.fundName], ["super member number", y.memberNumber]].forEach(([k, v]) => { if (v === "" || v === null || v === undefined) gap.push(k); });
  return gap;
}


// ---- Phone numbers. One style everywhere, the same as fmtPhone in the app. Mobiles (04), 1300 and 1800 read xxxx xxx xxx, other ten digit numbers xx xxxx xxxx, six digit 13 numbers 13 xx xx. Anything else is kept as typed. ----
function normPhone(v) {
  const raw = String(v === null || v === undefined ? "" : v).trim();
  if (!raw) return raw;
  let d = raw.replace(/[^0-9+]/g, "");
  if (/^\+61\d{9}$/.test(d)) d = `0${d.slice(3)}`;
  else if (/^61\d{9}$/.test(d)) d = `0${d.slice(2)}`;
  if (/^13\d{4}$/.test(d)) return `${d.slice(0, 2)} ${d.slice(2, 4)} ${d.slice(4)}`;
  if (!/^\d{10}$/.test(d)) return raw;
  if (d.startsWith("04") || d.startsWith("1300") || d.startsWith("1800")) return `${d.slice(0, 4)} ${d.slice(4, 7)} ${d.slice(7)}`;
  return `${d.slice(0, 2)} ${d.slice(2, 6)} ${d.slice(6)}`;
}
// Moves every number already stored into that style. Runs once when the Worker first starts after a deploy, and again on a later start only changes a number that is still in the old style.
let phonesChecked = false;
async function phonesNormalise(env) {
  if (phonesChecked) return;
  phonesChecked = true;
  await reservoirTables(env);
  let changed = 0;
  for (const [tbl, col, idCol] of [["dfp_xero", "mobile", "email"], ["dfp_reservoir", "em_mobile", "person_key"]]) {
    const rows = ((await env.DB.prepare(`SELECT ${idCol} AS id, ${col} AS v FROM ${tbl} WHERE ${col} IS NOT NULL AND ${col} <> ''`).all()).results) || [];
    for (const r of rows) {
      const n = normPhone(r.v);
      if (n !== r.v) { await env.DB.prepare(`UPDATE ${tbl} SET ${col} = ? WHERE ${idCol} = ?`).bind(n, r.id).run(); changed++; }
    }
  }
  try {
    const rows = ((await env.DB.prepare("SELECT id, servicePhone AS v FROM fleet WHERE servicePhone IS NOT NULL AND servicePhone <> ''").all()).results) || [];
    for (const r of rows) {
      const n = normPhone(r.v);
      if (n !== r.v) { await env.DB.prepare("UPDATE fleet SET servicePhone = ? WHERE id = ?").bind(n, r.id).run(); changed++; }
    }
  } catch (e) { console.log("phones: fleet not checked", String(e && e.message || e).slice(0, 120)); }
  if (changed) console.log("phones moved to the standard style:", changed);
}
// ---- Leave. The app reads only this table (xero_leave). Xero is written to only when the office approves a request or applies a change. ----
// Balances come from xero_payroll.leave_balances. History is brought in from Xero by the leave sync (leaveSync) and by nothing else.
let leaveTableMade = false;
async function leaveTable(env) {
  if (leaveTableMade) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_leave (id TEXT PRIMARY KEY, email TEXT NOT NULL, name TEXT, employee_id TEXT, type TEXT NOT NULL, type_name TEXT, leave_type_id TEXT, from_date TEXT NOT NULL, to_date TEXT NOT NULL, status TEXT NOT NULL, cancel_requested INTEGER NOT NULL DEFAULT 0, edit_requested INTEGER NOT NULL DEFAULT 0, new_type TEXT, new_from TEXT, new_to TEXT, xero_id TEXT, source TEXT, xero_note TEXT, created_at TEXT DEFAULT (datetime('now')), decided_by TEXT, decided_at TEXT)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS xero_leave_email ON xero_leave (email, from_date)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS xero_leave_xero ON xero_leave (xero_id)").run();
  for (const col of ["asked_type TEXT", "days TEXT", "split_of TEXT"]) {
    try { await env.DB.prepare(`ALTER TABLE xero_leave ADD COLUMN ${col}`).run(); } catch (e) { /* already there */ }
  }
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_leave_apps (xero_id TEXT PRIMARY KEY, leave_id TEXT NOT NULL)").run();
  leaveTableMade = true;
}
const LEAVE_LABELS = { annual: "Annual Leave", personal: "Personal/Carer's Leave", unpaid: "Personal/Carer's Leave (Unpaid)", compassionate: "Compassionate Leave (Unpaid)" };
const leaveTypeOf = (t) => (t === "personal" || t === "annual" || t === "compassionate" ? t : "");
function leaveKind(name) {
  const n = String(name || "");
  if (/compassionate/i.test(n)) return "compassionate";
  if (/unpaid/i.test(n) && /personal|carer|sick/i.test(n)) return "unpaid";
  if (/annual/i.test(n)) return "annual";
  if (/personal|carer|sick/i.test(n)) return "personal";
  return "other";
}
const leaveIsoOk = (s) => /^\d{4}-\d{2}-\d{2}$/.test(String(s || "")) && !isNaN(new Date(`${s}T00:00:00Z`).getTime());
// A working day is 8 hours for everyone except people on the Plumbing and Fire Sprinklers Award, who are on 7.6. Clerical, casual and no award are all 8. A day figure typed on a person\u2019s Staff page overrides this.
const LEAVE_DAY_HOURS = 8;
const PLUMBING_DAY_HOURS = 7.6;
const isPlumbingAward = (c) => /plumbing|MA000036/i.test(String(c || ""));
// Contracted hours in a day are kept per person (dayHours on their Reservoir record). Anyone without one is on 8 hours, or 7.6 on the plumbing award.
const leaveDayNum = (v) => { const n = Number(v); return n >= 1 && n <= 12 ? n : LEAVE_DAY_HOURS; };
async function leaveDayMap(env) {
  await reservoirTables(env);
  await dfpTables(env);
  const m = new Map();
  const pay = ((await env.DB.prepare("SELECT email, name, classification FROM xero_payroll").all()).results) || [];
  for (const r of pay) {
    const h = isPlumbingAward(r.classification) ? PLUMBING_DAY_HOURS : LEAVE_DAY_HOURS;
    if (r.email) m.set(`c:${String(r.email).toLowerCase()}`, h);
    if (r.name) m.set(`ck:${profileKey(r.name)}`, h);
  }
  const rows = ((await env.DB.prepare("SELECT email, person_key, day_hours FROM dfp_reservoir WHERE day_hours IS NOT NULL AND day_hours <> ''").all()).results) || [];
  for (const r of rows) {
    const h = leaveDayNum(r.day_hours);
    if (r.email) m.set(String(r.email).toLowerCase(), h);
    if (r.person_key) m.set(`k:${r.person_key}`, h);
  }
  return m;
}
const leaveDayFrom = (m, email, name) => {
  if (!m) return LEAVE_DAY_HOURS;
  const e = String(email || "").toLowerCase(), k = profileKey(name);
  return m.get(e) || m.get(`k:${k}`) || m.get(`c:${e}`) || m.get(`ck:${k}`) || LEAVE_DAY_HOURS;
};
async function leaveDayHours(env, email) {
  const m = await leaveDayMap(env);
  const pay = await env.DB.prepare("SELECT name FROM xero_payroll WHERE lower(email) = ?").bind(String(email || "").toLowerCase()).first();
  return leaveDayFrom(m, email, pay && pay.name);
}
function leaveDaysOf(r) {
  try { const d = JSON.parse(r.days || "null"); return Array.isArray(d) ? d : null; } catch (e) { return null; }
}
function leaveWorkDays(from, to) {
  const out = [];
  for (let d = from; d <= to && out.length < 400; d = tsAddDays(d, 1)) {
    const dow = new Date(`${d}T00:00:00Z`).getUTCDay();
    if (dow >= 1 && dow <= 5 && !TS_HOLIDAYS[d]) out.push(d);
  }
  return out;
}
const leaveHoursOf = (r, dh) => {
  const d = leaveDaysOf(r);
  const total = d ? d.reduce((a, x) => a + (Number(x.hours) || 0), 0) : leaveWorkDays(r.from_date, r.to_date).length * (dh || LEAVE_DAY_HOURS);
  return Math.round(total * 100) / 100;
};
// The day by day plan as the screens see it: what was planned, or every working day in the range at the person's hours.
const leavePlanOf = (r, dh) => {
  const d = leaveDaysOf(r);
  if (d) return d.map((x) => ({ date: x.date, hours: Math.round((Number(x.hours) || 0) * 10000) / 10000 }));
  return leaveWorkDays(r.from_date, r.to_date).map((date) => ({ date, hours: dh || LEAVE_DAY_HOURS }));
};
// What staff see: a request that was switched to unpaid at the office still shows as what they asked for.
const leaveShownType = (r) => (r.type === "unpaid" && r.asked_type ? r.asked_type : r.type);
function leaveOut(r, dm) {
  const dh = leaveDayFrom(dm, r.email, r.name);
  const shown = leaveShownType(r);
  return { id: r.id, name: r.name || "", email: r.email, type: shown, typeName: shown === r.type ? (r.type_name || LEAVE_LABELS[r.type] || "Leave") : (LEAVE_LABELS[shown] || "Leave"), actualType: r.type, hours: leaveHoursOf(r, dh), dayHours: dh, plan: leavePlanOf(r, dh), splitOf: r.split_of || "", unpaidPart: r.type === "unpaid" && !!r.asked_type, from: r.from_date, to: r.to_date, status: r.status, cancelRequested: Number(r.cancel_requested) === 1, editRequested: Number(r.edit_requested) === 1, newType: r.new_type || "", newFrom: r.new_from || "", newTo: r.new_to || "", createdAt: r.created_at || "", decidedAt: r.decided_at || "", decidedBy: r.decided_by || "", xeroNote: r.xero_note || "" };
}
function leaveActor(user) {
  return user.viewAs && user.realUser ? `${user.realUser.name} (viewing as ${user.name})` : user.name;
}
// Hours on the books for annual and personal leave, from the stored snapshot of the person's Xero leave balances.
async function leaveBalances(env, email) {
  await dfpTables(env);
  const row = await env.DB.prepare("SELECT employee_id, leave_balances, synced_at FROM xero_payroll WHERE email = ?").bind(email).first();
  const out = { found: !!row, hasLeave: false, annualHours: null, personalHours: null, syncedAt: row ? row.synced_at || "" : "", employeeId: row ? row.employee_id : "" };
  if (!row) return out;
  let list;
  try { list = JSON.parse(row.leave_balances || "[]"); } catch (e) { list = []; }
  // Xero only gives a person leave balances when they have a leave template. No balances at all means no leave for that person (a director who has chosen not to take formal leave, for example).
  out.hasLeave = Array.isArray(list) && list.length > 0;
  for (const b of list) {
    const k = leaveKind(b.name);
    const hours = /hour/i.test(b.unitType || "Hours") ? xnum(b.units) : null;
    if (hours === null) continue;
    if (k === "annual") out.annualHours = (out.annualHours || 0) + hours;
    if (k === "personal") out.personalHours = (out.personalHours || 0) + hours;
  }
  return out;
}
async function leaveXeroTypeId(env, row) {
  const pay = await env.DB.prepare("SELECT leave_balances FROM xero_payroll WHERE email = ?").bind(row.email).first();
  let list;
  try { list = JSON.parse((pay && pay.leave_balances) || "[]"); } catch (e) { list = []; }
  const hit = list.find((b) => b.typeId && leaveKind(b.name) === row.type);
  if (hit) return hit.typeId;
  const res = await xeroFetch(env, "payroll.xro/1.0/PayItems");
  if (!res.ok) return "";
  const types = (((await res.json()).PayItems || {}).LeaveTypes) || [];
  const t = types.find((x) => leaveKind(x.Name) === row.type);
  return t ? t.LeaveTypeID : "";
}
async function leaveEmployeeId(env, row) {
  if (row.employee_id) return row.employee_id;
  const pay = await env.DB.prepare("SELECT employee_id FROM xero_payroll WHERE email = ?").bind(row.email).first();
  return pay ? pay.employee_id : "";
}
// Personal leave hours a person can still use: their Xero balance less what their earlier waiting requests already hold.
async function leaveAvail(env, email, kind, exceptId, before) {
  const bal = await leaveBalances(env, email);
  let have = kind === "annual" ? bal.annualHours : bal.personalHours;
  have = Number.isFinite(have) ? have : 0;
  const rows = ((await env.DB.prepare("SELECT * FROM xero_leave WHERE email = ? AND status = 'pending' AND type = ? AND id <> ? AND created_at <= ?").bind(email, kind, exceptId || "", before || "9999").all()).results) || [];
  const dh = await leaveDayHours(env, email);
  const held = rows.reduce((a, r) => a + leaveHoursOf(r, dh), 0);
  return Math.max(0, Math.round((have - held) * 100) / 100);
}
// Builds the day by day plan for a waiting request. Annual leave stays one range. Personal leave uses the hours the person has and any balance of the day goes as a second, unpaid request that only the office sees. No hours at all makes the whole request unpaid.
async function leaveReplan(env, row, kind, from, to) {
  await env.DB.prepare("DELETE FROM xero_leave WHERE split_of = ? AND status = 'pending'").bind(row.id).run();
  const asked = row.asked_type || kind;
  if (kind !== "personal") {
    await env.DB.prepare("UPDATE xero_leave SET type = ?, type_name = ?, from_date = ?, to_date = ?, days = NULL, asked_type = ? WHERE id = ?").bind(kind, LEAVE_LABELS[kind], from, to, asked, row.id).run();
    return { ok: true };
  }
  const work = leaveWorkDays(from, to);
  if (!work.length) {
    await env.DB.prepare("UPDATE xero_leave SET type = ?, type_name = ?, from_date = ?, to_date = ?, days = NULL, asked_type = ? WHERE id = ?").bind(kind, LEAVE_LABELS[kind], from, to, asked, row.id).run();
    return { ok: true };
  }
  let avail = await leaveAvail(env, row.email, "personal", row.id, row.created_at);
  const DAY = await leaveDayHours(env, row.email);
  const paid = [], unpaid = [];
  for (const date of work) {
    const take = Math.round(Math.min(DAY, avail) * 100) / 100;
    avail = Math.round((avail - take) * 100) / 100;
    if (take > 0) paid.push({ date, hours: take });
    const left = Math.round((DAY - take) * 100) / 100;
    if (left > 0) unpaid.push({ date, hours: left });
  }
  if (!paid.length) {
    await env.DB.prepare("UPDATE xero_leave SET type = 'unpaid', type_name = ?, from_date = ?, to_date = ?, days = ?, asked_type = 'personal' WHERE id = ?").bind(LEAVE_LABELS.unpaid, from, to, JSON.stringify(unpaid), row.id).run();
    return { ok: true, unpaidOnly: true };
  }
  await env.DB.prepare("UPDATE xero_leave SET type = 'personal', type_name = ?, from_date = ?, to_date = ?, days = ?, asked_type = 'personal' WHERE id = ?").bind(LEAVE_LABELS.personal, from, to, JSON.stringify(paid), row.id).run();
  if (unpaid.length) {
    const id = `lv_${crypto.randomUUID().replace(/-/g, "").slice(0, 16)}`;
    await env.DB.prepare("INSERT INTO xero_leave (id, email, name, employee_id, type, type_name, from_date, to_date, status, source, asked_type, days, split_of, created_at) VALUES (?, ?, ?, ?, 'unpaid', ?, ?, ?, 'pending', 'reservoir', 'personal', ?, ?, ?)").bind(id, row.email, row.name || "", row.employee_id || null, LEAVE_LABELS.unpaid, unpaid[0].date, unpaid[unpaid.length - 1].date, JSON.stringify(unpaid), row.id, row.created_at || new Date().toISOString().replace("T", " ").slice(0, 19)).run();
  }
  return { ok: true, split: unpaid.length > 0 };
}
async function leaveCreate(env, o) {
  const id = `lv_${crypto.randomUUID().replace(/-/g, "").slice(0, 16)}`;
  const now = new Date().toISOString().replace("T", " ").slice(0, 19);
  await env.DB.prepare("INSERT INTO xero_leave (id, email, name, employee_id, type, type_name, from_date, to_date, status, source, asked_type, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'pending', 'reservoir', ?, ?)").bind(id, o.email, o.name, o.employeeId, o.type, LEAVE_LABELS[o.type], o.from, o.to, o.type, now).run();
  const row = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(id).first();
  await leaveReplan(env, row, o.type, o.from, o.to);
  return id;
}
// Rejects in Xero every leave application made for a request (one per day, or the single older id). Returns { fail } with the first reason Xero gave when any are refused.
async function leaveRejectInXero(env, row) {
  const ids = new Set();
  (leaveDaysOf(row) || []).forEach((d) => { if (d.xeroId) ids.add(d.xeroId); });
  if (row.xero_id) ids.add(row.xero_id);
  let fail = "";
  for (const id of ids) {
    try {
      const res = await xeroFetch(env, `payroll.xro/1.0/LeaveApplications/${id}/reject`, { method: "POST", headers: { "Content-Type": "application/json" } });
      if (!res.ok) { if (!fail) fail = await xeroWhy(res); continue; }
      let j;
      try { j = await res.json(); } catch (e) { j = {}; }
      const errs = (((j.LeaveApplications || [])[0] || {}).ValidationErrors || []).map((x) => x.Message).filter(Boolean);
      if (errs.length && !fail) fail = errs.join(" ");
    } catch (e) { if (!fail) fail = String(e && e.message || e).slice(0, 160); }
  }
  return fail ? { fail } : {};
}
// Sends a day by day request to Xero, one leave application a day, so Xero totals the weeks itself. A part day carries its hours. Days already sent keep their id so a second try never sends them twice.
async function leavePushDays(env, row) {
  const employeeId = await leaveEmployeeId(env, row);
  if (!employeeId) return { fail: "This person is not in the Xero payroll snapshot yet, so Xero cannot be told. Run the one time fill first." };
  const typeId = await leaveXeroTypeId(env, row);
  if (!typeId) return { fail: `Xero has no ${LEAVE_LABELS[row.type] || "matching"} leave type for this person.` };
  const days = leaveDaysOf(row) || [];
  const DAY = await leaveDayHours(env, row.email);
  for (const d of days) {
    if (d.xeroId) continue;
    const monday = tsAddDays(d.date, -((new Date(`${d.date}T00:00:00Z`).getUTCDay() + 6) % 7));
    const app = { EmployeeID: employeeId, LeaveTypeID: typeId, Title: LEAVE_LABELS[row.type] || "Leave", Description: "Requested in Reservoir", StartDate: `${d.date}T00:00:00`, EndDate: `${d.date}T00:00:00`, PayOutType: "DEFAULT" };
    if (Math.abs(Number(d.hours) - DAY) > 0.001) app.LeavePeriods = [{ PayPeriodStartDate: `${monday}T00:00:00`, PayPeriodEndDate: `${tsAddDays(monday, 6)}T00:00:00`, NumberOfUnits: Number(d.hours) }];
    const res = await xeroFetch(env, "payroll.xro/1.0/LeaveApplications", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify([app]) });
    if (!res.ok) return { fail: `${d.date}: ${await xeroWhy(res)}`, days };
    let j;
    try { j = await res.json(); } catch (e) { j = {}; }
    const made = (j.LeaveApplications || [])[0];
    if (!made) return { fail: `${d.date}: Xero answered but sent back no leave application.`, days };
    const errs = (made.ValidationErrors || []).map((x) => x.Message).filter(Boolean);
    if (errs.length) return { fail: `${d.date}: Xero refused it: ${errs.join(" ")}`, days };
    d.xeroId = made.LeaveApplicationID || "";
    if (d.xeroId) await env.DB.prepare("INSERT OR REPLACE INTO xero_leave_apps (xero_id, leave_id) VALUES (?, ?)").bind(d.xeroId, row.id).run();
    await env.DB.prepare("UPDATE xero_leave SET days = ? WHERE id = ?").bind(JSON.stringify(days), row.id).run();
  }
  return { id: (days.find((d) => d.xeroId) || {}).xeroId || "", days };
}
// Writes one leave application to Xero: create when there is no xero_id yet, update when there is. Returns { id } or { fail }.
async function leavePushXero(env, row, change) {
  const employeeId = await leaveEmployeeId(env, row);
  if (!employeeId) return { fail: "This person is not in the Xero payroll snapshot yet, so Xero cannot be told. Run the one time fill first." };
  const kind = change.type || row.type;
  const typeId = await leaveXeroTypeId(env, { ...row, type: kind });
  if (!typeId) return { fail: `Xero has no ${LEAVE_LABELS[kind] || "matching"} leave type for this person.` };
  const app = { EmployeeID: employeeId, LeaveTypeID: typeId, Title: LEAVE_LABELS[kind] || row.type_name || "Leave", Description: "Requested in Reservoir", StartDate: `${change.from}T00:00:00`, EndDate: `${change.to}T00:00:00`, PayOutType: "DEFAULT" };
  const updating = !!row.xero_id;
  if (updating) app.LeaveApplicationID = row.xero_id;
  const res = await xeroFetch(env, `payroll.xro/1.0/LeaveApplications${updating ? `/${row.xero_id}` : ""}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify([app]) });
  if (!res.ok) return { fail: await xeroWhy(res) };
  let j;
  try { j = await res.json(); } catch (e) { j = {}; }
  const made = (j.LeaveApplications || [])[0];
  if (!made) return { fail: "Xero answered but sent back no leave application." };
  const errs = (made.ValidationErrors || []).map((x) => x.Message).filter(Boolean);
  if (errs.length) return { fail: `Xero refused it: ${errs.join(" ")}` };
  return { id: made.LeaveApplicationID || row.xero_id || "" };
}
// After Xero changes, the person's balances are read once more so the Leave screen shows the new figure at once. Best effort.
async function leaveRefreshBalances(env, email) {
  try {
    const pay = await env.DB.prepare("SELECT employee_id FROM xero_payroll WHERE email = ?").bind(email).first();
    if (!pay || !pay.employee_id) return;
    const r = await xeroFetch(env, `payroll.xro/1.0/Employees/${pay.employee_id}`);
    if (!r.ok) return;
    const rec = ((await r.json()).Employees || [])[0];
    if (!rec) return;
    const balances = (Array.isArray(rec.LeaveBalances) ? rec.LeaveBalances : []).map((b) => ({ name: String(b.LeaveName || ""), typeId: String(b.LeaveTypeID || ""), units: xnum(b.NumberOfUnits), unitType: String(b.TypeOfUnits || "") }));
    await env.DB.prepare("UPDATE xero_payroll SET leave_balances = ?, synced_at = datetime('now') WHERE email = ?").bind(JSON.stringify(balances), email).run();
  } catch (e) { /* the weekly sync will catch up */ }
}
// The leave sync. Reads every active person's balances and the last 12 months of leave applications from Xero into D1. Nothing here ever writes to Xero.
async function leaveSync(env) {
  await leaveTable(env);
  await dfpTables(env);
  const people = ((await env.DB.prepare("SELECT employee_id, email, name FROM xero_payroll WHERE email IS NOT NULL AND email <> '' AND (status IS NULL OR status = 'ACTIVE')").all()).results) || [];
  const byEmp = new Map(people.map((p) => [p.employee_id, p]));
  let balances = 0;
  await xeroEach(people, async (p) => {
    try {
      const r = await xeroFetch(env, `payroll.xro/1.0/Employees/${p.employee_id}`);
      if (!r.ok) return;
      const rec = ((await r.json()).Employees || [])[0];
      if (!rec) return;
      const list = (Array.isArray(rec.LeaveBalances) ? rec.LeaveBalances : []).map((b) => ({ name: String(b.LeaveName || ""), typeId: String(b.LeaveTypeID || ""), units: xnum(b.NumberOfUnits), unitType: String(b.TypeOfUnits || "") }));
      await env.DB.prepare("UPDATE xero_payroll SET leave_balances = ?, synced_at = datetime('now') WHERE employee_id = ?").bind(JSON.stringify(list), p.employee_id).run();
      balances++;
    } catch (e) { /* skip this person, try again next sync */ }
  });
  const since = melbourneDay(Date.now() - 366 * 86400000);
  const typeNames = new Map();
  try {
    const pi = await xeroFetch(env, "payroll.xro/1.0/PayItems");
    if (pi.ok) (((await pi.json()).PayItems || {}).LeaveTypes || []).forEach((t) => typeNames.set(t.LeaveTypeID, t.Name));
  } catch (e) { /* names fall back to the balances */ }
  for (const p of people) {
    try {
      const pay = await env.DB.prepare("SELECT leave_balances FROM xero_payroll WHERE employee_id = ?").bind(p.employee_id).first();
      JSON.parse((pay && pay.leave_balances) || "[]").forEach((b) => { if (b.typeId && b.name && !typeNames.has(b.typeId)) typeNames.set(b.typeId, b.name); });
    } catch (e) { /* ignore */ }
  }
  let seen = 0, added = 0, updated = 0, pages = 0;
  for (let page = 1; page <= 20; page++) {
    const res = await xeroFetch(env, `payroll.xro/1.0/LeaveApplications?page=${page}`);
    if (!res.ok) return { ok: false, balances, error: await xeroWhy(res) };
    const list = (await res.json()).LeaveApplications || [];
    pages++;
    if (!list.length) break;
    for (const a of list) {
      const p = byEmp.get(a.EmployeeID);
      const from = xeroDate(a.StartDate), to = xeroDate(a.EndDate);
      if (!p || !from || !to || to < since) continue;
      seen++;
      const name = typeNames.get(a.LeaveTypeID) || a.Title || "";
      const kind = leaveKind(name);
      const dayApp = await env.DB.prepare("SELECT leave_id FROM xero_leave_apps WHERE xero_id = ?").bind(a.LeaveApplicationID).first();
      if (dayApp) continue;
      const known = await env.DB.prepare("SELECT id, status FROM xero_leave WHERE xero_id = ?").bind(a.LeaveApplicationID).first();
      if (known) {
        if (known.status === "approved") { await env.DB.prepare("UPDATE xero_leave SET from_date = ?, to_date = ?, type = ?, type_name = ? WHERE id = ?").bind(from, to, kind, name, known.id).run(); updated++; }
        continue;
      }
      await env.DB.prepare("INSERT INTO xero_leave (id, email, name, employee_id, type, type_name, leave_type_id, from_date, to_date, status, xero_id, source, decided_by, decided_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'approved', ?, 'xero', 'Xero', datetime('now'))").bind(`lv_${crypto.randomUUID().replace(/-/g, "").slice(0, 16)}`, p.email, p.name, p.employee_id, kind, name, a.LeaveTypeID || "", from, to, a.LeaveApplicationID).run();
      added++;
    }
    if (list.length < 100) break;
  }
  return { ok: true, balances, leaveSeen: seen, added, updated, pages };
}
// ---- Timesheet history. Xero timesheets from the start of 2026, kept in D1 as a read only record per person (xero_timesheets). Reads from Xero only, never writes. ----
const TIMESHEET_FROM = "2026-01-01";
let timesheetTableMade = false;
async function timesheetTable(env) {
  if (timesheetTableMade) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_timesheets (id TEXT PRIMARY KEY, employee_id TEXT NOT NULL, email TEXT, name TEXT, start_date TEXT NOT NULL, end_date TEXT NOT NULL, status TEXT, total_hours REAL NOT NULL DEFAULT 0, lines TEXT, xero_updated TEXT, synced_at TEXT DEFAULT (datetime('now')))").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS xero_timesheets_email ON xero_timesheets (email, start_date)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS xero_timesheets_emp ON xero_timesheets (employee_id, start_date)").run();
  timesheetTableMade = true;
}
// Brings every Xero timesheet that ends on or after 1 January 2026 into D1, for every person Reservoir knows from Xero payroll. Safe to run again, each timesheet is saved once and refreshed.
async function timesheetSync(env) {
  await timesheetTable(env);
  await dfpTables(env);
  const people = ((await env.DB.prepare("SELECT employee_id, email, name FROM xero_payroll").all()).results) || [];
  const byEmp = new Map(people.map((p) => [p.employee_id, p]));
  const rateNames = new Map();
  try {
    const pi = await xeroFetch(env, "payroll.xro/1.0/PayItems");
    if (pi.ok) (((await pi.json()).PayItems || {}).EarningsRates || []).forEach((r) => rateNames.set(r.EarningsRateID, r.Name));
  } catch (e) { /* lines keep their rate id */ }
  const out = { ok: true, pages: 0, seen: 0, saved: 0, byStatus: {}, people: 0, earliest: "", latest: "", skippedNotStaff: 0, rawCount: 0, rawEarliest: "", rawLatest: "", rawPeople: 0, beforeCutoff: 0 };
  const rawWho = new Set();
  const who = new Set();
  const pending = [];
  // Saved in groups so a full read takes a moment, not a long wait.
  const flush = async () => {
    while (pending.length) {
      const group = pending.splice(0, 40);
      if (typeof env.DB.batch === "function") await env.DB.batch(group);
      else for (const st of group) await st.run();
    }
  };
  // Every page is read with no date filter in the request and the 1 January cut off is applied here, so nothing depends on how Xero treats a filter. raw* figures describe everything Xero holds, to show whether there is anything earlier than what was saved.
  for (let page = 1; page <= 60; page++) {
    const res = await xeroFetch(env, `payroll.xro/1.0/Timesheets?page=${page}`);
    if (!res.ok) {
      await flush();
      const why = await xeroWhy(res);
      const hint = res.status === 401 || res.status === 403 ? " Xero needs the Timesheets permission (payroll.timesheets). Add it to the Xero app, then reconnect Xero from the Development tab." : "";
      return { ...out, ok: false, error: why + hint };
    }
    const list = (await res.json()).Timesheets || [];
    out.pages++;
    if (!list.length) break;
    for (const t of list) {
      const start = xeroDate(t.StartDate), end = xeroDate(t.EndDate);
      if (start && end) {
        out.rawCount++;
        rawWho.add(t.EmployeeID);
        if (!out.rawEarliest || start < out.rawEarliest) out.rawEarliest = start;
        if (!out.rawLatest || end > out.rawLatest) out.rawLatest = end;
      }
      if (!start || !end) continue;
      if (end < TIMESHEET_FROM) { out.beforeCutoff++; continue; }
      const p = byEmp.get(t.EmployeeID);
      if (!p) { out.skippedNotStaff++; continue; }
      const lines = (Array.isArray(t.TimesheetLines) ? t.TimesheetLines : []).map((l) => ({ rateId: String(l.EarningsRateID || ""), rateName: rateNames.get(l.EarningsRateID) || "", units: (Array.isArray(l.NumberOfUnits) ? l.NumberOfUnits : []).map(xnum) }));
      let total = xnum(t.Hours);
      if (!total) total = lines.reduce((a, l) => a + l.units.reduce((x, y) => x + y, 0), 0);
      const um = /\/Date\((-?\d+)/.exec(String(t.UpdatedDateUTC || ""));
      const updated = um ? new Date(Number(um[1])).toISOString() : String(t.UpdatedDateUTC || "");
      pending.push(env.DB.prepare("INSERT INTO xero_timesheets (id, employee_id, email, name, start_date, end_date, status, total_hours, lines, xero_updated, synced_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, datetime('now')) ON CONFLICT(id) DO UPDATE SET employee_id = excluded.employee_id, email = excluded.email, name = excluded.name, start_date = excluded.start_date, end_date = excluded.end_date, status = excluded.status, total_hours = excluded.total_hours, lines = excluded.lines, xero_updated = excluded.xero_updated, synced_at = datetime('now') WHERE xero_timesheets.status <> 'PROCESSED'").bind(t.TimesheetID, t.EmployeeID, p.email || null, p.name || "", start, end, String(t.Status || ""), total, JSON.stringify(lines), updated));
      out.seen++;
      out.saved++;
      const st = String(t.Status || "UNKNOWN");
      out.byStatus[st] = (out.byStatus[st] || 0) + 1;
      who.add(t.EmployeeID);
      if (!out.earliest || start < out.earliest) out.earliest = start;
      if (!out.latest || end > out.latest) out.latest = end;
    }
    if (list.length < 100) break;
  }
  await flush();
  out.people = who.size;
  out.rawPeople = rawWho.size;
  return out;
}
// A read only check of what Xero actually sends for timesheets, to settle whether drafts come through. Asks the older and the newer Payroll API with and without a status filter and reports plainly. Writes nothing.
async function timesheetProbe(env) {
  const people = ((await env.DB.prepare("SELECT employee_id, name FROM xero_payroll").all()).results) || [];
  const nameOf = new Map(people.map((p) => [p.employee_id, p.name]));
  const lines = [];
  const dmy = (iso) => String(iso || "").split("-").reverse().join(".");
  const ask = async (label, path, pick) => {
    try {
      const res = await xeroFetch(env, path);
      if (!res.ok) { lines.push(`${label}: ${await xeroWhy(res)}`); return null; }
      const j = await res.json();
      const list = pick(j);
      return list;
    } catch (e) { lines.push(`${label}: ${String(e && e.message || e).slice(0, 120)}`); return null; }
  };
  const count = (list, f) => list.reduce((m, t) => { const k = String(f(t) || "UNKNOWN"); m[k] = (m[k] || 0) + 1; return m; }, {});
  const show = (m) => Object.entries(m).map(([k, v]) => `${k} ${v}`).join(", ") || "none";
  const v1 = await ask("Older API, everything", "payroll.xro/1.0/Timesheets?page=1", (j) => j.Timesheets || []);
  if (v1) lines.push(`Older API, no filter, first page: ${v1.length} timesheets. By status: ${show(count(v1, (t) => t.Status))}.`);
  const v1d = await ask("Older API, Draft only", `payroll.xro/1.0/Timesheets?where=${encodeURIComponent('Status=="DRAFT"')}`, (j) => j.Timesheets || []);
  if (v1d) {
    lines.push(`Older API, Draft only: ${v1d.length} timesheets.`);
    v1d.slice(0, 12).forEach((t) => lines.push(`Draft: ${nameOf.get(t.EmployeeID) || t.EmployeeID}, ${dmy(xeroDate(t.StartDate))} to ${dmy(xeroDate(t.EndDate))}, ${xnum(t.Hours)} hours.`));
  }
  const v2 = await ask("Newer API, Draft only", "payroll.xro/2.0/Timesheets?status=Draft", (j) => j.timesheets || j.Timesheets || []);
  if (v2) {
    lines.push(`Newer API, Draft only: ${v2.length} timesheets.`);
    v2.slice(0, 12).forEach((t) => lines.push(`Draft (newer): ${nameOf.get(t.employeeId || t.EmployeeID) || t.employeeId || t.EmployeeID}, ${dmy(String(t.startDate || t.StartDate || "").slice(0, 10))} to ${dmy(String(t.endDate || t.EndDate || "").slice(0, 10))}.`));
  }
  const v2a = await ask("Newer API, everything", "payroll.xro/2.0/Timesheets", (j) => j.timesheets || j.Timesheets || []);
  if (v2a) lines.push(`Newer API, no filter: ${v2a.length} timesheets. By status: ${show(count(v2a, (t) => t.status || t.Status))}.`);
  return { ok: true, lines };
}
// The Timesheets screen reads only what is already stored here. It never calls Xero.
// ---- Real access levels, as set on each person in Microsoft Entra (the T1, T2 and T3 app roles on the Reservoir enterprise app). Read through Graph with the app's own login. Needs the Graph app to hold the Directory.Read.All application permission with admin consent. Read only. Kept for five minutes. ----
let accessTiersCache = { at: 0, value: null };
async function accessTiers(env) {
  if (accessTiersCache.value && Date.now() - accessTiersCache.at < 300000) return accessTiersCache.value;
  const get = async (path) => {
    let res = await graphFetch(env, path);
    // A 403 can come from a Graph token fetched before the permission was granted (roles are fixed inside the token). Fetch a fresh one and try once more.
    if (res.status === 403) { graphToken = { value: null, at: 0 }; res = await graphFetch(env, path); }
    const text = await res.text();
    let json;
    try { json = JSON.parse(text); } catch (e) { json = null; }
    if (!res.ok) {
      let who = "";
      if (res.status === 403) {
        try {
          const c = JSON.parse(atob(String(graphToken.value).split(".")[1].replace(/-/g, "+").replace(/_/g, "/")));
          who = ` The Worker's Graph login is app ${c.appid || c.azp || "unknown"} and its token holds: ${(c.roles || []).join(", ") || "no application permissions at all"}.`;
        } catch (e) { who = ""; }
      }
      throw new Error(`Graph answered ${res.status} for ${path.split("?")[0]}. ${json && json.error && json.error.message ? json.error.message : text.slice(0, 160)}${who}`);
    }
    return json || {};
  };
  const sp = (await get(`/servicePrincipals?$filter=appId eq '${env.CLIENT_ID}'&$select=id,appRoles`)).value || [];
  if (!sp.length) throw new Error("The Reservoir enterprise app was not found in Entra.");
  const roleOf = {};
  (sp[0].appRoles || []).forEach((r) => { if (/^T[123]$/.test(r.value || "")) roleOf[r.id] = r.value; });
  const rank = { T1: 1, T2: 2, T3: 3 };
  const tiers = {};
  const give = (u, tier) => {
    for (const k of [u.userPrincipalName, u.mail]) {
      const e = String(k || "").trim().toLowerCase();
      if (e && (!tiers[e] || rank[tier] < rank[tiers[e]])) tiers[e] = tier;
    }
  };
  let next = `/servicePrincipals/${sp[0].id}/appRoleAssignedTo?$top=100`;
  for (let page = 0; next && page < 5; page++) {
    const j = await get(next);
    for (const a of j.value || []) {
      const tier = roleOf[a.appRoleId];
      if (!tier) continue;
      if (a.principalType === "User") give(await get(`/users/${a.principalId}?$select=userPrincipalName,mail`), tier);
      else if (a.principalType === "Group") for (const m of ((await get(`/groups/${a.principalId}/transitiveMembers/microsoft.graph.user?$select=userPrincipalName,mail&$top=200`)).value || [])) give(m, tier);
    }
    next = j["@odata.nextLink"] || "";
  }
  accessTiersCache = { at: Date.now(), value: tiers };
  return tiers;
}
// ---- AroFlo (API v2, personal access token). Read only. T1 only. The token is the Worker secret AROFLO_TOKEN and is never returned. ----
const AROFLO_BASE = "https://api.aroflo.com/v2";
async function arofloGet(env, path) {
  const res = await fetch(AROFLO_BASE + path, { method: "GET", headers: { Authorization: "Bearer " + env.AROFLO_TOKEN, Accept: "application/json" } });
  const text = await res.text();
  let json;
  try { json = JSON.parse(text); } catch (e) { json = null; }
  return { ok: res.ok, status: res.status, json, text: text.slice(0, 1200), contentType: res.headers.get("content-type") || "" };
}
// Flattens any JSON into [path, value] pairs so the app can list every field AroFlo sent, whatever its shape.
function arofloFlat(value, prefix, out, depth) {
  if (out.length > 400) return out;
  if (value === null || value === undefined) { if (prefix) out.push([prefix, ""]); return out; }
  if (Array.isArray(value)) {
    if (!value.length) { if (prefix) out.push([prefix, "(empty list)"]); return out; }
    if (depth >= 4) { out.push([prefix, `(${value.length} items)`]); return out; }
    value.slice(0, 20).forEach((v, i) => arofloFlat(v, `${prefix}[${i}]`, out, depth + 1));
    return out;
  }
  if (typeof value === "object") {
    if (depth >= 4) { out.push([prefix, "(more fields)"]); return out; }
    Object.keys(value).forEach((k) => arofloFlat(value[k], prefix ? `${prefix}.${k}` : k, out, depth + 1));
    return out;
  }
  out.push([prefix, String(value)]);
  return out;
}
// The list of records inside whatever AroFlo returned: a bare list, or the first list found under a key.
function arofloList(json) {
  if (Array.isArray(json)) return json;
  if (json && typeof json === "object") {
    for (const k of ["data", "items", "results", "projects", "records", "rows"]) if (Array.isArray(json[k])) return json[k];
    for (const k of Object.keys(json)) if (Array.isArray(json[k])) return json[k];
    for (const k of Object.keys(json)) if (json[k] && typeof json[k] === "object") { const inner = arofloList(json[k]); if (inner) return inner; }
  }
  return null;
}
// Everything in the reply that is not the list itself (paging details, totals, links), so the way AroFlo pages its answers can be read and followed.
function arofloMeta(json) {
  if (!json || typeof json !== "object" || Array.isArray(json)) return [];
  const out = [];
  for (const k of Object.keys(json)) {
    const v = json[k];
    if (Array.isArray(v)) continue;
    arofloFlat(v, k, out, 1);
  }
  return out.slice(0, 40);
}
// Works out how many pages AroFlo has from its paging details, when it says. Returns 0 when it does not.
function arofloLastPage(meta, listLength) {
  const num = (re) => { const h = meta.find(([k, v]) => re.test(k.split(".").pop()) && v !== "" && Number.isFinite(Number(v))); return h ? Number(h[1]) : 0; };
  const pages = num(/^(total_?pages|page_?count|last_?page|num_?pages|pages|totalPages|pageCount|lastPage)$/i);
  if (pages > 1) return pages;
  const total = num(/^(total|count|total_?count|total_?items|total_?results|totalCount|totalItems|totalResults)$/i);
  const size = num(/^(per_?page|page_?size|limit|perPage|pageSize)$/i) || listLength;
  if (total > size && size > 0) return Math.ceil(total / size);
  return 0;
}
// Picks the project number, client and site address out of one record by field name. Anything not found stays blank, and the full record is always passed on too.
function arofloShape(rec) {
  const flat = arofloFlat(rec, "", [], 0).filter(([, v]) => v !== "" && v !== "(empty list)");
  const find = (want, avoid) => {
    for (const re of want) {
      const hit = flat.find(([k]) => re.test(k) && !(avoid && avoid.test(k)));
      if (hit) return hit[1];
    }
    return "";
  };
  const number = find([/project.*(number|no|num|code|ref)$/i, /(^|\.)(number|projectnumber|jobnumber|ref|reference)$/i, /(^|\.)id$/i]);
  const client = find([/client.*name/i, /customer.*name/i, /(^|\.)client$/i, /(^|\.)customer$/i, /client/i, /customer/i], /address|email|phone|id$/i);
  const parts = flat.filter(([k]) => /(site|location|project|job).*(address|street|suburb|city|town|state|postcode|postal)/i.test(k) && !/client|customer|billing|postal.*client/i.test(k)).map(([, v]) => v);
  let address = parts.length ? Array.from(new Set(parts)).join(", ") : "";
  if (!address) address = find([/(^|\.)address/i, /street/i], /client|customer|billing|email/i);
  const when = flat.find(([k, v]) => /(created|added|loaded|inserted|entered|logged)/i.test(k) && /^\d{4}-\d{2}-\d{2}/.test(v));
  return { number: String(number || ""), client: String(client || ""), address: String(address || ""), when: when ? when[1] : "" };
}
async function arofloDiscover(env) {
  const tries = ["/openapi.json", "/swagger.json", "/docs/json", "/", "/healthcheck"];
  const found = [];
  for (const p of tries) {
    try {
      const r = await arofloGet(env, p);
      const paths = r.json && r.json.paths ? Object.keys(r.json.paths).slice(0, 120) : [];
      found.push({ path: p, status: r.status, paths, text: paths.length ? "" : r.text.slice(0, 300) });
    } catch (e) {
      found.push({ path: p, status: 0, paths: [], text: String(e && e.message || e).slice(0, 120) });
    }
  }
  return found;
}
async function arofloRoutes(request, env, user, url) {
  if (!url.pathname.startsWith("/api/aroflo")) return null;
  if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
  const isPush = url.pathname === "/api/aroflo/push-test" && request.method === "POST";
  if (request.method !== "GET" && !isPush) return reply(env, 404, { error: "not found" });
  if (!env.AROFLO_TOKEN) return reply(env, 200, { ok: false, error: "The AROFLO_TOKEN secret is not set on the Worker.", worker: WORKER_BUILD });
  try {
    // A test only. Finds one project by its number, then tries to add a note and a small test PDF to it, and reports exactly what AroFlo answered for every attempt. Directors only, and only when the button is pressed.
    if (isPush) {
      let body = {};
      try { body = await request.json(); } catch (e) { body = {}; }
      const want = String(body.jobNumber || "").trim();
      if (!want || !/^[A-Za-z0-9_.-]{1,40}$/.test(want)) return reply(env, 400, { error: "Enter a job number" });
      const log = [];
      const note = (m) => log.push(m);
      let pid = "";
      const direct = await arofloGet(env, "/projects/" + encodeURIComponent(want));
      if (direct.ok && direct.json) { pid = want; note(`Found project ${want} directly.`); }
      else {
        note(`Direct lookup of ${want} answered ${direct.status}. Searching the list.`);
        const first = await arofloGet(env, "/projects");
        let lastPage = 1;
        if (first.ok) {
          const l0 = arofloList(first.json) || [];
          lastPage = arofloLastPage(arofloMeta(first.json), l0.length) || 1;
          const hit0 = l0.find((rec) => { const sh = arofloShape(rec); return sh.number === want || String(rec && (rec.id ?? rec.projectId ?? rec.ID)) === want; });
          if (hit0) pid = String(hit0.id ?? hit0.projectId ?? hit0.ID);
        }
        for (let pg = lastPage; !pid && pg >= 1 && pg > lastPage - 8; pg--) {
          const rp = await arofloGet(env, "/projects?page=" + pg);
          const lp = rp.ok ? arofloList(rp.json) : null;
          if (!lp) { note(`Page ${pg} answered ${rp.status}.`); continue; }
          const hit = lp.find((rec) => { const sh = arofloShape(rec); return sh.number === want || String(rec && (rec.id ?? rec.projectId ?? rec.ID)) === want; });
          if (hit) pid = String(hit.id ?? hit.projectId ?? hit.ID);
        }
        note(pid ? `Found project id ${pid} for ${want}.` : `No project matched ${want} in the newest ${Math.min(8, lastPage)} pages.`);
      }
      if (!pid) return reply(env, 200, { ok: false, log, attempts: [], worker: WORKER_BUILD });
      const disc = await arofloDiscover(env);
      const interesting = [];
      disc.forEach((d) => (d.paths || []).forEach((p) => { if (/note|document|attach|file|photo|comment|upload|image/i.test(p)) interesting.push(p); }));
      note(interesting.length ? `AroFlo lists these related paths: ${Array.from(new Set(interesting)).slice(0, 25).join(", ")}` : "AroFlo published no list of paths to read.");
      const attempts = [];
      const send = async (label, path, init) => {
        try {
          const res = await fetch(AROFLO_BASE + path, { ...init, headers: { Authorization: "Bearer " + env.AROFLO_TOKEN, Accept: "application/json", ...(init.headers || {}) } });
          const text = (await res.text()).slice(0, 300);
          attempts.push({ label, path, method: init.method, status: res.status, allow: res.headers.get("allow") || "", body: text });
          return res.ok;
        } catch (e) { attempts.push({ label, path, method: init.method, status: 0, allow: "", body: String(e && e.message || e).slice(0, 200) }); return false; }
      };
      const stamp = new Date().toISOString();
      const msg = `Reservoir test, please ignore. ${stamp}`;
      const enc = encodeURIComponent(pid);
      let noteOk = false;
      for (const key of ["note", "text", "content"]) {
        for (const seg of ["notes", "comments"]) {
          if (noteOk) break;
          noteOk = await send(`note as ${key}`, `/projects/${enc}/${seg}`, { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ [key]: msg }) });
        }
      }
      const pdf = "%PDF-1.4\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj\n2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj\n3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 200 100]/Contents 4 0 R/Resources<</Font<</F1 5 0 R>>>>>>endobj\n4 0 obj<</Length 60>>stream\nBT /F1 12 Tf 10 50 Td (Reservoir test) Tj ET\nendstream endobj\n5 0 obj<</Type/Font/Subtype/Type1/BaseFont/Helvetica>>endobj\ntrailer<</Root 1 0 R>>\n%%EOF";
      let docOk = false;
      for (const seg of ["documents", "attachments", "files", "photos"]) {
        if (docOk) break;
        const fd = new FormData();
        fd.append("file", new Blob([pdf], { type: "application/pdf" }), "Reservoir test.pdf");
        docOk = await send(`file upload to ${seg}`, `/projects/${enc}/${seg}`, { method: "POST", body: fd });
      }
      return reply(env, 200, { ok: noteOk || docOk, noteOk, docOk, projectId: pid, log, attempts, worker: WORKER_BUILD });
    }
    if (url.pathname === "/api/aroflo/test") {
      // Defaults to the healthcheck. A different GET path can be tried with ?path=/something, GET only.
      let path = String(url.searchParams.get("path") || "/healthcheck");
      if (!path.startsWith("/") || path.includes("..") || path.includes("//") || path.includes("\\")) path = "/healthcheck";
      const r = await arofloGet(env, path);
      const shown = r.json ? JSON.stringify(r.json).slice(0, 1500) : r.text;
      return reply(env, 200, { ok: r.ok, status: r.status, path, contentType: r.contentType, body: shown, worker: WORKER_BUILD });
    }
    // The 10 newest projects. Asks AroFlo for the project list, keeps the newest 10 (by a created or added date when the records have one, otherwise by the highest id), and shapes each for the card.
    if (url.pathname === "/api/aroflo/projects") {
      const r = await arofloGet(env, "/projects");
      let list = r.ok ? arofloList(r.json) : null;
      if (!list) {
        return reply(env, 200, { ok: false, status: r.status, path: "/projects", body: r.json ? JSON.stringify(r.json).slice(0, 600) : r.text, discovery: await arofloDiscover(env), worker: WORKER_BUILD });
      }
      // AroFlo hands the list over in pages, oldest first, so the first page holds the oldest projects. When it says how many pages there are, go to the last one (and the ones before it if the last is short) to reach the newest.
      const meta = arofloMeta(r.json);
      const first = list.length;
      const lastPage = arofloLastPage(meta, list.length);
      const steps = [`First page held ${first} projects.`];
      if (lastPage > 1) {
        const pagesRead = [];
        for (const pg of [lastPage, lastPage - 1, lastPage - 2]) {
          if (pg < 1) break;
          if (pagesRead.length && list.length >= 10) break;
          try {
            const rp = await arofloGet(env, "/projects?page=" + pg);
            const lp = rp.ok ? arofloList(rp.json) : null;
            if (lp && lp.length) { pagesRead.push(pg); list = pagesRead.length === 1 ? lp : lp.concat(list); steps.push(`Page ${pg} of ${lastPage} held ${lp.length}.`); }
            else steps.push(`Page ${pg} answered ${rp.status} with no list.`);
          } catch (e) { steps.push(`Page ${pg} failed.`); }
          if (list.length >= 10) break;
        }
      } else steps.push("AroFlo did not say how many pages there are, so only the first page was read.");
      const rows = list.map((rec, i) => ({ rec, i, shape: arofloShape(rec) }));
      const numericId = (x) => { const v = x.rec && typeof x.rec === "object" ? (x.rec.id ?? x.rec.projectId ?? x.rec.ID) : null; return Number(v); };
      const haveDates = rows.filter((x) => x.shape.when).length > rows.length / 2;
      let orderedBy = "list order, newest last";
      if (haveDates) { rows.sort((a, b) => String(b.shape.when).localeCompare(String(a.shape.when))); orderedBy = "created date"; }
      else if (rows.every((x) => Number.isFinite(numericId(x)))) { rows.sort((a, b) => numericId(b) - numericId(a)); orderedBy = "highest id"; }
      else rows.reverse();
      const top = rows.slice(0, 10).map((x) => ({ id: String((x.rec && (x.rec.id ?? x.rec.projectId ?? x.rec.ID)) ?? x.shape.number ?? x.i), number: x.shape.number, client: x.shape.client, address: x.shape.address, when: x.shape.when, fields: arofloFlat(x.rec, "", [], 0) }));
      return reply(env, 200, { ok: true, status: r.status, total: list.length, orderedBy, steps, meta, lastPage, items: top, worker: WORKER_BUILD });
    }
    // A file from AroFlo's document store (a photo or a document), passed through so the app can show it. Directors only, read only.
    if (url.pathname === "/api/aroflo/file") {
      const p = String(url.searchParams.get("path") || "");
      if (!/^\/DocStorage\/[A-Za-z0-9_-]{4,80}$/.test(p)) return reply(env, 400, { error: "bad path" });
      const tried = [];
      for (const base of [AROFLO_BASE, "https://api.aroflo.com", "https://app.aroflo.com"]) {
        try {
          const res = await fetch(base + p, { headers: { Authorization: "Bearer " + env.AROFLO_TOKEN }, redirect: "follow" });
          tried.push(`${base}${p} answered ${res.status}`);
          const ct = res.headers.get("content-type") || "";
          if (res.ok && !/text\/html|application\/json/i.test(ct)) return new Response(res.body, { status: 200, headers: { "Content-Type": ct || "application/octet-stream", "Cache-Control": "private, max-age=300", ...cors(env) } });
          if (res.ok) tried.push(`it was ${ct}, not a file`);
        } catch (e) { tried.push(`${base} failed`); }
      }
      return reply(env, 502, { error: "AroFlo would not give that file", tried });
    }
    // One job by its number: client, site address, then its notes and its documents. Read only. AroFlo's v2 paths are not published, so each possible path is tried and what AroFlo answered is reported.
    if (url.pathname === "/api/aroflo/job") {
      const want = String(url.searchParams.get("number") || "").trim();
      if (!want || !/^[A-Za-z0-9_.-]{1,40}$/.test(want)) return reply(env, 400, { error: "Enter a task number" });
      const steps = [];
      const idOf = (rec) => String(rec && (rec.id ?? rec.projectId ?? rec.taskId ?? rec.ID) || "");
      const numOf = (rec) => { const sh = arofloShape(rec); return sh.number; };
      let coll = "", rec = null, id = "";
      for (const c of ["/tasks", "/projects", "/jobs"]) {
        const r = await arofloGet(env, `${c}/${encodeURIComponent(want)}`);
        steps.push(`${c}/${want} answered ${r.status}`);
        if (r.ok && r.json && typeof r.json === "object") { coll = c; rec = (r.json.data && typeof r.json.data === "object" && !Array.isArray(r.json.data)) ? r.json.data : r.json; id = idOf(rec) || want; break; }
      }
      if (!rec) {
        for (const c of ["/tasks", "/projects", "/jobs"]) {
          const first = await arofloGet(env, c);
          if (!first.ok) { steps.push(`${c} list answered ${first.status}`); continue; }
          const l0 = arofloList(first.json) || [];
          const lastPage = arofloLastPage(arofloMeta(first.json), l0.length) || 1;
          let hit = l0.find((x) => numOf(x) === want || idOf(x) === want);
          for (let pg = lastPage; !hit && pg >= 1 && pg > lastPage - 8 && lastPage > 1; pg--) {
            const rp = await arofloGet(env, `${c}?page=${pg}`);
            const lp = rp.ok ? arofloList(rp.json) : null;
            if (lp) hit = lp.find((x) => numOf(x) === want || idOf(x) === want);
          }
          steps.push(`${c} list searched${hit ? ", found" : ", not found"}`);
          if (hit) { coll = c; rec = hit; id = idOf(hit) || want; break; }
        }
      }
      if (!rec) return reply(env, 200, { ok: false, steps, worker: WORKER_BUILD });
      const shape = arofloShape(rec);
      const jobFlat = arofloFlat(rec, "", [], 0);
      // The client's full name, not its short name or code.
      const nameKeys = jobFlat.filter(([k, v]) => /(client|customer|organi[sz]ation|company)/i.test(k) && /name/i.test(k) && !/(short|code|ref|(^|\.)id$|contact|email|phone|address|nick|abbrev)/i.test(k) && v !== "" && v !== "(empty list)");
      const bestName = nameKeys.find(([k]) => /(legal|full|trading|company|organi[sz]ation).*name|(^|\.)name$/i.test(k)) || nameKeys.sort((a, b) => String(b[1]).length - String(a[1]).length)[0];
      // The site address put in street, suburb, state, postcode order, from whichever group of fields looks like the site.
      const addrParts = jobFlat.filter(([k, v]) => /(address|street|suburb|city|town|locality|state|postcode|post_code|zip|line1|line2|unit|lot)/i.test(k) && !/(client|customer|billing|email|phone|contact|org)/i.test(k) && v !== "" && v !== "(empty list)");
      const groups = {};
      addrParts.forEach(([k, v]) => { const g = k.includes(".") ? k.slice(0, k.lastIndexOf(".")) : ""; (groups[g] = groups[g] || []).push([k.slice(k.lastIndexOf(".") + 1), String(v)]); });
      const gNames = Object.keys(groups).sort((a, b) => (/(site|location|job|task|project)/i.test(b) ? 1 : 0) - (/(site|location|job|task|project)/i.test(a) ? 1 : 0) || groups[b].length - groups[a].length);
      let siteAddress = "";
      if (gNames.length) {
        const g = groups[gNames[0]];
        const rank = (n) => /(line1|address1|street|^address$|unit|lot)/i.test(n) ? 1 : /line2|address2/i.test(n) ? 2 : /(suburb|city|town|locality)/i.test(n) ? 3 : /state/i.test(n) ? 4 : /(post|zip)/i.test(n) ? 5 : 6;
        siteAddress = Array.from(new Set(g.sort((a, b) => rank(a[0]) - rank(b[0])).map((x) => x[1]))).join(", ");
      }
      const pick = (flat, re, skip) => { const f = flat.find(([k, v]) => re.test(k) && !(skip && skip.test(k)) && v !== "" && v !== "(empty list)"); return f ? String(f[1]) : ""; };
      const subs = {};
      for (const seg of ["notes", "documents", "attachments", "files", "photos"]) {
        const r = await arofloGet(env, `${coll}/${encodeURIComponent(id)}/${seg}`);
        const list = r.ok ? arofloList(r.json) : null;
        steps.push(`${coll}/${id}/${seg} answered ${r.status}${list ? `, ${list.length} items` : r.ok ? ", no list" : ""}`);
        if (list) subs[seg] = list.slice(0, 50).map((it) => {
          const flat = arofloFlat(it, "", [], 0);
          const get = (re) => { const f = flat.find(([k, v]) => re.test(k) && v !== "" && v !== "(empty list)"); return f ? String(f[1]) : ""; };
          const person = (pre) => [get(new RegExp(`${pre}.*givenName$`, "i")), get(new RegExp(`${pre}.*familyName$`, "i"))].filter(Boolean).join(" ");
          const plain = (h) => String(h || "").replace(/<br\s*\/?>/gi, "\n").replace(/<\/p>\s*<p[^>]*>/gi, "\n").replace(/<[^>]+>/g, "").replace(/&nbsp;/g, " ").replace(/&amp;/g, "&").replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/\n{3,}/g, "\n\n").trim();
          if (seg === "notes") return { text: plain(get(/(^|\.)content$/i)), by: person("postedBy"), when: get(/(timestampCreated|created)/i), fields: [] };
          return { name: get(/(^|\.)displayName$/i) || get(/(^|\.)name$/i), fileName: get(/(^|\.)fileName$/i), by: person("insertedByUser"), when: get(/datetimeInserted/i), original: get(/(^|\.)original$/i), thumbnail: get(/(^|\.)thumbnail$/i), type: get(/documentType/i), size: Number(get(/fileSizeBytes/i)) || 0, fields: [] };
        });
      }
      return reply(env, 200, { ok: true, collection: coll, id, number: shape.number || want, client: bestName ? String(bestName[1]) : shape.client, address: siteAddress || shape.address, fields: jobFlat.slice(0, 120), notes: subs.notes || [], documents: subs.documents || subs.attachments || subs.files || subs.photos || [], haveNotes: !!subs.notes, haveDocs: !!(subs.documents || subs.attachments || subs.files || subs.photos), steps, worker: WORKER_BUILD });
    }
    // One project pulled raw. Lists every field AroFlo sends for it.
    if (url.pathname === "/api/aroflo/project") {
      const id = String(url.searchParams.get("id") || "").trim();
      if (!id || !/^[A-Za-z0-9_.-]{1,80}$/.test(id)) return reply(env, 400, { error: "No project id" });
      const r = await arofloGet(env, "/projects/" + encodeURIComponent(id));
      if (!r.json) return reply(env, 200, { ok: false, status: r.status, path: "/projects/" + id, body: r.text, worker: WORKER_BUILD });
      return reply(env, 200, { ok: r.ok, status: r.status, path: "/projects/" + id, fields: arofloFlat(r.json, "", [], 0), worker: WORKER_BUILD });
    }
    return reply(env, 404, { error: "not found" });
  } catch (e) {
    return reply(env, 200, { ok: false, error: String(e && e.message || e).slice(0, 300), worker: WORKER_BUILD });
  }
}
// /api/timesheets routes. Office only (T1). Returns a Response, or null when the path is not a timesheets path.

// ---- Timesheet submit and approval. Staff submit a week, the office approves it, and only then does it go to Xero as an approved timesheet. Directors and supervisors (T1 and T2) are trusted, so theirs goes straight through.
// This is the one narrow way the Worker writes to Xero for any tier: static hours go out in one go, and the person only ever hears success or failure. Nothing is read back to them.
let tsTablesMade = false;
async function tsTables(env) {
  if (tsTablesMade) return;
  tsTablesMade = true;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS timesheet_submissions (id TEXT PRIMARY KEY, email TEXT NOT NULL, name TEXT, employee_id TEXT, week_start TEXT NOT NULL, week_end TEXT, days TEXT, total REAL, status TEXT NOT NULL, xero_status TEXT, forwarded INTEGER NOT NULL DEFAULT 0, last_error TEXT, created_at TEXT, decided_at TEXT, decided_by TEXT)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS timesheet_submissions_email ON timesheet_submissions (email, week_start)").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS ts_notices (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT NOT NULL, text TEXT, created_at TEXT, seen INTEGER NOT NULL DEFAULT 0)").run();
}
const tsDmy = (iso) => { const [y, m, d] = String(iso || "").split("-"); return d && m && y ? `${d}.${m}.${y}` : String(iso || ""); };
const tsAddDays = (iso, n) => { const d = new Date(`${iso}T00:00:00Z`); d.setUTCDate(d.getUTCDate() + n); return d.toISOString().slice(0, 10); };
// Today's date in Victoria, close enough for the week checks (UTC plus 11 hours).
const tsToday = () => new Date(Date.now() + 11 * 3600 * 1000).toISOString().slice(0, 10);
function tsWeekProblem(week) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(week)) return "That is not a week start date.";
  const d = new Date(`${week}T00:00:00Z`);
  if (Number.isNaN(d.getTime()) || d.getUTCDay() !== 1) return "A timesheet week must start on a Monday.";
  const today = tsToday();
  const t = new Date(`${today}T00:00:00Z`);
  const thisMonday = tsAddDays(today, -((t.getUTCDay() + 6) % 7));
  if (week > thisMonday) return "That week has not started yet.";
  if (week < tsAddDays(thisMonday, -35)) return "That week is too far back to submit here.";
  return "";
}
function tsCleanDays(days) {
  if (!Array.isArray(days) || days.length < 6 || days.length > 7) return null;
  const out = [];
  for (let i = 0; i < 7; i++) {
    const d = days[i] || {};
    const row = {};
    for (const k of ["ord", "ot15", "ot2"]) {
      const n = Number(d[k] || 0);
      if (!Number.isFinite(n) || n < 0 || n > 24) return null;
      row[k] = Math.round(n * 100) / 100;
    }
    if (row.ord + row.ot15 + row.ot2 > 24) return null;
    out.push(row);
  }
  return out;
}
// A week that is five trade school days (the 7.6 hours the apprentices are paid for) and nothing else is approved without anyone looking at it. The person still presses Submit.
function tsAllTradeSchool(dayTypes, days) {
  if (!Array.isArray(dayTypes) || dayTypes.length < 5) return false;
  for (let i = 0; i < 5; i++) {
    const d = days[i];
    if (dayTypes[i] !== "trade_school" || d.ord !== 7.6 || d.ot15 !== 0 || d.ot2 !== 0) return false;
  }
  return days[5].ord + days[5].ot15 + days[5].ot2 + days[6].ord + days[6].ot15 + days[6].ot2 === 0;
}
const tsTotal = (days) => Math.round(days.reduce((a, d) => a + d.ord + d.ot15 + d.ot2, 0) * 100) / 100;
// Sends one week to Xero. First as Approved. If Xero will not take it that way it is sent as a Draft, and the caller is told. Returns { ok, xeroStatus, note } or { fail }.
async function tsPushXero(env, sub) {
  const res = await xeroFetch(env, `payroll.xro/1.0/Employees/${sub.employee_id}`);
  if (!res.ok) return { fail: `Could not read your pay record from Xero. ${await xeroWhy(res)}` };
  const record = ((await res.json()).Employees || [])[0];
  if (!record) return { fail: "Xero has no pay record for this person." };
  const items = await xeroItems(env);
  const picked = xeroPickRates(record, items.EarningsRates || []);
  const days = JSON.parse(sub.days);
  if (!picked.ordinaryItem) return { fail: "Xero has no ordinary earnings rate for this person." };
  // Salaried people are sent too, as an approved timesheet at their set hours, so Xero holds the hours. Check the draft pay run in Xero before posting that the salary line is unchanged.
  // Public holidays. Xero has a Public Holiday pay rate, so a holiday day goes to Xero on that rate and not as ordinary time. With no such rate in Xero the day is left blank, as the office does by hand.
  const phItem = (items.EarningsRates || []).find((r) => /public\s*holiday/i.test(String(r.Name || "")));
  const isHoliday = days.map((d, i) => !!TS_HOLIDAYS[tsAddDays(sub.week_start, i)]);
  const phUnits = days.map((d, i) => isHoliday[i] ? Math.round(((Number(d.ord) || 0) + (Number(d.ot15) || 0) + (Number(d.ot2) || 0)) * 100) / 100 : 0);
  const phTotal = phUnits.reduce((a, b) => a + b, 0);
  const rest = days.map((d, i) => isHoliday[i] ? { ord: 0, ot15: 0, ot2: 0 } : d);
  const sum = (k) => rest.reduce((a, d) => a + d[k], 0);
  const lines = [];
  const add = (item, k, label) => {
    if (sum(k) <= 0) return "";
    if (!item) return `Xero has no ${label} rate, but this week has ${label} hours.`;
    lines.push({ EarningsRateID: item.EarningsRateID, NumberOfUnits: rest.map((d) => d[k]) });
    return "";
  };
  const problem = add(picked.ordinaryItem, "ord", "ordinary") || add(picked.ot15Item, "ot15", "overtime under 2 hours") || add(picked.ot2Item, "ot2", "overtime over 2 hours");
  if (problem) return { fail: problem };
  const baseLines = lines.slice();
  const send = async (status, withPh) => {
    const sendLines = withPh && phItem && phTotal > 0 ? [...baseLines, { EarningsRateID: phItem.EarningsRateID, NumberOfUnits: phUnits }] : baseLines;
    if (!sendLines.length) return { fail: "There are no hours to send." };
    const body = [{ EmployeeID: sub.employee_id, StartDate: `${sub.week_start}T00:00:00`, EndDate: `${sub.week_end}T00:00:00`, Status: status, TimesheetLines: sendLines }];
    const r = await xeroFetch(env, "payroll.xro/1.0/Timesheets", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(body) });
    if (!r.ok) return { fail: await xeroWhy(r) };
    let j;
    try { j = await r.json(); } catch (e) { j = {}; }
    const made = (j.Timesheets || [])[0];
    if (!made) return { fail: "Xero answered but sent back no timesheet." };
    const errs = (made.ValidationErrors || []).map((x) => x.Message).filter(Boolean);
    if (errs.length) return { fail: `Xero refused it: ${errs.join(" ")}` };
    return { ok: true, xeroStatus: String(made.Status || status).toUpperCase() };
  };
  const phDates = days.map((d, i) => isHoliday[i] && phUnits[i] > 0 ? `${TS_HOLIDAYS[tsAddDays(sub.week_start, i)]} ${tsDmy(tsAddDays(sub.week_start, i))}` : "").filter(Boolean);
  let phNote = "";
  let withPh = phTotal > 0 && !!phItem;
  if (phTotal > 0 && !phItem) phNote = `Xero has no Public Holiday pay rate, so ${phDates.join(", ")} was left blank.`;
  let first = await send("APPROVED", withPh);
  if (!first.ok && withPh) {
    const retry = await send("APPROVED", false);
    if (retry.ok) { phNote = `Xero would not take the Public Holiday line (${first.fail || "refused"}), so ${phDates.join(", ")} was left blank.`; first = retry; withPh = false; }
  }
  if (first.ok) {
    if (withPh) phNote = `${phDates.join(", ")} sent on the ${phItem.Name} rate.`;
    return { ...first, ...phNote ? { note: phNote } : {} };
  }
  const second = await send("DRAFT", withPh);
  if (second.ok) return { ...second, note: `Xero would not take it as approved (${first.fail}) so it was sent as a draft. Approve it in Xero.${phNote ? ` ${phNote}` : ""}` };
  return { fail: second.fail };
}
// The hours a week a salaried person is paid for, from the payroll snapshot. 0 when they are not on a salary.
async function tsSalaryHours(env, email) {
  try { const r = await env.DB.prepare("SELECT annual_salary, week_hours FROM xero_payroll WHERE lower(email) = ?").bind(String(email || "").toLowerCase()).first(); return r && Number(r.annual_salary) > 0 ? Number(r.week_hours) || 0 : 0; } catch (e) { return 0; }
}
// Weeks that were kept in Reservoir only because the person is salaried (before w144) are sent to Xero now. It only touches weeks from 19 Oct 2026 and each row is done once.
async function tsResendSalaried(env) {
  await tsTables(env);
  const rows = ((await env.DB.prepare("SELECT * FROM timesheet_submissions WHERE xero_status = 'SALARIED' AND week_start >= '2026-10-19' LIMIT 10").all()).results) || [];
  for (const row of rows) {
    const out = await tsPushXero(env, row);
    if (out.fail) { await env.DB.prepare("UPDATE timesheet_submissions SET last_error = ? WHERE id = ?").bind(String(out.fail).slice(0, 400), row.id).run(); continue; }
    await env.DB.prepare("UPDATE timesheet_submissions SET xero_status = ?, last_error = NULL WHERE id = ?").bind(out.xeroStatus, row.id).run();
  }
}
// The Xero employee for a signed in person, from the payroll snapshot. Nothing is read from Xero.
async function tsEmployee(env, email) {
  const row = await env.DB.prepare("SELECT employee_id FROM xero_payroll WHERE lower(email) = ?").bind(email).first();
  return row && row.employee_id ? row.employee_id : "";
}
async function tsDuplicate(env, employeeId, email, week) {
  await timesheetTable(env);
  const inXero = await env.DB.prepare("SELECT id FROM xero_timesheets WHERE employee_id = ? AND start_date = ? AND (status IS NULL OR upper(status) <> 'REJECTED') LIMIT 1").bind(employeeId, week).first();
  if (inXero) return "Xero already has a timesheet for that week.";
  const mine = await env.DB.prepare("SELECT status FROM timesheet_submissions WHERE email = ? AND week_start = ? AND status IN ('pending','sent') LIMIT 1").bind(email, week).first();
  if (mine) return mine.status === "pending" ? "That week is already waiting for approval." : "That week has already been sent to Xero.";
  return "";
}
// Only a director may approve a waiting timesheet for now.
// ---- Notifications and the punch list ----
// One table for everything that needs to tell somebody something. A notice goes to one person (their work email) or to the whole office ('@office', seen by every director and manager). A notice with todo = 1 is a job for the person: it stays on their punch list, and keeps counting on the bell, until it is done. Any other notice counts until it has been read.
// data holds what the app needs to act on the notice, for example which card and which photo.
let noticesReady = false;
// ---- General requests (fuel card, replacement tools, test and tag, and the rest). The app keeps each request as one JSON record. Anyone signed in can send one and see their own. The office (T1 and T2) sees them all, and only Tayler can approve or decline.
let appRequestsReady = false;
async function appRequestsTable(env) {
  if (appRequestsReady) return;
  appRequestsReady = true;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS app_requests (id TEXT PRIMARY KEY, by_email TEXT NOT NULL, by_name TEXT, type TEXT, status TEXT, data TEXT NOT NULL, created_at TEXT, updated_at TEXT)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS app_requests_by ON app_requests (by_email)").run();
}
async function appRequestsRoutes(request, env, user, url) {
  const m = url.pathname.match(/^\/api\/requests(?:\/([A-Za-z0-9_-]{1,64}))?$/);
  if (!m) return null;
  const email = String(user.email || "").trim().toLowerCase();
  if (!email) return reply(env, 400, { error: "No email on your sign in" });
  const office = user.tier === "T1" || user.tier === "T2";
  await appRequestsTable(env);
  const out = (r) => { try { return { ...JSON.parse(r.data), id: r.id, server: true }; } catch (e) { return null; } };
  if (request.method === "GET" && !m[1]) {
    const rows = office
      ? ((await env.DB.prepare("SELECT * FROM app_requests ORDER BY created_at DESC LIMIT 300").all()).results) || []
      : ((await env.DB.prepare("SELECT * FROM app_requests WHERE by_email = ? ORDER BY created_at DESC LIMIT 300").bind(email).all()).results) || [];
    return reply(env, 200, { requests: rows.map(out).filter(Boolean), worker: WORKER_BUILD });
  }
  if (request.method === "POST" && !m[1]) {
    let b;
    try { b = await request.json(); } catch (e) { b = null; }
    if (!b || typeof b !== "object" || !/^[A-Za-z0-9_-]{1,64}$/.test(String(b.id || ""))) return reply(env, 400, { error: "Bad request record" });
    const rec = { ...b };
    delete rec.server;
    rec.photos = (Array.isArray(rec.photos) ? rec.photos : []).slice(0, 8).map((x) => ({ name: String(x && x.name || "Photo").slice(0, 120), thumb: x && typeof x.thumb === "string" && x.thumb.length < 120000 ? x.thumb : "" }));
    rec.kind = "request";
    const json = JSON.stringify(rec);
    if (json.length > 900000) return reply(env, 413, { error: "That request is too large" });
    const existing = await env.DB.prepare("SELECT * FROM app_requests WHERE id = ?").bind(rec.id).first();
    const now = new Date().toISOString();
    if (!existing) {
      const first = String(user.name || "").split(" ")[0];
      rec.by = first || rec.by;
      await env.DB.prepare("INSERT INTO app_requests (id, by_email, by_name, type, status, data, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)").bind(rec.id, email, user.name || "", String(rec.type || ""), String(rec.status || "pending"), JSON.stringify(rec), now, now).run();
      return reply(env, 200, { ok: true, created: true });
    }
    if (existing.by_email !== email && !office) return reply(env, 403, { error: "forbidden" });
    const oldStatus = String(existing.status || "pending");
    const newStatus = String(rec.status || oldStatus);
    if (newStatus !== oldStatus && (newStatus === "approved" || newStatus === "declined") && !isApprover(user)) return reply(env, 403, { error: "Only Tayler can approve or decline requests for now" });
    rec.by = (() => { try { return JSON.parse(existing.data).by || rec.by; } catch (e) { return rec.by; } })();
    await env.DB.prepare("UPDATE app_requests SET status = ?, type = ?, data = ?, updated_at = ? WHERE id = ?").bind(newStatus, String(rec.type || existing.type || ""), JSON.stringify(rec), now, rec.id).run();
    if (newStatus !== oldStatus && (newStatus === "approved" || newStatus === "declined")) {
      await notify(env, { to: existing.by_email, kind: "request", title: `Your request for ${String(rec.type || "request").toLowerCase().slice(0, 60)} has been ${newStatus}`, body: "Check your request screen for more details.", link: "makerequest", by: user.name || email });
    }
    // A note from the office sends the requester back a push, and a reply from the requester tells the office.
    try {
      const before = (() => { try { return (JSON.parse(existing.data).notes || []).length; } catch (e) { return 0; } })();
      const notes = Array.isArray(rec.notes) ? rec.notes : [];
      const last = notes[notes.length - 1];
      if (notes.length > before && last && newStatus === "pending") {
        if (last.role === "office") await notify(env, { to: existing.by_email, kind: "request", title: `More information needed for your ${String(rec.type || "request").toLowerCase().slice(0, 60)} request`, body: "Check your request screen for more details.", link: "makerequest", by: user.name || email });
        else await notify(env, { to: "@office", kind: "request", title: `Reply on request, ${String(rec.type || "Request").slice(0, 60)}`, body: "The requester replied. Open Team Requests.", link: "teamrequests", by: user.name || email });
      }
    } catch (e) { /* the notice is best effort */ }
    return reply(env, 200, { ok: true });
  }
  if (request.method === "DELETE" && m[1]) {
    const row = await env.DB.prepare("SELECT by_email FROM app_requests WHERE id = ?").bind(m[1]).first();
    if (!row) return reply(env, 200, { ok: true });
    if (row.by_email !== email && !isApprover(user)) return reply(env, 403, { error: "forbidden" });
    await env.DB.prepare("DELETE FROM app_requests WHERE id = ?").bind(m[1]).run();
    return reply(env, 200, { ok: true });
  }
  return reply(env, 405, { error: "method not allowed" });
}
// ---- Partners added in the app (online portals the office logs into). The seed list stays in the app, only partners added here live in D1. T1 and T2 only. Name, username and password are required.
let partnersReady = false;
async function partnersTable(env) {
  if (partnersReady) return;
  partnersReady = true;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS partners (id TEXT PRIMARY KEY, kind TEXT, name TEXT NOT NULL, phone TEXT, website TEXT, address TEXT, account TEXT, username TEXT, password TEXT, email TEXT, notes TEXT, by_name TEXT, created_at TEXT, updated_at TEXT)").run();
  try { await env.DB.prepare("ALTER TABLE partners ADD COLUMN removed INTEGER NOT NULL DEFAULT 0").run(); } catch (e) { /* already there */ }
}
async function partnersRoutes(request, env, user, url) {
  const m = url.pathname.match(/^\/api\/partners(?:\/([A-Za-z0-9_-]{1,64}))?$/);
  if (!m) return null;
  if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
  await partnersTable(env);
  if (request.method === "GET" && !m[1]) {
    const rows = ((await env.DB.prepare("SELECT id, kind, name, phone, website, address, account, username, password, email, notes FROM partners WHERE removed = 0 ORDER BY name COLLATE NOCASE").all()).results) || [];
    const gone = ((await env.DB.prepare("SELECT id FROM partners WHERE removed = 1").all()).results) || [];
    return reply(env, 200, { partners: rows.map((r) => ({ ...r, kind: r.kind || void 0 })), removed: gone.map((r) => r.id), worker: WORKER_BUILD });
  }
  if (request.method === "PUT" && m[1]) {
    let b;
    try { b = await request.json(); } catch (e) { b = null; }
    if (!b || typeof b !== "object") return reply(env, 400, { error: "bad json" });
    const f = (k, n) => String(b[k] == null ? "" : b[k]).trim().slice(0, n || 200);
    const name = f("name"), username = f("username"), password = String(b.password == null ? "" : b.password).slice(0, 500);
    // Partners added in the app (ids starting with u) must have a username and password. The built in list can be renamed or changed without them.
    const added = /^u/.test(m[1]);
    if (!name || (added && (!username || !password))) return reply(env, 400, { error: "A partner needs a name, a username and a password" });
    const now = new Date().toISOString();
    const kind = f("kind") === "supplier" ? "supplier" : "";
    await env.DB.prepare("INSERT INTO partners (id, kind, name, phone, website, address, account, username, password, email, notes, by_name, created_at, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET removed = 0, kind = excluded.kind, name = excluded.name, phone = excluded.phone, website = excluded.website, address = excluded.address, account = excluded.account, username = excluded.username, password = excluded.password, email = excluded.email, notes = excluded.notes, by_name = excluded.by_name, updated_at = excluded.updated_at")
      .bind(m[1], kind, name, f("phone", 40), f("website", 300), f("address", 300), f("account", 100), username, password, f("email", 200), String(b.notes == null ? "" : b.notes).slice(0, 2000), user.name || "", now, now).run();
    return reply(env, 200, { ok: true });
  }
  if (request.method === "DELETE" && m[1]) {
    // Removes a partner for everyone. The built in list lives in the app, so a removed one is kept as a marker the app hides.
    const now = new Date().toISOString();
    await env.DB.prepare("INSERT INTO partners (id, name, removed, by_name, created_at, updated_at) VALUES (?, '(removed)', 1, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET removed = 1, by_name = excluded.by_name, updated_at = excluded.updated_at").bind(m[1], user.name || "", now, now).run();
    return reply(env, 200, { ok: true });
  }
  return reply(env, 405, { error: "method not allowed" });
}
// The Action List. Workspaces (a name and who is in it) and action items live here so they survive reloads and everyone sees the same list. Staff see items assigned to them, items they created, and items in workspaces they belong to. The DFP Team workspace shows staff (T3) only their own items. Only T1 and T2 can edit an item; the person it is assigned to, or who created it, can tick it done.
let todosReady = false;
async function todosTable(env) {
  if (todosReady) return;
  todosReady = true;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS todo_workspaces (name TEXT PRIMARY KEY, members TEXT NOT NULL DEFAULT '[]', by_name TEXT, created_at TEXT, updated_at TEXT)").run();
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS todo_items (id TEXT PRIMARY KEY, title TEXT NOT NULL, note TEXT, assigned_to TEXT, ws TEXT, due TEXT, repeat TEXT, done INTEGER NOT NULL DEFAULT 0, by_name TEXT, created TEXT, updated_at TEXT)").run();
  try { await env.DB.prepare("ALTER TABLE todo_items ADD COLUMN ord INTEGER NOT NULL DEFAULT 0").run(); } catch (e) {}
}
async function todosRoutes(request, env, user, url) {
  const m = url.pathname.match(/^\/api\/todos(?:\/(items|workspaces)\/([^/]{1,80}))?$/);
  if (!m) return null;
  await todosTable(env);
  const me = String(user.name || "").trim().split(/\s+/)[0];
  const senior = user.tier === "T1" || user.tier === "T2";
  const firstOf = (n) => String(n || "").trim().split(/\s+/)[0];
  if (request.method === "GET" && !m[1]) {
    await env.DB.prepare("DELETE FROM todo_items WHERE done = 1 AND updated_at IS NOT NULL AND updated_at < ?").bind(new Date(Date.now() - 92 * 86400000).toISOString()).run();
    const ws = ((await env.DB.prepare("SELECT name, members FROM todo_workspaces ORDER BY name COLLATE NOCASE").all()).results) || [];
    const workspaces = {};
    for (const w of ws) { let a = []; try { a = JSON.parse(w.members || "[]"); } catch (e) { a = []; } workspaces[w.name] = a; }
    const rows = ((await env.DB.prepare("SELECT id, title, note, assigned_to, ws, due, repeat, done, ord, by_name, created FROM todo_items ORDER BY created DESC, rowid DESC").all()).results) || [];
    const items = rows.filter((r) => {
      if (senior) return true;
      const mine = firstOf(r.assigned_to) === me || r.by_name === me;
      if (mine) return true;
      if (!r.ws || !(workspaces[r.ws] || []).includes(me)) return false;
      return r.ws === "DFP Team" ? senior : true;
    }).map((r) => { let rp = null; try { rp = r.repeat ? JSON.parse(r.repeat) : null; } catch (e) { rp = null; } return { id: r.id, title: r.title, note: r.note || "", assignedTo: r.assigned_to || "", group: r.ws || "", due: r.due || "", importance: "date", repeat: rp, done: r.done === 1, ord: Number(r.ord) || 0, by: r.by_name || "", created: r.created || "" }; });
    return reply(env, 200, { workspaces, items, worker: WORKER_BUILD });
  }
  if (request.method === "PUT" && m[1] === "workspaces") {
    const name = decodeURIComponent(m[2]).trim().slice(0, 50);
    let b; try { b = await request.json(); } catch (e) { b = null; }
    if (b && typeof b.rename === "string") {
      if (!senior) return reply(env, 403, { error: "forbidden" });
      const nu = b.rename.trim().slice(0, 50);
      if (!name || !nu) return reply(env, 400, { error: "bad request" });
      const clash = await env.DB.prepare("SELECT name FROM todo_workspaces WHERE lower(name) = lower(?) AND name <> ?").bind(nu, name).first();
      if (clash) return reply(env, 409, { error: "A workspace with that name already exists" });
      const now = new Date().toISOString();
      await env.DB.prepare("UPDATE todo_workspaces SET name = ?, updated_at = ? WHERE name = ?").bind(nu, now, name).run();
      await env.DB.prepare("UPDATE todo_items SET ws = ?, updated_at = ? WHERE ws = ?").bind(nu, now, name).run();
      return reply(env, 200, { ok: true });
    }
    if (!name || !b || !Array.isArray(b.members)) return reply(env, 400, { error: "bad request" });
    const members = [...new Set(b.members.map((x) => String(x || "").trim().slice(0, 40)).filter(Boolean))].slice(0, 60);
    const cur = await env.DB.prepare("SELECT members FROM todo_workspaces WHERE name = ?").bind(name).first();
    let curMembers = []; try { curMembers = cur ? JSON.parse(cur.members || "[]") : []; } catch (e) { curMembers = []; }
    if (!senior && !(cur && curMembers.includes(me))) return reply(env, 403, { error: "forbidden" });
    const now = new Date().toISOString();
    await env.DB.prepare("INSERT INTO todo_workspaces (name, members, by_name, created_at, updated_at) VALUES (?, ?, ?, ?, ?) ON CONFLICT(name) DO UPDATE SET members = excluded.members, updated_at = excluded.updated_at").bind(name, JSON.stringify(members), me, now, now).run();
    return reply(env, 200, { ok: true });
  }
  if (request.method === "DELETE" && m[1] === "workspaces") {
    if (!senior) return reply(env, 403, { error: "forbidden" });
    const name = decodeURIComponent(m[2]).trim().slice(0, 50);
    const openCount = await env.DB.prepare("SELECT COUNT(*) AS n FROM todo_items WHERE ws = ? AND done = 0").bind(name).first();
    if (openCount && openCount.n > 0) return reply(env, 409, { error: "The workspace still has open action items" });
    await env.DB.prepare("DELETE FROM todo_items WHERE ws = ? AND done = 1").bind(name).run();
    await env.DB.prepare("DELETE FROM todo_workspaces WHERE name = ?").bind(name).run();
    return reply(env, 200, { ok: true });
  }
  if (request.method === "DELETE" && m[1] === "items") {
    if (!senior) return reply(env, 403, { error: "forbidden" });
    const id = m[2];
    if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) return reply(env, 400, { error: "bad id" });
    const cur = await env.DB.prepare("SELECT id, title, assigned_to, ws FROM todo_items WHERE id = ?").bind(id).first();
    if (!cur) return reply(env, 200, { ok: true });
    if (url.searchParams.get("series") === "1") await env.DB.prepare("DELETE FROM todo_items WHERE done = 1 AND title = ? AND COALESCE(assigned_to, '') = ? AND COALESCE(ws, '') = ?").bind(cur.title, cur.assigned_to || "", cur.ws || "").run();
    await env.DB.prepare("DELETE FROM todo_items WHERE id = ?").bind(id).run();
    return reply(env, 200, { ok: true });
  }
  if (request.method === "PUT" && m[1] === "items") {
    const id = m[2];
    if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) return reply(env, 400, { error: "bad id" });
    let b; try { b = await request.json(); } catch (e) { b = null; }
    if (!b || typeof b !== "object") return reply(env, 400, { error: "bad json" });
    const cur = await env.DB.prepare("SELECT id, assigned_to, by_name FROM todo_items WHERE id = ?").bind(id).first();
    const now = new Date().toISOString();
    if (cur) {
      const mineItem = firstOf(cur.assigned_to) === me || cur.by_name === me;
      if (senior && b.title != null) {
        const rp = b.repeat && b.repeat.type ? JSON.stringify({ type: String(b.repeat.type).slice(0, 10), dow: Number(b.repeat.dow) || 0, dom: Number(b.repeat.dom) || 1 }) : null;
        await env.DB.prepare("UPDATE todo_items SET title = ?, note = ?, assigned_to = ?, ws = ?, due = ?, repeat = ?, done = ?, ord = ?, updated_at = ? WHERE id = ?").bind(String(b.title || "").trim().slice(0, 60), String(b.note || "").slice(0, 2000), String(b.assignedTo || "").slice(0, 80), String(b.group || "").slice(0, 50), String(b.due || "").slice(0, 10), rp, b.done ? 1 : 0, Math.max(0, Math.min(999, Math.round(Number(b.ord) || 0))), now, id).run();
        return reply(env, 200, { ok: true });
      }
      if (mineItem || senior) { await env.DB.prepare("UPDATE todo_items SET done = ?, updated_at = ? WHERE id = ?").bind(b.done ? 1 : 0, now, id).run(); return reply(env, 200, { ok: true }); }
      return reply(env, 403, { error: "forbidden" });
    }
    const title = String(b.title || "").trim().slice(0, 60);
    if (!title) return reply(env, 400, { error: "A title is needed" });
    const rp = b.repeat && b.repeat.type ? JSON.stringify({ type: String(b.repeat.type).slice(0, 10), dow: Number(b.repeat.dow) || 0, dom: Number(b.repeat.dom) || 1 }) : null;
    await env.DB.prepare("INSERT INTO todo_items (id, title, note, assigned_to, ws, due, repeat, done, ord, by_name, created, updated_at) VALUES (?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?)").bind(id, title, String(b.note || "").slice(0, 2000), String(b.assignedTo || "").slice(0, 80), String(b.group || "").slice(0, 50), String(b.due || "").slice(0, 10), rp, Math.max(0, Math.min(999, Math.round(Number(b.ord) || 0))), me, String(b.created || now.slice(0, 10)).slice(0, 10), now).run();
    return reply(env, 200, { ok: true });
  }
  return reply(env, 405, { error: "method not allowed" });
}
async function noticesTable(env) {
  if (noticesReady) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS notices (id TEXT PRIMARY KEY, to_email TEXT NOT NULL, kind TEXT, title TEXT, body TEXT, todo INTEGER NOT NULL DEFAULT 0, link TEXT, data TEXT, by_name TEXT, created_at TEXT DEFAULT (datetime('now')), read_at TEXT, done_at TEXT, done_by TEXT)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS notices_to ON notices (to_email, created_at)").run();
  noticesReady = true;
  // One time: every notice is cleared so everyone starts fresh (w144).
  if (!(await migDone(env, "notices_fresh_start_w144"))) {
    try { await env.DB.prepare("DELETE FROM notices").run(); await migMark(env, "notices_fresh_start_w144", "all notices cleared"); } catch (e) { noticesReady = false; throw e; }
  }
  // One time: the timesheet notices that were kept on their own are brought across.
  if (!(await migDone(env, "ts_notices_to_notices_v1"))) {
    try {
      const old = ((await env.DB.prepare("SELECT id, email, text, created_at, seen FROM ts_notices").all()).results) || [];
      for (const n of old) await env.DB.prepare("INSERT OR IGNORE INTO notices (id, to_email, kind, title, body, todo, created_at, read_at) VALUES (?, ?, 'timesheet', 'Timesheet', ?, 0, ?, ?)").bind(`ts${n.id}`, String(n.email || "").toLowerCase(), n.text || "", n.created_at || null, n.seen === 1 ? (n.created_at || "seen") : null).run();
      await migMark(env, "ts_notices_to_notices_v1", `${old.length} timesheet notices copied`);
    } catch (e) { /* no old table yet */ await migMark(env, "ts_notices_to_notices_v1", "nothing to copy"); }
  }
}
// Push notifications. Which kinds of notice also go to people's phones is decided here, once. Anything marked todo (someone has to act) also pushes. To change what pushes, change this list.
const PUSH_KINDS = ["timesheet", "payroll", "cardreset", "card", "request", "policy", "dailyops"];
const b64uEnc = (buf) => { const u = new Uint8Array(buf); let s = ""; for (let i = 0; i < u.length; i++) s += String.fromCharCode(u[i]); return btoa(s).replace(/\+/g, "-").replace(/\//g, "_").replace(/=+$/, ""); };
const b64uDec = (str) => { const s = String(str || "").replace(/-/g, "+").replace(/_/g, "/"); const bin = atob(s + "=".repeat((4 - (s.length % 4)) % 4)); const u = new Uint8Array(bin.length); for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i); return u; };
const concatBytes = (...parts) => { const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0)); let o = 0; for (const p of parts) { out.set(p, o); o += p.length; } return out; };
let pushReady = false;
async function pushTable(env) {
  if (pushReady) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS push_subs (endpoint TEXT PRIMARY KEY, email TEXT NOT NULL, tier TEXT, p256dh TEXT NOT NULL, auth TEXT NOT NULL, ua TEXT, created_at TEXT DEFAULT (datetime('now')), last_ok TEXT)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS push_subs_email ON push_subs (email)").run();
  try { await env.DB.prepare("ALTER TABLE push_subs ADD COLUMN all_hours INTEGER NOT NULL DEFAULT 0").run(); } catch (e) { /* already there */ }
  try { await env.DB.prepare("ALTER TABLE push_subs ADD COLUMN hours_mode TEXT").run(); } catch (e) { /* already there */ }
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS push_held (id INTEGER PRIMARY KEY AUTOINCREMENT, endpoint TEXT NOT NULL, created_at TEXT DEFAULT (datetime('now')))").run();
  pushReady = true;
}
// Each phone picks when push may reach it. Weekday windows are Monday to Friday in Melbourne time, or 24 hours every day. Anything held back is counted and sent as one push when that phone's window next opens. The notice itself is always in the shield.
const PUSH_MODES = { "0730-1600": [450, 960], "0700-1630": [420, 990], "0600-1800": [360, 1080], "24": null };
const pushModeOf = (row) => (row && row.hours_mode && Object.prototype.hasOwnProperty.call(PUSH_MODES, row.hours_mode)) ? row.hours_mode : (row && row.all_hours ? "24" : "0700-1630");
function pushInHours(mode, d) {
  const w = PUSH_MODES[mode || "0700-1630"];
  if (w === null) return true;
  const parts = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", weekday: "short", hour: "2-digit", minute: "2-digit", hour12: false }).formatToParts(d || new Date());
  const get = (t) => (parts.find((x) => x.type === t) || {}).value || "";
  const mins = (Number(get("hour")) % 24) * 60 + Number(get("minute"));
  return !["Sat", "Sun"].includes(get("weekday")) && mins >= w[0] && mins < w[1];
}
async function pushFlushHeld(env) {
  try {
    await pushTable(env);
    const rows = ((await env.DB.prepare("SELECT endpoint, COUNT(*) AS n FROM push_held GROUP BY endpoint").all()).results) || [];
    for (const r of rows) {
      const sub = await env.DB.prepare("SELECT * FROM push_subs WHERE endpoint = ?").bind(r.endpoint).first();
      if (sub && !pushInHours(pushModeOf(sub))) continue;
      if (sub) await pushSendOne(env, sub, { title: "Reservoir", body: r.n === 1 ? "You have 1 notification waiting in Reservoir" : `You have ${r.n} notifications waiting in Reservoir`, tag: "held", open: "notices" });
      await env.DB.prepare("DELETE FROM push_held WHERE endpoint = ?").bind(r.endpoint).run();
    }
  } catch (e) { console.log("push flush failed", String(e && e.message || e).slice(0, 160)); }
}
// The signing keys for push, made once and kept in app_settings. The public half goes to the app, the private half never leaves the worker.
async function pushKeys(env) {
  await settingsTable(env);
  const read = async () => { const r = await env.DB.prepare("SELECT value FROM app_settings WHERE key = 'secret.push_vapid'").first(); try { return r && r.value ? JSON.parse(r.value) : null; } catch (e) { return null; } };
  let k = await read();
  if (!k) {
    const pair = await crypto.subtle.generateKey({ name: "ECDSA", namedCurve: "P-256" }, true, ["sign", "verify"]);
    const pub = b64uEnc(await crypto.subtle.exportKey("raw", pair.publicKey));
    const priv = await crypto.subtle.exportKey("jwk", pair.privateKey);
    await env.DB.prepare("INSERT OR IGNORE INTO app_settings (key, value, detail, by_name) VALUES ('secret.push_vapid', ?, 'Push notification signing keys, made by the worker', 'worker')").bind(JSON.stringify({ pub, priv })).run();
    k = await read();
  }
  return k;
}
async function vapidHeader(env, endpoint) {
  const k = await pushKeys(env);
  const key = await crypto.subtle.importKey("jwk", k.priv, { name: "ECDSA", namedCurve: "P-256" }, false, ["sign"]);
  const enc = (o) => b64uEnc(new TextEncoder().encode(JSON.stringify(o)));
  const head = enc({ typ: "JWT", alg: "ES256" });
  const body = enc({ aud: new URL(endpoint).origin, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: "mailto:maintenance@directflow.com.au" });
  const sig = await crypto.subtle.sign({ name: "ECDSA", hash: "SHA-256" }, key, new TextEncoder().encode(`${head}.${body}`));
  return `vapid t=${head}.${body}.${b64uEnc(sig)}, k=${k.pub}`;
}
// Message encryption for web push (RFC 8291, aes128gcm). One record, no padding.
async function pushEncrypt(p256dh, authSecret, text) {
  const uaPub = b64uDec(p256dh), auth = b64uDec(authSecret);
  const eph = await crypto.subtle.generateKey({ name: "ECDH", namedCurve: "P-256" }, true, ["deriveBits"]);
  const asPub = new Uint8Array(await crypto.subtle.exportKey("raw", eph.publicKey));
  const uaKey = await crypto.subtle.importKey("raw", uaPub, { name: "ECDH", namedCurve: "P-256" }, false, []);
  const shared = await crypto.subtle.deriveBits({ name: "ECDH", public: uaKey }, eph.privateKey, 256);
  const hkdf = async (ikm, salt, info, bytes) => crypto.subtle.deriveBits({ name: "HKDF", hash: "SHA-256", salt, info }, await crypto.subtle.importKey("raw", ikm, "HKDF", false, ["deriveBits"]), bytes * 8);
  const ikm = new Uint8Array(await hkdf(new Uint8Array(shared), auth, concatBytes(new TextEncoder().encode("WebPush: info\0"), uaPub, asPub), 32));
  const salt = crypto.getRandomValues(new Uint8Array(16));
  const cek = await hkdf(ikm, salt, new TextEncoder().encode("Content-Encoding: aes128gcm\0"), 16);
  const nonce = await hkdf(ikm, salt, new TextEncoder().encode("Content-Encoding: nonce\0"), 12);
  const aes = await crypto.subtle.importKey("raw", cek, "AES-GCM", false, ["encrypt"]);
  const plain = concatBytes(new TextEncoder().encode(text), new Uint8Array([2]));
  const cipher = new Uint8Array(await crypto.subtle.encrypt({ name: "AES-GCM", iv: nonce }, aes, plain));
  return concatBytes(salt, new Uint8Array([0, 0, 16, 0]), new Uint8Array([asPub.length]), asPub, cipher);
}
async function pushSendOne(env, sub, payload) {
  try {
    const body = await pushEncrypt(sub.p256dh, sub.auth, JSON.stringify(payload));
    const res = await fetch(sub.endpoint, { method: "POST", headers: { Authorization: await vapidHeader(env, sub.endpoint), "Content-Encoding": "aes128gcm", "Content-Type": "application/octet-stream", TTL: "86400", Urgency: "normal" }, body, signal: AbortSignal.timeout(8000) });
    if (res.status === 404 || res.status === 410) { await env.DB.prepare("DELETE FROM push_subs WHERE endpoint = ?").bind(sub.endpoint).run(); return "gone"; }
    if (res.ok) { await env.DB.prepare("UPDATE push_subs SET last_ok = datetime('now') WHERE endpoint = ?").bind(sub.endpoint).run(); return "sent"; }
    console.log("push refused", res.status, String(sub.endpoint).slice(0, 40));
    return "refused";
  } catch (e) { console.log("push failed", String(e && e.message || e).slice(0, 160)); return "failed"; }
}
// Sends one notice to the phones of whoever it was addressed to. Notices for the office go to everyone signed up at the senior levels (T1 and T2).
async function pushForNotice(env, n, id) {
  try {
    if (!(PUSH_KINDS.includes(String(n.kind || "")) || n.todo)) return;
    await pushTable(env);
    const to = String(n.to || "").toLowerCase();
    const rows = to === "@office"
      ? ((await env.DB.prepare("SELECT * FROM push_subs WHERE tier IN ('T1', 'T2')").all()).results) || []
      : ((await env.DB.prepare("SELECT * FROM push_subs WHERE email = ?").bind(to).all()).results) || [];
    if (!rows.length) return;
    const payload = { title: String(n.title || "Reservoir").slice(0, 120), body: n.kind === "request" ? String(n.body || "Open Reservoir to read it").slice(0, 160) : "Open Reservoir to read it", tag: `n-${id || Date.now()}`, open: /^[a-z]+$/.test(String(n.link || "")) ? String(n.link) : "notices" };
    await Promise.allSettled(rows.map(async (r) => {
      if (pushInHours(pushModeOf(r))) return pushSendOne(env, r, payload);
      await env.DB.prepare("INSERT INTO push_held (endpoint) VALUES (?)").bind(r.endpoint).run();
    }));
  } catch (e) { console.log("push for notice failed", String(e && e.message || e).slice(0, 160)); }
}
async function pushRoutes(request, env, user, url) {
  if (!url.pathname.startsWith("/api/push/")) return null;
  const email = String(user.email || "").toLowerCase();
  if (!email) return reply(env, 400, { error: "No email on your sign in" });
  await pushTable(env);
  if (url.pathname === "/api/push/key" && request.method === "GET") return reply(env, 200, { key: (await pushKeys(env)).pub });
  if (url.pathname === "/api/push/status" && request.method === "GET") {
    const ep = url.searchParams.get("endpoint") || "";
    const row = ep ? await env.DB.prepare("SELECT email, all_hours, hours_mode FROM push_subs WHERE endpoint = ? AND email = ?").bind(ep, email).first() : null;
    const n = await env.DB.prepare("SELECT COUNT(*) AS n FROM push_subs WHERE email = ?").bind(email).first();
    return reply(env, 200, { subscribed: !!row, mode: row ? pushModeOf(row) : "", devices: n ? n.n : 0 });
  }
  if (request.method !== "POST") return reply(env, 405, { error: "method not allowed" });
  if (user.viewAs) return reply(env, 403, { error: "Viewing as someone else is read only." });
  let body = {};
  try { body = await request.json(); } catch (e) { body = {}; }
  if (url.pathname === "/api/push/subscribe") {
    const s = body.subscription || {};
    const ep = String(s.endpoint || ""), p = String(s.keys && s.keys.p256dh || ""), a = String(s.keys && s.keys.auth || "");
    if (!/^https:\/\//.test(ep) || ep.length > 600 || !p || !a) return reply(env, 400, { error: "That subscription is not valid" });
    const subMode = Object.prototype.hasOwnProperty.call(PUSH_MODES, String(body.mode)) ? String(body.mode) : null;
    await env.DB.prepare("INSERT INTO push_subs (endpoint, email, tier, p256dh, auth, ua, hours_mode, all_hours) VALUES (?, ?, ?, ?, ?, ?, COALESCE(?, '0700-1630'), ?) ON CONFLICT(endpoint) DO UPDATE SET email = excluded.email, tier = excluded.tier, p256dh = excluded.p256dh, auth = excluded.auth, ua = excluded.ua, hours_mode = COALESCE(?, hours_mode), all_hours = CASE WHEN ? IS NULL THEN all_hours WHEN ? = '24' THEN 1 ELSE 0 END").bind(ep, email, String(user.tier || "T3"), p, a, String(body.ua || "").slice(0, 160), subMode, subMode === "24" ? 1 : 0, subMode, subMode, subMode).run();
    return reply(env, 200, { ok: true });
  }
  if (url.pathname === "/api/push/hours") {
    const mode = Object.prototype.hasOwnProperty.call(PUSH_MODES, String(body.mode)) ? String(body.mode) : "0700-1630";
    await env.DB.prepare("UPDATE push_subs SET hours_mode = ?, all_hours = ? WHERE endpoint = ? AND email = ?").bind(mode, mode === "24" ? 1 : 0, String(body.endpoint || ""), email).run();
    return reply(env, 200, { ok: true });
  }
  if (url.pathname === "/api/push/unsubscribe") {
    await env.DB.prepare("DELETE FROM push_subs WHERE endpoint = ? AND email = ?").bind(String(body.endpoint || ""), email).run();
    return reply(env, 200, { ok: true });
  }
  if (url.pathname === "/api/push/test") {
    const rows = ((await env.DB.prepare("SELECT * FROM push_subs WHERE email = ?").bind(email).all()).results) || [];
    const out = await Promise.all(rows.map((r) => pushSendOne(env, r, { title: "Reservoir test", body: "Push notifications are working", tag: `t-${Date.now()}`, open: "notices" })));
    return reply(env, 200, { sent: out.filter((x) => x === "sent").length, devices: rows.length });
  }
  return reply(env, 404, { error: "not found" });
}
async function notify(env, n) {
  try {
    await noticesTable(env);
    const id = crypto.randomUUID();
    await env.DB.prepare("INSERT INTO notices (id, to_email, kind, title, body, todo, link, data, by_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)").bind(id, String(n.to || "").toLowerCase(), String(n.kind || "info").slice(0, 30), String(n.title || "").slice(0, 160), String(n.body || "").slice(0, 600), n.todo ? 1 : 0, n.link ? String(n.link).slice(0, 40) : null, n.data ? JSON.stringify(n.data).slice(0, 1500) : null, String(n.by || "").slice(0, 120)).run();
    await pushForNotice(env, n, id);
    return id;
  } catch (e) { console.log("notify failed", String(e && e.message || e).slice(0, 200)); return null; }
}
const noticeOut = (r) => { let data = null; try { data = r.data ? JSON.parse(r.data) : null; } catch (e) { data = null; } return { id: r.id, kind: r.kind || "info", title: r.title || "", body: r.body || "", todo: r.todo === 1, link: r.link || "", data, by: r.by_name || "", createdAt: r.created_at || "", read: !!r.read_at, done: !!r.done_at, doneBy: r.done_by || "", office: r.to_email === "@office" }; };
// At 10am on a weekday, if any Daily Operations task is still not ticked, tell the directors (T1). One alert a day. The cron runs every 30 minutes, so a run between 10:00 and 12:00 sends it if it has not gone yet.
async function dailyOpsAlert(env) {
  const parts = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", weekday: "short", hour: "2-digit", minute: "2-digit", hour12: false }).formatToParts(new Date());
  const get = (t) => (parts.find((x) => x.type === t) || {}).value || "";
  const mins = (Number(get("hour")) % 24) * 60 + Number(get("minute"));
  if (["Sat", "Sun"].includes(get("weekday")) || mins < 600 || mins >= 720) return;
  await todosTable(env);
  await settingsTable(env);
  const today = melbourneDay();
  const sent = await env.DB.prepare("SELECT value FROM app_settings WHERE key = 'alert.dailyops'").first();
  if (sent && sent.value === today) return;
  const rows = ((await env.DB.prepare("SELECT done, due FROM todo_items WHERE ws = 'Daily Operations' AND due <> '' AND due <= ?").bind(today).all()).results) || [];
  const todays = rows.filter((r) => r.done !== 1 || r.due === today);
  const left = todays.filter((r) => r.done !== 1).length;
  if (!left) return;
  const tiers = await accessTiers(env);
  const to = Object.keys(tiers).filter((e) => tiers[e] === "T1");
  let ok = false;
  for (const e of to) { const id = await notify(env, { to: e, kind: "dailyops", title: `Daily Operations ${left}/${todays.length} Incomplete`, body: "Open the Action List to see what is left.", link: "todo", by: "Reservoir" }); if (id) ok = true; }
  if (ok) await env.DB.prepare("INSERT INTO app_settings (key, value, detail, by_name, at) VALUES ('alert.dailyops', ?, 'alert sent', 'Daily Operations alert', datetime('now')) ON CONFLICT(key) DO UPDATE SET value = excluded.value, detail = excluded.detail, at = excluded.at").bind(today).run();
}
async function noticesRoutes(request, env, user, url) {
  if (!url.pathname.startsWith("/api/notices")) return null;
  const email = String(user.email || "").toLowerCase();
  const office = user.tier === "T1" || user.tier === "T2";
  if (!email) return reply(env, 400, { error: "No email on your sign in" });
  await noticesTable(env);
  const mine = office ? "(to_email = ? OR to_email = '@office')" : "to_email = ?";
  if (url.pathname === "/api/notices" && request.method === "GET") {
    const rows = ((await env.DB.prepare(`SELECT * FROM notices WHERE ${mine} ORDER BY created_at DESC, rowid DESC LIMIT 80`).bind(email).all()).results) || [];
    return reply(env, 200, { notices: rows.map(noticeOut), worker: WORKER_BUILD });
  }
  if (url.pathname === "/api/notices/read-all" && request.method === "POST") {
    await env.DB.prepare("UPDATE notices SET read_at = datetime('now') WHERE to_email = ? AND read_at IS NULL AND todo = 0").bind(email).run();
    return reply(env, 200, { ok: true });
  }
  const m = url.pathname.match(/^\/api\/notices\/([0-9a-zA-Z-]{1,40})\/(read|done)$/);
  if (m && request.method === "POST") {
    const row = await env.DB.prepare(`SELECT * FROM notices WHERE id = ? AND ${mine}`).bind(m[1], email).first();
    if (!row) return reply(env, 404, { error: "That notice is not there any more." });
    if (m[2] === "read") await env.DB.prepare("UPDATE notices SET read_at = COALESCE(read_at, datetime('now')) WHERE id = ?").bind(row.id).run();
    else await env.DB.prepare("UPDATE notices SET done_at = COALESCE(done_at, datetime('now')), read_at = COALESCE(read_at, datetime('now')), done_by = COALESCE(done_by, ?) WHERE id = ?").bind(user.name || email, row.id).run();
    return reply(env, 200, { ok: true });
  }
  return reply(env, 404, { error: "not found" });
}
// For now every approval, of any kind, is Tayler's alone. Andrew and Mitch can still see things; they cannot decide them. This is the one place to change when approvals are shared out.
const isApprover = (user) => user.tier === "T1" && /^tayler\b/i.test(String(user.name || "").trim());
const tsCanApprove = (user) => isApprover(user);
const tsRowOut = (r) => ({ id: r.id, name: r.name || r.email, week: r.week_start, weekEnd: r.week_end, total: Number(r.total) || 0, days: (() => { try { return JSON.parse(r.days || "[]"); } catch (e) { return []; } })(), createdAt: r.created_at || "", error: r.last_error || "" });

// The director's weekly view: one card per person for the week being paid, then a history of earlier pay periods.
async function tsOfficeRoute(env, user) {
  await tsTables(env);
  const today = tsToday();
  const t = new Date(`${today}T00:00:00Z`);
  const thisMonday = tsAddDays(today, -((t.getUTCDay() + 6) % 7));
  const lastMonday = tsAddDays(thisMonday, -7);
  let filedEnd = "";
  try {
    await dfpTables(env);
    const r = await env.DB.prepare("SELECT MAX(period_end) AS e FROM payslip_runs WHERE filed = 1").first();
    filedEnd = (r && r.e) ? String(r.e).slice(0, 10) : "";
  } catch (e) { filedEnd = ""; }
  const staff = (((await env.DB.prepare("SELECT employee_id, email, name, employment_basis FROM xero_payroll WHERE email IS NOT NULL AND email <> '' AND (status IS NULL OR status = 'ACTIVE')").all()).results) || [])
    .filter((p) => !/andrew\s+wowk/i.test(String(p.name || "")))
    .map((p) => ({ email: String(p.email).toLowerCase(), name: String(p.name || p.email), casual: /casual/i.test(String(p.employment_basis || "")) }))
    .sort((a, b) => a.name.localeCompare(b.name));
  const subs = ((await env.DB.prepare("SELECT * FROM timesheet_submissions WHERE status IN ('pending','sent') AND week_start >= ? ORDER BY created_at DESC").bind(tsAddDays(today, -200)).all()).results) || [];
  // Timesheets the team put in through Xero Me, as last pulled from Xero. Shown for anyone with nothing submitted in Reservoir that week, and only for weeks not yet filed.
  const xeroBy = new Map();
  try {
    await timesheetTable(env);
    const xs = ((await env.DB.prepare("SELECT id, email, start_date, status, total_hours, lines FROM xero_timesheets WHERE start_date >= ? AND (status IS NULL OR upper(status) NOT IN ('REJECTED','DELETED'))").bind(tsAddDays(today, -200)).all()).results) || [];
    for (const x of xs) { if (x.email) xeroBy.set(`${String(x.email).toLowerCase()}|${x.start_date}`, x); }
  } catch (e) { /* the Xero copy is not there yet, the screen shows Reservoir submissions only */ }
  const xeroCard = (x, who) => {
    const st = String(x.status || "").toUpperCase();
    const sums = [0, 0, 0, 0, 0, 0, 0];
    try { JSON.parse(x.lines || "[]").forEach((l) => (l.units || []).forEach((u, i) => { if (i < 7) sums[i] += Number(u) || 0; })); } catch (e) { /* no day lines */ }
    const done = st === "APPROVED" || st === "PROCESSED";
    return { id: `xero:${x.id}`, name: who.name, email: who.email, state: done ? "approved" : "xerodraft", total: Number(x.total_hours) || 0, days: sums.map((h) => ({ ord: h, ot15: 0, ot2: 0 })), auto: done ? "Approved in Xero" : st === "REQUESTED" ? "Submitted in Xero, not approved yet" : "Draft in Xero", error: "" };
  };
  const latest = new Map();
  for (const r of subs) { const k = `${r.email}|${r.week_start}`; if (!latest.has(k)) latest.set(k, r); }
  const byWeek = new Map();
  for (const r of latest.values()) { if (!byWeek.has(r.week_start)) byWeek.set(r.week_start, []); byWeek.get(r.week_start).push(r); }
  const filed = (week) => !!filedEnd && tsAddDays(week, 6) <= filedEnd;
  const card = (r, who) => {
    if (!r) return { id: "", name: who.name, email: who.email, state: "notsubmitted", total: 0, days: [], auto: "", error: "" };
    const out = tsRowOut(r);
    const state = r.status === "pending" ? "pending" : (/^(APPROVED|SALARIED)$/.test(String(r.xero_status || "").toUpperCase()) ? "approved" : "xerodraft");
    const auto = r.decided_by === "auto" ? "Automatic approval" : r.decided_by === "auto trade school" ? "Five days trade school, automatic approval" : "";
    return { id: out.id, name: out.name, email: r.email, state, total: out.total, days: out.days, auto, error: out.error };
  };
  const build = (week) => {
    const rows = byWeek.get(week) || [];
    const used = new Set();
    // Casual staff are left off until they have a timesheet that week. Everyone else shows by their Xero name, so a name saved on an old submission never differs from the rest.
    const cards = [];
    for (const who of staff) {
      const r = rows.find((x) => x.email === who.email);
      if (r) used.add(r.id);
      const xr = !r && xeroBy.get(`${who.email}|${week}`);
      if (!r && !xr && who.casual) continue;
      cards.push({ ...(xr ? xeroCard(xr, who) : card(r, who)), name: who.name });
    }
    rows.filter((r) => !used.has(r.id)).forEach((r) => cards.push(card(r, { name: r.name || r.email, email: r.email })));
    return { week, weekEnd: tsAddDays(week, 6), payDate: tsAddDays(week, 10), cards };
  };
  const weeks = new Set(byWeek.keys());
  const xeroWeeks = new Set();
  for (const k of xeroBy.keys()) { const w = k.split("|")[1]; if (w <= thisMonday && !filed(w)) { weeks.add(w); xeroWeeks.add(w); } }
  weeks.add(lastMonday);
  const active = [], history = [];
  [...weeks].sort().reverse().forEach((w) => {
    if (w > thisMonday) return;
    if (filed(w)) return; // once a pay run is filed its weeks leave this screen
    if (w === thisMonday && !byWeek.has(w) && !xeroWeeks.has(w)) return;
    active.push(build(w));
  });
  return reply(env, 200, { active, history, filedEnd, worker: WORKER_BUILD });
}

// ---- Automatic timesheets for the two people on a fixed 40 hours (Tayler and Andrew). Runs from the scheduled trigger and from a manual run.
// Each weekday from 4:30pm the day is filled in as 7:30 to 4:00 with a 30 minute break, unless it is a public holiday or approved leave. From 6pm Friday the week is submitted and sent to Xero as approved.
// Anyone who has already done their week by hand, or whose week is already in Xero, is left alone. Keep this holiday list in step with VIC_HOLIDAYS_COMMON_2026 in the app.
const TS_AUTO_NAMES = [/^tayler\b/i, /^andrew\b/i];
const TS_HOLIDAYS = {
  "2025-09-26": "Friday before the AFL Grand Final", "2025-10-22": "Geelong Cup", "2025-12-25": "Christmas Day", "2025-12-26": "Boxing Day",
  "2026-01-01": "New Year's Day", "2026-01-26": "Australia Day", "2026-03-09": "Labour Day", "2026-04-03": "Good Friday", "2026-04-04": "Easter Saturday", "2026-04-05": "Easter Sunday", "2026-04-06": "Easter Monday",
  "2026-04-25": "ANZAC Day", "2026-06-08": "King's Birthday", "2026-09-25": "Grand Final Friday", "2026-10-21": "Geelong Cup", "2026-12-25": "Christmas Day", "2026-12-26": "Boxing Day", "2026-12-28": "Additional day (Boxing Day falls on a Saturday)"
};
const TS_AUTO_HOURS = 8;
function tsMelbourne(now) {
  const parts = new Intl.DateTimeFormat("en-AU", { timeZone: "Australia/Melbourne", year: "numeric", month: "2-digit", day: "2-digit", hour: "2-digit", minute: "2-digit", hour12: false, weekday: "short" }).formatToParts(now);
  const g = (t) => (parts.find((x) => x.type === t) || {}).value || "";
  const dow = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"].indexOf(g("weekday"));
  return { date: `${g("year")}-${g("month")}-${g("day")}`, dow, mins: (Number(g("hour")) % 24) * 60 + Number(g("minute")) };
}
// The week as the app keeps it, so the person sees the filled days when they open Timesheets.
const tsAppKey = (iso) => { const [y, m, d] = iso.split("-").map(Number); return `${y}-${m}-${d}`; };
async function tsAutoRun(env) {
  await tsTables(env);
  await timesheetDraftsTable(env);
  const now = tsMelbourne(new Date());
  const thisMonday = tsAddDays(now.date, -((now.dow + 6) % 7));
  // This week, and last week while it is still unpaid, so a Friday night run that did not go through is picked up again.
  const lastMonday = tsAddDays(thisMonday, -7);
  const weeks = tsAddDays(lastMonday, 10) >= now.date ? [lastMonday, thisMonday] : [thisMonday];
  const out = [];
  const people = (((await env.DB.prepare("SELECT employee_id, email, name FROM xero_payroll WHERE email IS NOT NULL AND email <> '' AND (status IS NULL OR status = 'ACTIVE')").all()).results) || []).filter((p) => TS_AUTO_NAMES.some((re) => re.test(String(p.name || ""))));
  for (const who of people) {
    const email = String(who.email).toLowerCase();
    for (const monday of weeks) {
    const sunday = tsAddDays(monday, 6);
    const note = (m) => out.push(`${who.name}${monday === thisMonday ? "" : ` (week of ${monday})`}: ${m}`);
    const already = await tsDuplicate(env, who.employee_id, email, monday);
    if (already) { note(`left alone. ${already}`); continue; }
    const key = tsAppKey(monday);
    const row = await env.DB.prepare("SELECT data FROM timesheet_drafts WHERE email = ? AND week_start = ?").bind(email, key).first();
    let period = null;
    try { period = row ? JSON.parse(row.data) : null; } catch (e) { period = null; }
    if (period && period.status === "submitted") { note("left alone, already submitted in Reservoir"); continue; }
    if (!period) period = { key, status: "draft", days: [] };
    while (period.days.length < 6) { const i = period.days.length; period.days.push({ type: "work", holiday: TS_HOLIDAYS[tsAddDays(monday, i)] || null, entries: [], breakMins: 0, shortDayReason: null, shortDayApprover: null }); }
    const leave = ((await env.DB.prepare("SELECT from_date, to_date, type FROM xero_leave WHERE email = ? AND status = 'approved' AND to_date >= ? AND from_date <= ?").bind(email, monday, sunday).all()).results) || [];
    let changed = !row;
    for (let i = 0; i < 5; i++) {
      const date = tsAddDays(monday, i);
      const closed = date < now.date || (date === now.date && now.mins >= 16 * 60 + 30);
      const day = period.days[i];
      if (!closed || day.type !== "work" || (day.entries || []).length || day.breakMins) continue;
      if (TS_HOLIDAYS[date]) { if (!day.holiday) { day.holiday = TS_HOLIDAYS[date]; changed = true; } continue; }
      const l = leave.find((x) => x.from_date <= date && x.to_date >= date);
      if (l) { day.type = l.type === "personal" ? "personal_leave" : "annual_leave"; changed = true; continue; }
      day.entries = [{ id: `auto-${date}`, start: 450, end: 960 }];
      day.breakMins = 30;
      changed = true;
    }
    if (changed) {
      await env.DB.prepare("INSERT INTO timesheet_drafts (email, week_start, data, updated_at) VALUES (?, ?, ?, ?) ON CONFLICT(email, week_start) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at").bind(email, key, JSON.stringify(period), new Date().toISOString()).run();
      note("days filled in");
    }
    const submitNow = monday < thisMonday || (now.dow === 5 && now.mins >= 18 * 60) || now.dow === 6 || now.dow === 0;
    if (!submitNow) continue;
    const hours = [0, 1, 2, 3, 4, 5, 6].map((i) => {
      const d = period.days[i];
      if (!d || d.type === "annual_leave" || d.type === "personal_leave") return 0;
      if (d.holiday) return TS_AUTO_HOURS;
      const e = (d.entries || [])[0];
      return e ? Math.max(0, Math.round(((e.end - e.start - (d.breakMins || 0)) / 60) * 100) / 100) : 0;
    });
    const days = hours.map((h) => ({ ord: h, ot15: 0, ot2: 0 }));
    if (tsTotal(days) <= 0) { note("nothing to submit, no working hours this week"); continue; }
    const sub = { id: crypto.randomUUID(), email, name: who.name || "", employee_id: who.employee_id, week_start: monday, week_end: sunday, days: JSON.stringify(days), total: tsTotal(days), created_at: new Date().toISOString() };
    const res = await tsPushXero(env, sub);
    if (res.fail) { note(`could not send to Xero. ${res.fail}`); continue; }
    await env.DB.prepare("INSERT INTO timesheet_submissions (id, email, name, employee_id, week_start, week_end, days, total, status, xero_status, forwarded, last_error, created_at, decided_at, decided_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'sent', ?, 0, NULL, ?, ?, 'auto')").bind(sub.id, sub.email, sub.name, sub.employee_id, sub.week_start, sub.week_end, sub.days, sub.total, res.xeroStatus, sub.created_at, sub.created_at).run();
    period.status = "submitted";
    period.submittedAt = sub.created_at;
    period.approval = "sent";
    period.xeroStatus = res.xeroStatus;
    await env.DB.prepare("INSERT INTO timesheet_drafts (email, week_start, data, updated_at) VALUES (?, ?, ?, ?) ON CONFLICT(email, week_start) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at").bind(email, key, JSON.stringify(period), new Date().toISOString()).run();
    note(`submitted ${sub.total} hours to Xero as ${res.xeroStatus}${res.note ? `. ${res.note}` : ""}`);
    }
  }
  ["Tayler", "Andrew"].forEach((n, i) => { if (!people.some((p) => TS_AUTO_NAMES[i].test(String(p.name || "")))) out.push(`${n}: not found in the Xero payroll list with an email and an active status, so nothing is run for him.`); });
  return out;
}
async function tsSubmitRoutes(request, env, user, url, parts) {
  const rawEnv = Object.prototype.hasOwnProperty.call(env, "__callerTier") ? Object.getPrototypeOf(env) : env;
  const email = String(user.email || "").trim().toLowerCase();
  if (parts[2] === "submit" && parts.length === 3 && request.method === "POST") {
    if (!email) return reply(env, 400, { error: "No email on your sign in" });
    const body = await request.json().catch(() => null) || {};
    const week = String(body.week || "");
    const bad = tsWeekProblem(week);
    if (bad) return reply(env, 400, { error: bad });
    const days = tsCleanDays(body.days);
    if (!days) return reply(env, 400, { error: "The hours on that timesheet are not valid." });
    // Casual staff are not paid public holidays, so those days go through as zero.
    try {
      const basis = await env.DB.prepare("SELECT employment_basis FROM xero_payroll WHERE lower(email) = ?").bind(email).first();
      if (basis && /casual/i.test(String(basis.employment_basis || ""))) days.forEach((d, i) => { if (TS_HOLIDAYS[tsAddDays(week, i)]) { d.ord = 0; d.ot15 = 0; d.ot2 = 0; } });
    } catch (e) { /* the app already shows those days as zero */ }
    if (tsTotal(days) <= 0) return reply(env, 400, { error: "There are no hours to submit." });
    await tsTables(env);
    const employeeId = await tsEmployee(env, email);
    if (!employeeId) return reply(env, 409, { error: "You are not in the Xero payroll list yet, so this cannot be sent. Tell Tayler." });
    const dup = await tsDuplicate(env, employeeId, email, week);
    if (dup) return reply(env, 409, { error: dup });
    const sub = { id: crypto.randomUUID(), email, name: user.name || "", employee_id: employeeId, week_start: week, week_end: tsAddDays(week, 6), days: JSON.stringify(days), total: tsTotal(days), created_at: new Date().toISOString() };
    const trade = user.tier !== "T1" && user.tier !== "T2" && tsAllTradeSchool(body.dayTypes, days);
    const insert = (status, xeroStatus, error) => env.DB.prepare("INSERT INTO timesheet_submissions (id, email, name, employee_id, week_start, week_end, days, total, status, xero_status, forwarded, last_error, created_at, decided_at, decided_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?, ?, ?, ?)").bind(sub.id, sub.email, sub.name, sub.employee_id, sub.week_start, sub.week_end, sub.days, sub.total, status, xeroStatus || null, error || null, sub.created_at, status === "sent" ? sub.created_at : null, status === "sent" ? (trade ? "auto trade school" : "auto") : null).run();
    // Tayler, Andrew, Mitch and Andrey are approved on the spot. The plumbers wait for Tayler.
    const autoOk = user.tier === "T1" || user.tier === "T2" || trade || /^(tayler|andrew|mitch|andrey)\b/i.test(String(user.name || "").trim());
    if (autoOk) {
      const out = await tsPushXero(rawEnv, sub);
      if (out.fail) return reply(env, 502, { error: out.fail });
      await insert("sent", out.xeroStatus);
      await tsClearReminder(env, email);
      await notify(env, { to: sub.email, kind: "timesheet", title: "Timesheet approved", body: `Your timesheet for ${tsDmy(sub.week_start)} to ${tsDmy(sub.week_end)} has been approved.`, link: "timesheets", by: "Reservoir" });
      return reply(env, 200, { ok: true, state: "sent", xeroStatus: out.xeroStatus, note: out.note || "", worker: WORKER_BUILD });
    }
    await insert("pending");
    await tsClearReminder(env, email);
    return reply(env, 200, { ok: true, state: "pending", worker: WORKER_BUILD });
  }
  if (parts[2] === "auto-run" && parts.length === 3 && request.method === "POST") {
    if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
    const actions = await tsAutoRun(rawEnv);
    return reply(env, 200, { ok: true, actions, worker: WORKER_BUILD });
  }
  if (parts[2] === "office" && parts.length === 3 && request.method === "GET") {
    if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
    return tsOfficeRoute(env, user);
  }
  if (parts[2] === "inbox" && parts.length === 3 && request.method === "GET") {
    if (!email) return reply(env, 400, { error: "No email on your sign in" });
    await tsTables(env);
    const notices = (((await env.DB.prepare("SELECT id, text, created_at, seen FROM ts_notices WHERE email = ? ORDER BY id DESC LIMIT 20").bind(email).all()).results) || []).map((n) => ({ id: n.id, text: n.text || "", createdAt: n.created_at || "", seen: n.seen === 1 }));
    const mine = (((await env.DB.prepare("SELECT week_start, status, xero_status FROM timesheet_submissions WHERE email = ? ORDER BY week_start DESC LIMIT 12").bind(email).all()).results) || []).map((r) => ({ week: r.week_start, status: r.status, xeroStatus: r.xero_status || "" }));
    let pending = [];
    if (tsCanApprove(user)) {
      const rows = ((await env.DB.prepare("SELECT * FROM timesheet_submissions WHERE status = 'pending' ORDER BY created_at ASC LIMIT 50").all()).results) || [];
      pending = rows.map(tsRowOut);
    }
    return reply(env, 200, { notices, mine, pending, worker: WORKER_BUILD });
  }
  if (parts[2] === "notices" && parts[3] === "seen" && parts.length === 4 && request.method === "POST") {
    await tsTables(env);
    await env.DB.prepare("UPDATE ts_notices SET seen = 1 WHERE email = ? AND seen = 0").bind(email).run();
    return reply(env, 200, { ok: true });
  }
  if (parts[2] === "approve" && parts.length === 4 && request.method === "POST") {
    if (!tsCanApprove(user)) return reply(env, 403, { error: "forbidden" });
    await tsTables(env);
    const row = await env.DB.prepare("SELECT * FROM timesheet_submissions WHERE id = ?").bind(parts[3]).first();
    if (!row) return reply(env, 404, { error: "That timesheet is not there any more." });
    if (row.status !== "pending") return reply(env, 409, { error: row.status === "sent" ? "That timesheet has already been approved." : "That timesheet is not waiting for approval." });
    const out = await tsPushXero(rawEnv, row);
    if (out.fail) {
      await env.DB.prepare("UPDATE timesheet_submissions SET last_error = ? WHERE id = ?").bind(out.fail.slice(0, 400), row.id).run();
      return reply(env, 502, { error: out.fail });
    }
    await env.DB.prepare("UPDATE timesheet_submissions SET status = 'sent', xero_status = ?, last_error = NULL, decided_at = ?, decided_by = ? WHERE id = ?").bind(out.xeroStatus, new Date().toISOString(), user.name || email, row.id).run();
    await notify(env, { to: row.email, kind: "timesheet", title: "Timesheet approved", body: `Your timesheet for ${tsDmy(row.week_start)} to ${tsDmy(row.week_end)} has been approved.`, link: "timesheets", by: user.name || email });
    return reply(env, 200, { ok: true, xeroStatus: out.xeroStatus, note: out.note || "" });
  }
  return null;
}
// Submitting a week clears the "Submit your timesheet" reminder so it does not sit in Notices.
async function tsClearReminder(env, email) {
  try { await noticesTable(env); await env.DB.prepare("UPDATE notices SET read_at = datetime('now') WHERE to_email = ? AND kind = 'timesheet' AND title = 'Submit your timesheet' AND read_at IS NULL").bind(String(email || "").toLowerCase()).run(); } catch (e) { /* best effort */ }
}
let timesheetDraftsReady = false;
async function timesheetDraftsTable(env) {
  if (timesheetDraftsReady) return;
  timesheetDraftsReady = true;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS timesheet_drafts (email TEXT NOT NULL, week_start TEXT NOT NULL, data TEXT NOT NULL, updated_at TEXT, PRIMARY KEY (email, week_start))").run();
}
async function timesheetRoutes(request, env, user, url) {
  if (!url.pathname.startsWith("/api/timesheets")) return null;
  const parts = url.pathname.split("/").filter(Boolean);
  try {
    // The signed in person's own synced Xero timesheets, newest week first. Any tier, only their own record (or the person being viewed as). Read only.
    if (parts.length === 3 && parts[2] === "me" && request.method === "GET") {
      const email = String(user.email || "").trim().toLowerCase();
      if (!email) return reply(env, 400, { error: "No email on your sign in" });
      await timesheetTable(env);
      const rows = ((await env.DB.prepare("SELECT id, start_date, end_date, status, total_hours, lines, xero_updated, synced_at FROM xero_timesheets WHERE lower(email) = ? AND end_date >= ? ORDER BY start_date DESC").bind(email, TIMESHEET_FROM).all()).results) || [];
      const items = rows.map((r) => {
        let lines;
        try { lines = JSON.parse(r.lines || "[]"); } catch (e) { lines = []; }
        return { id: r.id, start: r.start_date, end: r.end_date, status: r.status || "", totalHours: Number(r.total_hours) || 0, lines, updated: r.xero_updated || "" };
      });
      const synced = rows.reduce((a, r) => (r.synced_at && r.synced_at > a ? r.synced_at : a), "");
      // End date of the latest pay run filed in Reservoir. The app shows a week as Processed only once its pay run is filed, and as Approved until then.
      let filedEnd = "";
      try { await payslipTables(env); const fr = await env.DB.prepare("SELECT MAX(period_end) AS e FROM payslip_runs WHERE filed = 1").first(); filedEnd = fr && fr.e ? String(fr.e).slice(0, 10) : ""; } catch (e) { filedEnd = ""; }
      return reply(env, 200, { items, syncedAt: synced, filedEnd, worker: WORKER_BUILD });
    }
    if (["submit", "inbox", "notices", "approve", "office", "auto-run"].includes(parts[2])) {
      const out = await tsSubmitRoutes(request, env, user, url, parts);
      if (out) return out;
    }
    // The signed in person's own half filled weeks, saved as they type so a closed app loses nothing. Any tier, own record only.
    if (parts.length === 3 && parts[2] === "drafts" && request.method === "GET") {
      const email = String(user.email || "").trim().toLowerCase();
      if (!email) return reply(env, 400, { error: "No email on your sign in" });
      await timesheetDraftsTable(env);
      const rows = ((await env.DB.prepare("SELECT week_start, data, updated_at FROM timesheet_drafts WHERE email = ? ORDER BY week_start DESC LIMIT 60").bind(email).all()).results) || [];
      const weeks = [];
      for (const r of rows) {
        try { weeks.push({ key: r.week_start, period: JSON.parse(r.data), updated: r.updated_at || "" }); } catch (e) { /* skip a damaged row */ }
      }
      return reply(env, 200, { weeks, worker: WORKER_BUILD });
    }
    if (parts.length === 4 && parts[2] === "drafts" && request.method === "PUT") {
      const email = String(user.email || "").trim().toLowerCase();
      if (!email) return reply(env, 400, { error: "No email on your sign in" });
      const week = decodeURIComponent(parts[3]);
      if (!/^\d{4}-\d{1,2}-\d{1,2}$/.test(week)) return reply(env, 400, { error: "Bad week" });
      const body = await request.json().catch(() => null);
      const period = body && body.period;
      if (!period || typeof period !== "object") return reply(env, 400, { error: "No timesheet sent" });
      const text = JSON.stringify(period);
      if (text.length > 60000) return reply(env, 413, { error: "Timesheet too large" });
      await timesheetDraftsTable(env);
      await env.DB.prepare("INSERT INTO timesheet_drafts (email, week_start, data, updated_at) VALUES (?, ?, ?, ?) ON CONFLICT(email, week_start) DO UPDATE SET data = excluded.data, updated_at = excluded.updated_at").bind(email, week, text, new Date().toISOString()).run();
      return reply(env, 200, { ok: true });
    }
    if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
    if (parts.length === 3 && parts[2] === "probe") {
      const out = await timesheetProbe(env);
      return reply(env, 200, { ...out, worker: WORKER_BUILD });
    }
    if (parts.length === 3 && parts[2] === "sync" && request.method === "POST") {
      const out = await timesheetSync(env);
      return reply(env, out.ok ? 200 : 502, { ...out, worker: WORKER_BUILD });
    }
    if (parts.length === 3 && parts[2] === "status" && request.method === "GET") {
      await timesheetTable(env);
      const r = await env.DB.prepare("SELECT COUNT(*) AS n, COUNT(DISTINCT employee_id) AS people, MIN(start_date) AS earliest, MAX(end_date) AS latest, MAX(synced_at) AS synced FROM xero_timesheets").first();
      return reply(env, 200, { count: Number(r && r.n) || 0, people: Number(r && r.people) || 0, earliest: (r && r.earliest) || "", latest: (r && r.latest) || "", syncedAt: (r && r.synced) || "", worker: WORKER_BUILD });
    }
    return reply(env, 404, { error: "not found" });
  } catch (e) {
    console.log("timesheets error", String(e && e.message || e));
    return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
  }
}
// All the /api/leave routes. Returns a Response, or null when the path is not a leave path.
async function leaveRoutes(request, env, user, url) {
  if (!url.pathname.startsWith("/api/leave")) return null;
  const office = user.tier === "T1" || user.tier === "T2";
  const email = String(user.email || "").trim().toLowerCase();
  const parts = url.pathname.split("/").filter(Boolean); // api, leave, ...
  const method = request.method;
  const body = async () => { try { return await request.json(); } catch (e) { return {}; } };
  const today = melbourneDay();
  const mineOnly = async (id) => {
    const row = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(id).first();
    return row && String(row.email).toLowerCase() === email ? row : null;
  };
  try {
    await leaveTable(env);
    if (parts.length === 3 && parts[2] === "me" && method === "GET") {
      if (!email) return reply(env, 400, { error: "No email on your sign in" });
      const since = melbourneDay(Date.now() - 366 * 86400000);
      const rows = ((await env.DB.prepare("SELECT * FROM xero_leave WHERE email = ? AND status IN ('pending', 'approved') AND to_date >= ? AND split_of IS NULL ORDER BY from_date DESC").bind(email, since).all()).results) || [];
      const bal = await leaveBalances(env, email);
      const dm = await leaveDayMap(env);
      return reply(env, 200, { balance: { found: bal.found, hasLeave: bal.hasLeave, annualHours: bal.annualHours, personalHours: bal.personalHours, syncedAt: bal.syncedAt, dayHours: leaveDayFrom(dm, email, user.name) }, items: rows.map((x) => leaveOut(x, dm)), worker: WORKER_BUILD });
    }
    if (parts.length === 3 && parts[2] === "request" && method === "POST") {
      if (!email) return reply(env, 400, { error: "No email on your sign in" });
      const b = await body();
      const type = leaveTypeOf(b.type);
      if (!type) return reply(env, 400, { error: "Choose a leave type" });
      if (!leaveIsoOk(b.from) || !leaveIsoOk(b.to) || b.to < b.from) return reply(env, 400, { error: "Check the dates" });
      const clash = await env.DB.prepare("SELECT id FROM xero_leave WHERE email = ? AND status IN ('pending', 'approved') AND from_date <= ? AND to_date >= ?").bind(email, b.to, b.from).first();
      if (clash) return reply(env, 409, { error: "You already have leave over some of those dates" });
      const have = await leaveBalances(env, email);
      if (have.found && !have.hasLeave) return reply(env, 403, { error: "You have no leave set up in Xero, so there is nothing to request" });
      const pay = await env.DB.prepare("SELECT employee_id FROM xero_payroll WHERE email = ?").bind(email).first();
      const id = await leaveCreate(env, { email, name: user.name || "", employeeId: pay ? pay.employee_id : null, type, from: b.from, to: b.to });
      return reply(env, 200, { ok: true, id });
    }
    if (parts.length === 4 && parts[2] !== "me" && parts[3] === "edit" && method === "POST") {
      const row = await mineOnly(parts[2]);
      if (!row) return reply(env, 404, { error: "not found" });
      const b = await body();
      const type = leaveTypeOf(b.type);
      if (!type || !leaveIsoOk(b.from) || !leaveIsoOk(b.to) || b.to < b.from) return reply(env, 400, { error: "Check the leave type and dates" });
      const clash = await env.DB.prepare("SELECT id FROM xero_leave WHERE email = ? AND status IN ('pending', 'approved') AND from_date <= ? AND to_date >= ? AND id <> ? AND (split_of IS NULL OR split_of <> ?)").bind(email, b.to, b.from, row.id, row.id).first();
      if (clash) return reply(env, 409, { error: "You already have leave over some of those dates" });
      if (row.status === "pending") {
        await env.DB.prepare("UPDATE xero_leave SET asked_type = ? WHERE id = ?").bind(type, row.id).run();
        await leaveReplan(env, { ...row, asked_type: type }, type, b.from, b.to);
        return reply(env, 200, { ok: true, requested: false });
      }
      if (row.status === "approved" && row.from_date > today) {
        await env.DB.prepare("UPDATE xero_leave SET edit_requested = 1, new_type = ?, new_from = ?, new_to = ? WHERE id = ?").bind(type, b.from, b.to, row.id).run();
        return reply(env, 200, { ok: true, requested: true });
      }
      return reply(env, 409, { error: "That leave can no longer be changed" });
    }
    if (parts.length === 4 && parts[2] !== "me" && parts[3] === "cancel" && method === "POST") {
      const row = await mineOnly(parts[2]);
      if (!row) return reply(env, 404, { error: "not found" });
      if (row.status === "pending") {
        await env.DB.prepare("DELETE FROM xero_leave WHERE id = ? OR (split_of = ? AND status = 'pending')").bind(row.id, row.id).run();
        return reply(env, 200, { ok: true, requested: false });
      }
      if (row.status === "approved" && row.from_date > today) {
        await env.DB.prepare("UPDATE xero_leave SET cancel_requested = 1 WHERE id = ?").bind(row.id).run();
        return reply(env, 200, { ok: true, requested: true });
      }
      return reply(env, 409, { error: "That leave can no longer be cancelled" });
    }
    if (parts.length === 3 && parts[2] === "office" && method === "GET") {
      if (!office) return reply(env, 403, { error: "forbidden" });
      const recent = melbourneDay(Date.now() - 30 * 86400000);
      const rows = ((await env.DB.prepare("SELECT * FROM xero_leave WHERE status = 'pending' OR cancel_requested = 1 OR edit_requested = 1 OR (status IN ('approved', 'declined', 'cancelled') AND source = 'reservoir' AND decided_at >= ?) ORDER BY from_date ASC").bind(recent).all()).results) || [];
      const away = ((await env.DB.prepare("SELECT * FROM xero_leave WHERE status = 'approved' AND from_date <= ? AND to_date >= ?").bind(today, today).all()).results) || [];
      const balances = new Map();
      const dm = await leaveDayMap(env);
      const withBalance = async (r) => {
        const out = leaveOut(r, dm);
        if (r.status === "pending") {
          if (!balances.has(r.email)) { const bal = await leaveBalances(env, r.email); balances.set(r.email, { annualHours: bal.annualHours, personalHours: bal.personalHours }); }
          out.balance = balances.get(r.email);
        }
        return out;
      };
      const items = [];
      for (const r of rows) items.push(await withBalance(r));
      return reply(env, 200, { items, away: away.map((x) => leaveOut(x, dm)), worker: WORKER_BUILD });
    }
    if (parts.length === 4 && parts[3] === "decide" && method === "POST") {
      if (!isApprover(user)) return reply(env, 403, { error: "forbidden" });
      const row = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(parts[2]).first();
      if (!row || row.status !== "pending") return reply(env, 409, { error: "That request is no longer waiting" });
      const b = await body();
      if (b.decision === "decline") {
        await env.DB.prepare("UPDATE xero_leave SET status = 'declined', decided_by = ?, decided_at = datetime('now') WHERE id = ? OR (split_of = ? AND status = 'pending')").bind(leaveActor(user), row.id, row.id).run();
        return reply(env, 200, { ok: true });
      }
      if (b.decision !== "approve") return reply(env, 400, { error: "bad decision" });
      let live = row;
      const want = leaveTypeOf(b.type);
      if (want && (want !== row.type || row.type === "unpaid")) {
        if (row.split_of) return reply(env, 400, { error: "Change the leave type on the main request, not on the unpaid part" });
        await leaveReplan(env, row, want, row.from_date, row.to_date);
        live = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(row.id).first();
      }
      const pushed = leaveDaysOf(live) ? await leavePushDays(env, live) : await leavePushXero(env, live, { type: live.type, from: live.from_date, to: live.to_date });
      if (pushed.fail) return reply(env, 502, { error: "Xero did not take the leave, so it is still waiting here", detail: pushed.fail });
      await env.DB.prepare("UPDATE xero_leave SET status = 'approved', xero_id = ?, decided_by = ?, decided_at = datetime('now') WHERE id = ?").bind(pushed.id, leaveActor(user), row.id).run();
      await leaveRefreshBalances(env, row.email);
      return reply(env, 200, { ok: true });
    }
    if (parts.length === 4 && parts[3] === "retype" && method === "POST") {
      if (!office) return reply(env, 403, { error: "forbidden" });
      const row = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(parts[2]).first();
      if (!row || row.status !== "pending") return reply(env, 409, { error: "That request is no longer waiting" });
      if (row.split_of) return reply(env, 400, { error: "Change the leave type on the main request, not on the unpaid part" });
      const b = await body();
      const want = leaveTypeOf(b.type);
      if (!want) return reply(env, 400, { error: "Choose a leave type" });
      await leaveReplan(env, row, want, row.from_date, row.to_date);
      return reply(env, 200, { ok: true });
    }
    // The office edits a waiting request before approving it: the leave type and the hours on each day. Nothing goes to Xero until Approve.
    if (parts.length === 4 && parts[3] === "adjust" && method === "POST") {
      if (!office) return reply(env, 403, { error: "forbidden" });
      const row = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(parts[2]).first();
      if (!row || row.status !== "pending") return reply(env, 409, { error: "That request is no longer waiting" });
      const b = await body();
      const DAY = await leaveDayHours(env, row.email);
      const allowed = new Set(leavePlanOf(row, DAY).map((x) => x.date));
      for (const d of leaveWorkDays(row.from_date, row.to_date)) allowed.add(d);
      if (!Array.isArray(b.days) || !b.days.length) return reply(env, 400, { error: "There are no days to save" });
      const days = [];
      const seen = new Set();
      for (const x of b.days.slice(0, 400)) {
        const date = String((x && x.date) || "");
        const hours = Math.round(Number(x && x.hours) * 10000) / 10000;
        if (!leaveIsoOk(date) || seen.has(date)) return reply(env, 400, { error: "Check the days" });
        if (!allowed.has(date)) return reply(env, 400, { error: `${date} is not one of the days on this request` });
        if (!Number.isFinite(hours) || hours < 0 || hours > 24) return reply(env, 400, { error: "Hours for a day must be between 0 and 24" });
        seen.add(date);
        if (hours > 0) days.push({ date, hours });
      }
      if (!days.length) return reply(env, 400, { error: "Every day is at 0 hours. Decline the request instead." });
      days.sort((a, c) => (a.date < c.date ? -1 : 1));
      let type = row.type;
      const want = leaveTypeOf(b.type);
      if (want && !row.split_of && want !== row.type && !(row.type === "unpaid" && want === (row.asked_type || ""))) type = want;
      // A hand edited plan replaces the automatic split, so any unpaid part made for it goes.
      if (!row.split_of) await env.DB.prepare("DELETE FROM xero_leave WHERE split_of = ? AND status = 'pending'").bind(row.id).run();
      const keepAsked = type === "unpaid" ? row.asked_type || "personal" : type;
      await env.DB.prepare("UPDATE xero_leave SET type = ?, type_name = ?, asked_type = ?, from_date = ?, to_date = ?, days = ? WHERE id = ?").bind(type, LEAVE_LABELS[type] || row.type_name, row.split_of ? row.asked_type : keepAsked, days[0].date, days[days.length - 1].date, JSON.stringify(days), row.id).run();
      return reply(env, 200, { ok: true });
    }
    if (parts.length === 4 && parts[3] === "decide-cancel" && method === "POST") {
      if (!office) return reply(env, 403, { error: "forbidden" });
      const row = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(parts[2]).first();
      if (!row || Number(row.cancel_requested) !== 1) return reply(env, 409, { error: "That cancellation is no longer waiting" });
      const b = await body();
      if (b.decision === "keep") {
        await env.DB.prepare("UPDATE xero_leave SET cancel_requested = 0 WHERE id = ?").bind(row.id).run();
        return reply(env, 200, { ok: true });
      }
      if (b.decision !== "cancel") return reply(env, 400, { error: "bad decision" });
      // Xero's payroll API has a reject call for each leave application. Every day sent to Xero is rejected there. If Xero refuses any of them, the leave is still cancelled here and the note says which ones the office must remove by hand.
      const rejected = await leaveRejectInXero(env, row);
      const note = rejected.fail ? `Cancelled here. Xero would not reject all of it (${rejected.fail}), so remove the rest in Xero by hand.` : "";
      await env.DB.prepare("UPDATE xero_leave SET status = 'cancelled', cancel_requested = 0, xero_note = ?, decided_by = ?, decided_at = datetime('now') WHERE id = ?").bind(note, leaveActor(user), row.id).run();
      return reply(env, 200, { ok: true, note });
    }
    if (parts.length === 4 && parts[3] === "decide-edit" && method === "POST") {
      if (!office) return reply(env, 403, { error: "forbidden" });
      const row = await env.DB.prepare("SELECT * FROM xero_leave WHERE id = ?").bind(parts[2]).first();
      if (!row || Number(row.edit_requested) !== 1) return reply(env, 409, { error: "That change is no longer waiting" });
      const b = await body();
      if (b.decision === "keep") {
        await env.DB.prepare("UPDATE xero_leave SET edit_requested = 0, new_type = NULL, new_from = NULL, new_to = NULL WHERE id = ?").bind(row.id).run();
        return reply(env, 200, { ok: true });
      }
      if (b.decision !== "apply") return reply(env, 400, { error: "bad decision" });
      const change = { type: row.new_type || row.type, from: row.new_from || row.from_date, to: row.new_to || row.to_date };
      if (leaveDaysOf(row)) return reply(env, 409, { error: "This leave went to Xero one day at a time. Change it in Xero by hand, then press Keep Leave here." });
      if (row.xero_id) {
        const pushed = await leavePushXero(env, row, change);
        if (pushed.fail) return reply(env, 502, { error: "Xero did not take the change, so the old dates stand", detail: pushed.fail });
      }
      await env.DB.prepare("UPDATE xero_leave SET type = ?, type_name = ?, from_date = ?, to_date = ?, edit_requested = 0, new_type = NULL, new_from = NULL, new_to = NULL, decided_by = ?, decided_at = datetime('now') WHERE id = ?").bind(change.type, LEAVE_LABELS[change.type] || row.type_name, change.from, change.to, leaveActor(user), row.id).run();
      await leaveRefreshBalances(env, row.email);
      return reply(env, 200, { ok: true });
    }
    if (parts.length === 3 && parts[2] === "sync" && method === "POST") {
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      const out = await leaveSync(env);
      return reply(env, out.ok ? 200 : 502, { ...out, worker: WORKER_BUILD });
    }
    return reply(env, 404, { error: "not found" });
  } catch (e) {
    console.log("leave error", String(e && e.message || e));
    return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
  }
}

// ---- Payroll sweeps. Xero is the source for payroll, so Reservoir keeps checking it and does what the Payroll screen buttons do, so nobody has to remember to tap them. They run from the half hourly cron and check the Melbourne clock themselves, so the first run in a window is the first half hour mark at or after its start time. ----
// Tuesday from 10:00 to 14:00: looks for a posted pay run and, if one is there, stores its slips exactly as Post Pay Run does (without replacing a run already stored). Stops once a run waiting to be filed exists. If nothing is posted by 14:00, sends one notice to Tayler.
// Thursday from 7:45: files the pay run that is waiting, exactly as File Pay Run does (marks it filed, moves its slips from Next Pay to Paid, banks the processed timesheets). If Tuesday was missed it first looks for a posted run in Xero. Only a run with a payment date of today or earlier is filed. Stops once a run paying today is filed.
// The last thing each sweep did is kept in app_settings under sweep.payroll_last, so it can be checked.
async function payrollSweepNote(env, day, text) {
  try {
    await settingsTable(env);
    await env.DB.prepare("INSERT INTO app_settings (key, value, detail, by_name, at) VALUES ('sweep.payroll_last', ?, ?, 'Payroll sweep', datetime('now')) ON CONFLICT(key) DO UPDATE SET value = excluded.value, detail = excluded.detail, by_name = excluded.by_name, at = excluded.at").bind(day, String(text).slice(0, 500)).run();
  } catch (e) { /* the note is best effort */ }
}
async function payrollSweep(env) {
  const now = tsMelbourne(new Date());
  const tue = now.dow === 2 && now.mins >= 10 * 60 && now.mins <= 14 * 60;
  const tueAfter = now.dow === 2 && now.mins > 14 * 60;
  const thu = now.dow === 4 && now.mins >= 7 * 60 + 45;
  if (!tue && !tueAfter && !thu) return;
  await payslipTables(env);
  const waitingRun = (from) => env.DB.prepare("SELECT r.run_id, r.payment_date FROM payslip_runs r WHERE r.filed = 0 AND r.rules >= ? AND r.payment_date >= ? AND EXISTS (SELECT 1 FROM payslip_store s WHERE s.run_id = r.run_id) ORDER BY r.payment_date DESC LIMIT 1").bind(PAYSLIP_RULES, from || "").first();
  if (tue) {
    if (await waitingRun(now.date)) return;
    const pull = await payslipPull(env, { runs: 3 });
    if (pull && pull.fail) { await payrollSweepNote(env, now.date, `Tuesday check failed: ${String(pull.fail.error || "").slice(0, 200)}`); return; }
    const got = await waitingRun(now.date);
    if (got) await payrollSweepNote(env, now.date, `Pay run paying ${got.payment_date} found posted in Xero and stored. Slips now show under Next Pay. Slips stored this check: ${(pull && pull.slipsStored) || 0}`);
    else await payrollSweepNote(env, now.date, "Tuesday check: no posted pay run in Xero yet");
    return;
  }
  if (tueAfter) {
    // Past 14:00 on a Tuesday with nothing posted: one notice only.
    if (await waitingRun(now.date)) return;
    await settingsTable(env);
    const sent = await env.DB.prepare("SELECT value FROM app_settings WHERE key = 'sweep.payroll_unposted'").first();
    if (sent && sent.value === now.date) return;
    const rows = ((await env.DB.prepare("SELECT name, email FROM xero_payroll WHERE email IS NOT NULL AND email <> '' AND (status IS NULL OR status = 'ACTIVE')").all()).results) || [];
    const me = rows.find((p) => /^tayler\b/i.test(String(p.name || "")));
    const id = await notify(env, { to: me ? me.email : "@office", kind: "payroll", title: "Payroll not posted", body: "Nothing has been posted in Xero for this pay run yet. Reservoir checked until 2pm.", by: "Payroll sweep" });
    if (id) await env.DB.prepare("INSERT INTO app_settings (key, value, detail, by_name, at) VALUES ('sweep.payroll_unposted', ?, 'notice sent', 'Payroll sweep', datetime('now')) ON CONFLICT(key) DO UPDATE SET value = excluded.value, detail = excluded.detail, at = excluded.at").bind(now.date).run();
    await payrollSweepNote(env, now.date, id ? "Tuesday 2pm: nothing posted, notice sent" : "Tuesday 2pm: nothing posted, notice failed to send");
    return;
  }
  // Thursday
  const doneToday = await env.DB.prepare("SELECT run_id FROM payslip_runs WHERE payment_date = ? AND filed = 1 AND rules >= ? LIMIT 1").bind(now.date, PAYSLIP_RULES).first();
  if (doneToday) return;
  let run = await waitingRun();
  if (!run) {
    const pull = await payslipPull(env, { runs: 3 });
    if (pull && pull.fail) { await payrollSweepNote(env, now.date, `Thursday check failed: ${String(pull.fail.error || "").slice(0, 200)}`); return; }
    run = await waitingRun();
  }
  if (!run) { await payrollSweepNote(env, now.date, "Thursday check: no posted pay run to file"); return; }
  if (String(run.payment_date) > now.date) { await payrollSweepNote(env, now.date, `Thursday check: waiting run pays ${run.payment_date}, not filed yet`); return; }
  await env.DB.prepare("UPDATE payslip_runs SET filed = 1, filed_at = datetime('now'), filed_by = 'Payroll sweep' WHERE run_id = ? AND filed = 0").bind(run.run_id).run();
  // The same one pull the button makes after filing: processed timesheets are banked as permanent history. A failure here does not undo the filing.
  let ts = "timesheets not banked";
  try { const r = await timesheetSync(env); ts = r && r.ok ? `timesheets banked: ${r.saved || 0}` : `timesheets failed: ${String((r && r.error) || "").slice(0, 120)}`; } catch (e) { ts = `timesheets failed: ${String(e && e.message || e).slice(0, 120)}`; }
  await payrollSweepNote(env, now.date, `Filed pay run paid ${run.payment_date}. ${ts}`);
}

// One off test: Saturday 24 Oct 2026, first half hour tick at or after 3:30pm Melbourne sends Andrey a timesheet reminder push. Sends once only.
async function pushOneOffAndrey(env) {
  const now = tsMelbourne(new Date());
  if (now.date !== "2026-10-24" || now.mins < 15 * 60 + 30 || now.mins > 18 * 60) return;
  await settingsTable(env);
  const done = await env.DB.prepare("SELECT value FROM app_settings WHERE key = 'oneoff.andrey_ts_push'").first();
  if (done) return;
  await env.DB.prepare("INSERT INTO app_settings (key, value, detail, by_name, at) VALUES ('oneoff.andrey_ts_push', '1', 'sent', 'One off', datetime('now'))").run();
  await pushTable(env);
  const rows = ((await env.DB.prepare("SELECT DISTINCT email FROM push_subs WHERE lower(email) LIKE 'andrey%'").all()).results) || [];
  for (const r of rows) await notify(env, { to: r.email, kind: "timesheet", title: "Submit your timesheet", body: "Please submit your timesheet for this week.", link: "timesheets", by: "Reminder" });
}
// Xero contacts, names only. Keeps each contact's Xero ID and name so a repair record can name who the tax invoice is from. Read only, and nothing else about the contact is kept.
let xeroContactsReady = false;
async function xeroContactsTable(env) {
  if (xeroContactsReady) return;
  await env.DB.prepare("CREATE TABLE IF NOT EXISTS xero_contacts (contact_id TEXT PRIMARY KEY, name TEXT NOT NULL, name_lc TEXT NOT NULL, synced_at TEXT)").run();
  await env.DB.prepare("CREATE INDEX IF NOT EXISTS xero_contacts_name ON xero_contacts (name_lc)").run();
  xeroContactsReady = true;
}
async function xeroContactsSync(env) {
  await xeroContactsTable(env);
  const stamp = new Date().toISOString();
  const pending = [];
  let saved = 0;
  const flush = async () => {
    while (pending.length) {
      const group = pending.splice(0, 40);
      if (typeof env.DB.batch === "function") await env.DB.batch(group);
      else for (const st of group) await st.run();
    }
  };
  for (let page = 1; page <= 100; page++) {
    const res = await xeroFetch(env, `api.xro/2.0/Contacts?page=${page}&summaryOnly=true`);
    if (!res.ok) { await flush(); return { ok: false, saved, error: await xeroWhy(res) }; }
    const list = (await res.json()).Contacts || [];
    if (!list.length) break;
    for (const c of list) {
      const name = String(c.Name || "").trim().slice(0, 200);
      if (!c.ContactID || !name) continue;
      pending.push(env.DB.prepare("INSERT INTO xero_contacts (contact_id, name, name_lc, synced_at) VALUES (?, ?, ?, ?) ON CONFLICT(contact_id) DO UPDATE SET name = excluded.name, name_lc = excluded.name_lc, synced_at = excluded.synced_at").bind(c.ContactID, name, name.toLowerCase(), stamp));
      saved++;
    }
    if (list.length < 100) break;
  }
  await flush();
  // A full read worked, so a contact that is no longer in Xero (archived or deleted) goes from the list too.
  await env.DB.prepare("DELETE FROM xero_contacts WHERE synced_at IS NULL OR synced_at <> ?").bind(stamp).run();
  return { ok: true, saved };
}
// Runs one Xero pull at most once per key (read only). A failed pull clears its key so the next half hourly tick tries again.
async function xeroPullOnce(env, key, label, fn) {
  await settingsTable(env);
  const done = await env.DB.prepare("SELECT value FROM app_settings WHERE key = ?").bind(key).first();
  if (done) return;
  await env.DB.prepare("INSERT INTO app_settings (key, value, detail, by_name, at) VALUES (?, '1', ?, 'Scheduled', datetime('now'))").bind(key, `${label} started`).run();
  let ok = false, note = "";
  try { const r = await fn(env); ok = !!(r && r.ok); note = ok ? r.note : `failed: ${String((r && r.error) || "").slice(0, 160)}`; } catch (e) { note = `failed: ${String(e && e.message || e).slice(0, 160)}`; }
  if (ok) await env.DB.prepare("UPDATE app_settings SET detail = ?, at = datetime('now') WHERE key = ?").bind(`${label} ${note}`, key).run();
  else await env.DB.prepare("DELETE FROM app_settings WHERE key = ?").bind(key).run();
}
async function tsPullFn(env) { const r = await timesheetSync(env); return r && r.ok ? { ok: true, note: `pulled ${r.saved || 0} timesheets, ${r.people || 0} people, ${JSON.stringify(r.byStatus || {})}` } : { ok: false, error: r && r.error }; }
async function contactsPullFn(env) { const r = await xeroContactsSync(env); return r.ok ? { ok: true, note: `${r.saved} contacts` } : { ok: false, error: r.error }; }
// Timesheets: Monday from 4:00pm Melbourne, and once as a one off after w158. Contacts: every weekday from 6:00pm Melbourne, and once as a one off after w163.
async function tsXeroPullSchedule(env) {
  const now = tsMelbourne(new Date());
  if (now.date >= "2026-10-26") await xeroPullOnce(env, "oneoff.ts_pull_20261026", "One off timesheet pull", tsPullFn);
  if (now.dow === 1 && now.mins >= 16 * 60 && now.mins < 17 * 60) await xeroPullOnce(env, `sweep.ts_pull.${now.date}`, "Monday timesheet pull", tsPullFn);
  if (now.date >= "2026-10-26") await xeroPullOnce(env, "oneoff.contacts_pull_20261026", "One off contacts pull", contactsPullFn);
  if (now.dow >= 1 && now.dow <= 5 && now.mins >= 18 * 60 && now.mins < 19 * 60) await xeroPullOnce(env, `sweep.contacts_pull.${now.date}`, "Weekday contacts pull", contactsPullFn);
}
export default {
  // Runs on a Cloudflare cron trigger if one is added (for example Thursday early morning). It only refreshes leave balances and leave history from Xero into D1.
  async scheduled(event, env, ctx) {
    // Every run checks the automatic timesheets and the payroll sweeps. The leave sync keeps its own schedule and is skipped on the half hourly one.
    ctx.waitUntil(tsAutoRun(env).catch((e) => console.log("auto timesheets failed", String(e && e.message || e))));
    ctx.waitUntil(pushFlushHeld(env));
    ctx.waitUntil(dailyOpsAlert(env).catch((e) => console.log("daily ops alert failed", String(e && e.message || e))));
    ctx.waitUntil(tsResendSalaried(env).catch((e) => console.log("resend salaried failed", String(e && e.message || e))));
    ctx.waitUntil(tsXeroPullSchedule(env).catch((e) => console.log("timesheet pull failed", String(e && e.message || e))));
    ctx.waitUntil(pushOneOffAndrey(env).catch((e) => console.log("one off failed", String(e && e.message || e))));
    ctx.waitUntil(payrollSweep(env).catch((e) => console.log("payroll sweep failed", String(e && e.message || e))));
    if (event && event.cron === "*/30 * * * *") return;
    ctx.waitUntil(leaveSync(env).catch((e) => console.log("leave sync failed", String(e && e.message || e))));
  },
  async fetch(request, rawEnv) {
    let env = rawEnv;
    if (request.method === "OPTIONS") return new Response(null, { status: 204, headers: cors(env) });
    const earlyUrl = new URL(request.url);
    if (earlyUrl.pathname === "/api/ping" && request.method === "GET") return reply(env, 200, { worker: WORKER_BUILD, xeroMe: true, xeroStaff: true });
    if (earlyUrl.pathname === "/api/xero/callback" && request.method === "GET") {
      // Xero sends the browser back here after consent. No Microsoft token on a redirect, so the signed state proves who started it.
      const back = (flag) => Response.redirect(`${env.ALLOWED_ORIGIN}/?xero=${flag}`, 302);
      try {
        if (!env.XERO_CLIENT_SECRET) return back("not_set_up");
        const st = await xeroStateOk(env, earlyUrl.searchParams.get("state"));
        const code = earlyUrl.searchParams.get("code");
        if (!st || !code) return back("failed");
        const t = await xeroTokenCall(env, { grant_type: "authorization_code", code, redirect_uri: xeroRedirect(env) });
        const cr = await fetch("https://api.xero.com/connections", { headers: { Authorization: `Bearer ${t.access_token}`, Accept: "application/json" } });
        await xeroCountCall(env, cr);
        const conns = cr.ok ? await cr.json() : [];
        const org = conns.find((c) => c.tenantType === "ORGANISATION") || conns[0];
        if (!org) return back("no_org");
        await xeroTable(env);
        await env.DB.prepare("INSERT INTO xero_auth (id, refresh_token, access_token, expires_at, tenant_id, tenant_name, connected_by) VALUES (1, ?, ?, ?, ?, ?, ?) ON CONFLICT(id) DO UPDATE SET refresh_token = excluded.refresh_token, access_token = excluded.access_token, expires_at = excluded.expires_at, tenant_id = excluded.tenant_id, tenant_name = excluded.tenant_name, connected_by = excluded.connected_by, updated_at = datetime('now')").bind(t.refresh_token, t.access_token, Date.now() + t.expires_in * 1000, org.tenantId, org.tenantName || "", st.name || "").run();
        return back("connected");
      } catch (e) { return back("failed"); }
    }
    let user;
    try { user = await verifyToken(request, env); } catch (e) { return reply(env, 401, { error: "unauthorised" }); }

    const url = new URL(request.url);

    // View As. Only a real director (T1) can send this header. The request then runs as the person named in it, at the access level the app sent. Everything is read only, with no exceptions.
    const viewHeader = request.headers.get("X-View-As");
    if (viewHeader) {
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      let v;
      try { v = JSON.parse(decodeURIComponent(viewHeader)); } catch (e) { v = null; }
      if (!v || !v.e || !["T1", "T2", "T3"].includes(v.t)) return reply(env, 400, { error: "bad view as" });
      if (request.method !== "GET") return reply(env, 403, { error: "Viewing as someone else is read only." });
      console.log("view as", user.email, "->", String(v.e).toLowerCase(), v.t);
      user = { tier: v.t, name: String(v.n || ""), email: String(v.e).toLowerCase(), viewAs: true, realUser: user };
    }

    // From here on every call carries the caller's access level, which is what the Xero gate reads. A level that is missing or unknown is treated as closed.
    env = Object.create(rawEnv, { __callerTier: { value: String(user.tier || "T3") } });

    try { await phonesNormalise(env); } catch (e) { console.log("phones normalise failed", String(e && e.message || e).slice(0, 160)); }
    const leaveReply = await leaveRoutes(request, env, user, url);
    if (leaveReply) return leaveReply;
    const timesheetReply = await timesheetRoutes(request, env, user, url);
    if (timesheetReply) return timesheetReply;
    const pushReply = await pushRoutes(request, env, user, url);
    if (pushReply) return pushReply;
    const reqsReply = await appRequestsRoutes(request, env, user, url);
    if (reqsReply) return reqsReply;
    const partnersReply = await partnersRoutes(request, env, user, url);
    if (partnersReply) return partnersReply;
    const todosReply = await todosRoutes(request, env, user, url);
    if (todosReply) return todosReply;
    const noticesReply = await noticesRoutes(request, env, user, url);
    if (noticesReply) return noticesReply;
    // The app background everyone sees. Any tier can read which one is chosen, only a director (T1) can change it. The pictures themselves are built into the app, so only the name is stored.
    if (url.pathname === "/api/background") {
      const KEYS = ["wave", "papercut", "curve", "aqua", "orib", "ofac", "oecl"];
      if (request.method === "GET") {
        await launchTable(env);
        const r = await env.DB.prepare("SELECT value FROM app_settings WHERE key = 'launch.app_background'").first();
        const v = r && KEYS.includes(r.value) ? r.value : "wave";
        return reply(env, 200, { key: v, worker: WORKER_BUILD });
      }
      if (request.method === "PUT") {
        if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
        let b;
        try { b = await request.json(); } catch (e) { b = {}; }
        if (!KEYS.includes(b.key)) return reply(env, 400, { error: "That background is not available" });
        await launchTable(env);
        await env.DB.prepare("INSERT INTO app_settings (key, value, by_name, at) VALUES ('launch.app_background', ?, ?, datetime('now')) ON CONFLICT(key) DO UPDATE SET value = excluded.value, by_name = excluded.by_name, at = datetime('now')").bind(b.key, user.name || user.email || "").run();
        return reply(env, 200, { key: b.key, worker: WORKER_BUILD });
      }
    }
    const arofloReply = await arofloRoutes(request, env, user, url);
    if (arofloReply) return arofloReply;

    if (url.pathname === "/api/policies/acks" && request.method === "GET") {
      // The signed in person's own acknowledgements. The latest one for each policy.
      try {
        await policyAckTable(env);
        const who = ppePerson(user);
        if (!who.key) return reply(env, 200, { acks: {} });
        const rows = ((await env.DB.prepare("SELECT policy_id, version, opened_at, read_at, acked_at FROM policy_ack_log WHERE person_key = ? ORDER BY id").bind(who.key).all()).results) || [];
        const acks = {};
        for (const r of rows) acks[r.policy_id] = { version: r.version, openedAt: r.opened_at, readAt: r.read_at, ackedAt: r.acked_at };
        return reply(env, 200, { acks });
      } catch (e) {
        console.log("policy acks read error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not read the acknowledgements" });
      }
    }
    const ackRoute = url.pathname.match(/^\/api\/policies\/([^/]+)\/ack$/);
    if (ackRoute && request.method === "POST") {
      // An acknowledgement. Every one is kept for good: who, which policy and version, when it was opened, when it was read to the end, and when it was acknowledged.
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const iso = (v) => { const t = Date.parse(String(v || "")); return Number.isFinite(t) ? new Date(t).toISOString() : ""; };
      const openedAt = iso(body.openedAt), readAt = iso(body.readAt);
      if (!openedAt || !readAt) return reply(env, 400, { error: "Open the policy and read it to the end before acknowledging." });
      try {
        await policyTable(env);
        await policyAckTable(env);
        const id = decodeURIComponent(ackRoute[1]);
        const pol = await env.DB.prepare("SELECT id, title, version FROM policies WHERE id = ? AND removed = 0").bind(id).first();
        if (!pol) return reply(env, 404, { error: "That policy was not found" });
        const who = ppePerson(user);
        if (!who.key) return reply(env, 400, { error: "No name on your sign in" });
        const ackedAt = new Date().toISOString();
        await env.DB.prepare("INSERT INTO policy_ack_log (person_key, name, email, policy_id, policy_title, version, opened_at, read_at, acked_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)")
          .bind(who.key, who.name, who.email, pol.id, pol.title, pol.version || 1, openedAt, readAt, ackedAt).run();
        return reply(env, 200, { saved: true, ackedAt, version: pol.version || 1 });
      } catch (e) {
        console.log("policy ack error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not save the acknowledgement" });
      }
    }
    if (url.pathname === "/api/policies/search" && request.method === "GET") {
      // Staff search: matches the policy title or the text inside the current Word version. Any signed in tier.
      try {
        await policyTable(env);
        const terms = String(url.searchParams.get("q") || "").toLowerCase().split(/\s+/).map((x) => x.replace(/[%_\\]/g, "")).filter((x) => x.length >= 2).slice(0, 6);
        if (!terms.length) return reply(env, 200, { hits: [] });
        const where = terms.map(() => "(lower(title) LIKE ? ESCAPE '\\' OR lower(COALESCE(body, '')) LIKE ? ESCAPE '\\')").join(" AND ");
        const args = terms.flatMap((x) => [`%${x}%`, `%${x}%`]);
        const rows = await env.DB.prepare(`SELECT id, title, body FROM policies WHERE removed = 0 AND COALESCE(kind, 'policy') = ? AND ${where} ORDER BY title COLLATE NOCASE LIMIT 30`).bind(url.searchParams.get("kind") === "handbook" ? "handbook" : "policy", ...args).all();
        const hits = (rows.results || []).map((r) => {
          const body = String(r.body || "");
          const at = body.toLowerCase().indexOf(terms[0]);
          const inTitle = String(r.title || "").toLowerCase().includes(terms[0]);
          let snippet = "";
          if (at >= 0 && !inTitle) {
            const from = Math.max(0, at - 50);
            snippet = (from > 0 ? "..." : "") + body.slice(from, at + 90).replace(/\s+/g, " ").trim() + (at + 90 < body.length ? "..." : "");
          }
          return { id: r.id, snippet };
        });
        return reply(env, 200, { hits });
      } catch (e) {
        return reply(env, 500, { error: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/policies" && request.method === "GET") {
      // The company policy list. Any signed in tier can read it.
      try {
        await policyTable(env);
        // Policies saved before dates were read get their effective date worked out from the stored text once.
        const pend = await env.DB.prepare("SELECT id, body FROM policies WHERE removed = 0 AND effective IS NULL AND body IS NOT NULL AND body <> ''").all();
        for (const r of pend.results || []) await env.DB.prepare("UPDATE policies SET effective = ? WHERE id = ?").bind(findEffective(r.body), r.id).run();
        const rows = await env.DB.prepare("SELECT id, title, item_id, file_name, needs_ack, updated_at, updated_on, updated_src, version, folder, docx_item_id, docx_name, effective, kind FROM policies WHERE removed = 0 ORDER BY title COLLATE NOCASE").all();
        return reply(env, 200, { policies: (rows.results || []).map(policyOut) });
      } catch (e) {
        return reply(env, 500, { error: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/policies" && request.method === "POST") {
      // Director only. Adds a policy: the title and tick come in the query string, the PDF is the request body.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      // The title comes from the form. When it is left blank the name of the uploaded file is used.
      // The Handbook is the same as Policies but kept in Workplace Culture / Handbook.
      const kind = docKind(url.searchParams.get("kind"));
      const title = txt(url.searchParams.get("title"), 120) || txt(titleFromName(url.searchParams.get("name")), 120);
      if (!title) return reply(env, 400, { error: "A policy needs a title" });
      const bytes = await request.arrayBuffer();
      // Nothing is added without a file, so a card can only ever stand for a file that was saved.
      if (!bytes.byteLength) return reply(env, 400, { error: "Choose a file to upload" });
      if (bytes.byteLength > MAX_UPLOAD) return reply(env, 413, { error: "file too large" });
      const madeOn = "";
      try {
        await policyTable(env);
        const id = crypto.randomUUID();
        const folder = cleanName(title) || id;
        // One document per title. A title that was removed earlier still has its files in SharePoint, so numbering carries on from the last version rather than clashing with them.
        const same = await env.DB.prepare("SELECT removed, version FROM policies WHERE COALESCE(kind, 'policy') = ? AND lower(folder) = lower(?) ORDER BY removed ASC, version DESC").bind(kind, folder).first();
        if (same && !same.removed) return reply(env, 409, { error: `${title} has already been added. Edit it to upload a new version.` });
        // A title added again after it was removed starts back at v1. Its old files are archived first.
        const first = 1;
        if (same) await archiveLeftovers(env, [...(await docHome(env, kind)), folder]);
        const pub = await policyPublish(env, null, folder, first, bytes, request.headers.get("Content-Type"), url.searchParams.get("name"), kind);
        // The last updated date is read from the document. A date typed in by hand wins over it.
        const found = !madeOn && pub.body ? findLastUpdated(pub.body) : null;
        const updOn = madeOn || (found ? found.iso : null);
        const updSrc = madeOn ? "manual" : policySrc(found);
        await env.DB.prepare("INSERT INTO policies (id, title, item_id, file_name, docx_item_id, docx_name, body, effective, needs_ack, updated_by, version, folder, updated_on, updated_src, kind) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").bind(id, title, pub.itemId, pub.fileName, pub.docxId, pub.docxName, pub.body, pub.effective, 1, (user.email || "").toLowerCase(), first, folder, updOn, updSrc, kind).run();
        const row = await env.DB.prepare("SELECT id, title, item_id, file_name, needs_ack, updated_at, updated_on, updated_src, version, folder, docx_item_id, docx_name, effective, kind FROM policies WHERE id = ?").bind(id).first();
        return reply(env, 200, { policy: policyOut(row) });
      } catch (e) {
        console.log("policy add error", String(e && e.message || e));
        return reply(env, e.status || 502, { error: e.say || "Could not add the policy", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    const pm = url.pathname.match(/^\/api\/policies\/([0-9a-fA-F-]{20,40})(\/file|\/text|\/location)?$/);
    if (pm) {
      try {
        await policyTable(env);
        const row = await env.DB.prepare("SELECT id, title, item_id, file_name, needs_ack, removed, updated_at, updated_on, updated_src, version, folder, docx_item_id, docx_name, effective, kind FROM policies WHERE id = ?").bind(pm[1]).first();
        if (!row || row.removed) return reply(env, 404, { error: "policy not found" });
        if (pm[2] === "/file" && request.method === "GET") {
          // Open the policy PDF. Any signed in tier.
          if (!row.item_id) return reply(env, 404, { error: "No file has been uploaded for this policy" });
          const dl = await graphFetch(env, `/sites/${siteId(env)}/drive/items/${row.item_id}/content`, { redirect: "manual" });
          const loc = dl.headers.get("Location");
          const file = loc ? await fetch(loc) : dl;
          if (!file.ok) return reply(env, 502, { error: "download failed" });
          return new Response(file.body, { status: 200, headers: { ...cors(env), "Content-Type": "application/pdf", "Content-Disposition": `inline; filename*=UTF-8''${encodeURIComponent(row.file_name || "policy.pdf")}` } });
        }
        if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
        const who = (user.email || "").toLowerCase();
        if (pm[2] === "/location" && request.method === "GET") {
          // Director only. Asks SharePoint where the current file really is, so the app can show the path and a link rather than take the record's word for it.
          if (!row.item_id) return reply(env, 200, { found: false });
          let it = null;
          try { const r = await graphFetch(env, `/sites/${siteId(env)}/drive/items/${row.item_id}?select=id,name,webUrl,parentReference,file`); if (r.ok) it = await r.json(); } catch (e) { it = null; }
          if (!it || !it.file) return reply(env, 200, { found: false });
          const parent = decodeURIComponent(String(it.parentReference && it.parentReference.path || "").replace(/^.*?root:/, "")).replace(/^\//, "");
          return reply(env, 200, { found: true, path: [...parent.split("/").filter(Boolean), it.name].join(" / "), webUrl: it.webUrl || "" });
        }
        if (pm[2] === "/file" && request.method === "POST") {
          // Upload a new version: the next v number becomes the current file and the previous one moves into the policy's Archive folder.
          const bytes = await request.arrayBuffer();
          if (!bytes.byteLength) return reply(env, 400, { error: "empty file" });
          if (bytes.byteLength > MAX_UPLOAD) return reply(env, 413, { error: "file too large" });
          const folder = row.folder || cleanName(row.title) || row.id;
          const next = (row.item_id ? (row.version || 1) : 0) + 1;
          const pub = await policyPublish(env, row.item_id ? row : null, folder, next, bytes, request.headers.get("Content-Type"), url.searchParams.get("name"), docKind(row.kind));
          const archived = pub.archived;
          // A new version takes its last updated date from its own text. A date set by hand stays. A PDF has no text here, so its date is read by the app straight afterwards.
          const keep = policyDateManual(row);
          const found = !keep && pub.body ? findLastUpdated(pub.body) : null;
          await env.DB.prepare("UPDATE policies SET item_id = ?, file_name = ?, docx_item_id = ?, docx_name = ?, body = ?, effective = ?, version = ?, folder = ?, updated_by = ?, updated_at = datetime('now')" + (keep ? "" : ", updated_on = ?, updated_src = ?") + " WHERE id = ?").bind(...[pub.itemId, pub.fileName, pub.docxId, pub.docxName, pub.body, pub.effective, next, folder, who, ...(keep ? [] : [found ? found.iso : null, policySrc(found)]), row.id]).run();
          const outV = await env.DB.prepare("SELECT id, title, item_id, file_name, needs_ack, updated_at, updated_on, updated_src, version, folder, docx_item_id, docx_name, effective, kind FROM policies WHERE id = ?").bind(row.id).first();
          // A new version wipes everyone's acknowledgement (the app matches on version), so every active staff member except the director is told to read and acknowledge it again.
          if (row.needs_ack) {
            try {
              const isHb = docKind(row.kind) !== "policy";
              const staff = (((await env.DB.prepare("SELECT email, name FROM xero_payroll WHERE email IS NOT NULL AND email <> '' AND (status IS NULL OR status = 'ACTIVE')").all()).results) || []).filter((p) => !/andrew\s+wowk/i.test(String(p.name || "")));
              for (const p of staff) await notify(env, { to: String(p.email).toLowerCase(), kind: "policy", title: `${row.title} has been updated`, body: `A new version of the ${isHb ? "handbook" : "policy"} is out. Open it, read it to the end and acknowledge it again.`, link: isHb ? "handbook" : "policies", by: user.name || who });
            } catch (e) { console.log("policy notice failed", String(e && e.message || e).slice(0, 200)); }
          }
          return reply(env, 200, { policy: policyOut(outV), archived });
        } else if (pm[2] === "/text" && request.method === "POST") {
          return reply(env, 200, await policyTextSave(env, row, request));
        } else if (!pm[2] && request.method === "PUT") {
          let body;
          try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
          const title = txt(body.title, 120);
          if (!title) return reply(env, 400, { error: "A policy needs a title" });
          await env.DB.prepare("UPDATE policies SET title = ?, needs_ack = 1, updated_by = ?, updated_at = datetime('now') WHERE id = ?").bind(title, who, row.id).run();
        } else if (!pm[2] && request.method === "DELETE") {
          await env.DB.prepare("UPDATE policies SET removed = 1, updated_by = ?, updated_at = datetime('now') WHERE id = ?").bind(who, row.id).run();
          return reply(env, 200, { removed: true });
        } else {
          return reply(env, 405, { error: "method not allowed" });
        }
        const out = await env.DB.prepare("SELECT id, title, item_id, file_name, needs_ack, updated_at, updated_on, updated_src, version, folder, docx_item_id, docx_name, effective, kind FROM policies WHERE id = ?").bind(row.id).first();
        return reply(env, 200, { policy: policyOut(out) });
      } catch (e) {
        console.log("policy error", String(e && e.message || e));
        return reply(env, e.status || 502, { error: e.say || "Could not change the policy", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/personalemail" && request.method === "PUT") {
      // Saves the signed in person's own personal email. Any tier, only ever their own record.
      if (!user.email) return reply(env, 400, { error: "No email on your sign in" });
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const value = String(body.personalEmail === undefined || body.personalEmail === null ? "" : body.personalEmail).replace(/\s+/g, "").slice(0, 160);
      if (value && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(value)) return reply(env, 400, { error: "That does not look like an email address" });
      try {
        const row = await reservoirRowFor(env, user.email, user.name, true);
        await env.DB.prepare("UPDATE dfp_reservoir SET personal_email = ?, updated_by = ?, updated_at = datetime('now') WHERE person_key = ?").bind(value, user.email.toLowerCase(), row.person_key).run();
        return reply(env, 200, { saved: true, personalEmail: value });
      } catch (e) {
        console.log("personal email save error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not save the personal email", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/photo" && request.method === "POST") {
      // A photo of a licence or registration card, sent with a change request. Held for the signed in person until the office processes the request. Any tier, own card only.
      const slot = String(url.searchParams.get("slot") || "");
      if (!["driver", "truck", "forklift", "excavator", "plumbReg"].includes(slot)) return reply(env, 400, { error: "Unknown card" });
      const bytes = await request.arrayBuffer();
      if (!bytes.byteLength) return reply(env, 400, { error: "empty photo" });
      if (bytes.byteLength > 1800000) return reply(env, 413, { error: "That photo is too large" });
      try {
        await profileTables(env);
        await env.DB.prepare("CREATE TABLE IF NOT EXISTS profile_photos (id TEXT PRIMARY KEY, person_key TEXT, slot TEXT, name TEXT, mime TEXT, data TEXT, created_at TEXT DEFAULT (datetime('now')))").run();
        const u8 = new Uint8Array(bytes);
        let bin = "";
        for (let i = 0; i < u8.length; i += 8192) bin += String.fromCharCode.apply(null, u8.subarray(i, i + 8192));
        const id = crypto.randomUUID();
        const mime = /^image\/(jpeg|png|webp)$/.test(String(request.headers.get("Content-Type") || "")) ? request.headers.get("Content-Type") : "image/jpeg";
        await env.DB.prepare("INSERT INTO profile_photos (id, person_key, slot, name, mime, data) VALUES (?, ?, ?, ?, ?, ?)").bind(id, profileKey(user.name || user.email), slot, cleanName(url.searchParams.get("name")) || "photo.jpg", mime, btoa(bin)).run();
        return reply(env, 200, { id });
      } catch (e) {
        console.log("profile photo save error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not save the photo" });
      }
    }
    const ppo = url.pathname.match(/^\/api\/profile\/photo\/([0-9a-fA-F-]{20,40})$/);
    if (ppo && request.method === "GET") {
      // Opens one card photo. The person it belongs to, or a director or manager.
      try {
        await env.DB.prepare("CREATE TABLE IF NOT EXISTS profile_photos (id TEXT PRIMARY KEY, person_key TEXT, slot TEXT, name TEXT, mime TEXT, data TEXT, created_at TEXT DEFAULT (datetime('now')))").run();
        const row = await env.DB.prepare("SELECT * FROM profile_photos WHERE id = ?").bind(ppo[1]).first();
        if (!row) return reply(env, 404, { error: "photo not found" });
        if (user.tier !== "T1" && user.tier !== "T2" && row.person_key !== profileKey(user.name || user.email)) return reply(env, 403, { error: "forbidden" });
        const bin = atob(row.data);
        const out = new Uint8Array(bin.length);
        for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
        return new Response(out, { status: 200, headers: { ...cors(env), "Content-Type": row.mime || "image/jpeg", "Cache-Control": "no-store" } });
      } catch (e) {
        console.log("profile photo open error", String(e && e.message || e));
        return reply(env, 502, { error: "Could not open the photo" });
      }
    }
    if (url.pathname === "/api/profile/card" && request.method === "POST") {
      // The signed in person saving one of their own cards. No approval. The office gets a notice.
      if (!user.email && !user.name) return reply(env, 400, { error: "No name on your sign in" });
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const def = CARD_DEFS[String(body.card || "")];
      if (!def || !body.values || typeof body.values !== "object") return reply(env, 400, { error: "That card is not known" });
      try {
        await profileTables(env);
        await env.DB.prepare("CREATE TABLE IF NOT EXISTS profile_photos (id TEXT PRIMARY KEY, person_key TEXT, slot TEXT, name TEXT, mime TEXT, data TEXT, created_at TEXT DEFAULT (datetime('now')))").run();
        const keep = {};
        for (const k of def.keys) if (Object.prototype.hasOwnProperty.call(body.values, k)) keep[k] = body.values[k];
        let newIds = [];
        if (Object.prototype.hasOwnProperty.call(body.values, def.copy)) {
          newIds = String(body.values[def.copy] || "").split("|").filter(Boolean).slice(0, 2);
          for (const pid of newIds) {
            const p = await env.DB.prepare("SELECT person_key FROM profile_photos WHERE id = ?").bind(pid).first();
            if (!p || p.person_key !== profileKey(user.name || user.email)) return reply(env, 400, { error: "That photo was not found. Take it again." });
          }
          keep[def.copy] = newIds.join("|");
        }
        if (!Object.keys(keep).length) return reply(env, 400, { error: "Nothing to save" });
        const before = await reservoirRowFor(env, user.email, user.name, false);
        const oldIds = String((before && before[PROFILE_COLS[def.copy]]) || "").split("|").filter(Boolean);
        const saved = await profileWrite(env, { name: user.name, email: user.email }, keep, user.email || "unknown");
        for (const oid of oldIds) if (!newIds.includes(oid)) await env.DB.prepare("DELETE FROM profile_photos WHERE id = ?").bind(oid).run();
        // A job sent back to the person by the office for this card is now done.
        await noticesTable(env);
        await env.DB.prepare("UPDATE notices SET done_at = datetime('now'), read_at = COALESCE(read_at, datetime('now')), done_by = ? WHERE to_email = ? AND kind = 'cardreset' AND done_at IS NULL AND data LIKE ?").bind(user.name || user.email, String(user.email || "").toLowerCase(), `%"card":"${body.card}"%`).run();
        const bits = [];
        for (const k of def.keys) if (Object.prototype.hasOwnProperty.call(keep, k)) bits.push(/Expiry$/.test(k) ? `Expiry ${dmyOf(saved[k])}` : /Class$/.test(k) ? `Class ${saved[k]}` : `Number ${saved[k]}`);
        if (newIds.length) bits.push(newIds.length > 1 ? "Front and back photos added" : "Front photo added");
        const first = String(user.name || user.email).split(" ")[0];
        await notify(env, { to: "@office", kind: "card", title: `${first} updated their ${def.label}`, body: bits.join(". "), todo: true, link: "", data: { card: body.card, label: def.label, email: String(user.email || "").toLowerCase(), name: user.name || "", photo: saved[def.copy] || "" }, by: user.name || user.email });
        await env.DB.prepare("INSERT INTO profile_audit (event, person, email, section, section_title, fields, by_name, note) VALUES ('card saved', ?, ?, ?, ?, ?, ?, ?)").bind(user.name || "", String(user.email || "").toLowerCase(), String(body.card), def.label, JSON.stringify(Object.keys(keep)), user.name || user.email || "", bits.join(". ").slice(0, 300)).run();
        return reply(env, 200, { saved: true, profile: saved });
      } catch (e) {
        const msg = String(e && e.message || e);
        console.log("profile card save error", msg);
        return reply(env, /^bad /.test(msg) ? 400 : 500, { error: /^bad /.test(msg) ? "One of those values is not valid" : "Could not save the card", detail: msg.slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/card/reset" && request.method === "POST") {
      // The office sending a card back: it is cleared, and the person gets a job to do it again.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const def = CARD_DEFS[String(body.card || "")];
      const target = String(body.email || "").toLowerCase();
      if (!def || !target) return reply(env, 400, { error: "That card or person is not known" });
      try {
        await profileTables(env);
        const row = await reservoirRowFor(env, target, "", false);
        if (!row) return reply(env, 404, { error: "That person has no saved details" });
        const ids = String(row[PROFILE_COLS[def.copy]] || "").split("|").filter(Boolean);
        const blank = {};
        for (const k of [...def.keys, def.copy]) blank[k] = "";
        await profileWrite(env, { name: row.name, email: target }, blank, user.email || "unknown");
        for (const pid of ids) await env.DB.prepare("DELETE FROM profile_photos WHERE id = ?").bind(pid).run();
        const note = String(body.note || "").trim().slice(0, 300);
        await notify(env, { to: target, kind: "cardreset", title: `Your ${def.label} needs doing again`, body: note || "The photo was not clear enough. Please scan it again and take your time.", todo: true, link: "mylicences", data: { card: String(body.card) }, by: user.name || user.email });
        await noticesTable(env);
        if (body.noticeId) await env.DB.prepare("UPDATE notices SET done_at = COALESCE(done_at, datetime('now')), read_at = COALESCE(read_at, datetime('now')), done_by = COALESCE(done_by, ?) WHERE id = ? AND to_email = '@office'").bind(user.name || user.email, String(body.noticeId)).run();
        await env.DB.prepare("INSERT INTO profile_audit (event, person, email, section, section_title, fields, by_name, note) VALUES ('card reset', ?, ?, ?, ?, '[]', ?, ?)").bind(row.name || "", target, String(body.card), def.label, user.name || user.email || "", note).run();
        return reply(env, 200, { reset: true });
      } catch (e) {
        console.log("profile card reset error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not reset the card", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/sizes" && request.method === "PUT") {
      // Uniform sizes are self managed: the signed in person saves their own, no approval. Only polo, jumper and jacket are ever taken.
      if (!user.email) return reply(env, 400, { error: "No email on your sign in" });
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const keep = {};
      for (const k of ["polo", "jumper", "jacket"]) if (Object.prototype.hasOwnProperty.call(body, k)) keep[k] = String(body[k] === null || body[k] === undefined ? "" : body[k]).trim().slice(0, 12);
      if (!Object.keys(keep).length) return reply(env, 400, { error: "No sizes given" });
      try {
        const saved = await profileWrite(env, { name: user.name, email: user.email }, keep, user.email || "unknown");
        return reply(env, 200, { saved: true, profile: saved });
      } catch (e) {
        const msg = String(e && e.message || e);
        return reply(env, /^bad /.test(msg) ? 400 : 500, { error: /^bad /.test(msg) ? "That size is not valid" : "Could not save the sizes", detail: msg.slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile" && request.method === "GET") {
      // The signed in person's own saved profile details. Any tier, only ever their own record.
      try {
        const mine = await profileFor(env, user);
        return reply(env, 200, { found: !!mine, profile: mine });
      } catch (e) {
        console.log("profile read error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not read the profile", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/all" && request.method === "GET") {
      // Everyone's saved profile details, for the Staff screen. Directors and managers only.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      try {
        await profileTables(env);
        const rows = await env.DB.prepare("SELECT * FROM dfp_reservoir WHERE name IS NOT NULL AND name <> '' ORDER BY name COLLATE NOCASE").all();
        return reply(env, 200, { people: (rows.results || []).map(profileOut) });
      } catch (e) {
        console.log("profile list error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not read the profiles", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/person" && request.method === "PUT") {
      // A director or manager editing one person's details directly from the Staff screen.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const name = String(body.name || "").trim();
      if (!profileKey(name) || !body.fields || typeof body.fields !== "object") return reply(env, 400, { error: "A name and fields are needed" });
      try {
        const saved = await profileWrite(env, { name }, body.fields, user.email || "unknown");
        return reply(env, 200, { saved: true, profile: saved });
      } catch (e) {
        const msg = String(e && e.message || e);
        return reply(env, /^bad /.test(msg) ? 400 : 500, { error: /^bad /.test(msg) ? "That value is not valid" : "Could not save the profile", detail: msg.slice(0, 200) });
      }
    }
    if (url.pathname === "/api/dfp/files" && request.method === "GET") {
      // Director only (T1). What is stored, without the file contents.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      await filesTable(env);
      const rows = ((await env.DB.prepare("SELECT slot, name, mime, size, updated_at, updated_by FROM dfp_files").all()).results) || [];
      return reply(env, 200, { slots: Object.keys(FILE_SLOTS).map((k) => { const r = rows.find((x) => x.slot === k); return { slot: k, label: FILE_SLOTS[k].label, stored: !!r, name: r ? r.name : "", size: r ? r.size : 0, updatedAt: r ? r.updated_at : "", updatedBy: r ? r.updated_by : "" }; }), worker: WORKER_BUILD });
    }
    const fs = url.pathname.match(/^\/api\/dfp\/files\/([a-z0-9-]{1,40})$/);
    if (fs && (request.method === "PUT" || request.method === "DELETE")) {
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      const slot = FILE_SLOTS[fs[1]];
      if (!slot) return reply(env, 404, { error: "unknown file slot" });
      await filesTable(env);
      if (request.method === "DELETE") {
        await env.DB.prepare("DELETE FROM dfp_files WHERE slot = ?").bind(fs[1]).run();
        return reply(env, 200, { ok: true });
      }
      let body;
      try { body = await request.json(); } catch (e) { return reply(env, 400, { error: "bad request" }); }
      const m = String(body.data || "").match(/^data:([a-z]+\/[a-z+.-]+);base64,([A-Za-z0-9+\/=]+)$/);
      if (!m || !slot.mimes.includes(m[1])) return reply(env, 400, { error: `That file type is not allowed here. Use ${slot.mimes.map((x) => x.split("/")[1].toUpperCase()).join(" or ")}.` });
      if (String(body.data).length > slot.maxChars) return reply(env, 413, { error: "That file is too large." });
      const name = String(body.name || "").replace(/[^\w .()-]/g, "").slice(0, 80);
      await env.DB.prepare("INSERT OR REPLACE INTO dfp_files (slot, name, mime, size, data, updated_at, updated_by) VALUES (?, ?, ?, ?, ?, datetime('now'), ?)").bind(fs[1], name, m[1], Math.floor(m[2].length * 3 / 4), body.data, String(user.email || user.name || "")).run();
      return reply(env, 200, { ok: true });
    }
    const fr = url.pathname.match(/^\/api\/files\/([a-z0-9-]{1,40})$/);
    if (fr && request.method === "GET") {
      // The stored file for a slot that every signed in person may read (the pay slip logo). Other slots are director only.
      const slot = FILE_SLOTS[fr[1]];
      if (!slot || (!slot.readAll && user.tier !== "T1")) return reply(env, 404, { found: false });
      await filesTable(env);
      const r = await env.DB.prepare("SELECT mime, data FROM dfp_files WHERE slot = ?").bind(fr[1]).first();
      if (!r) return reply(env, 200, { found: false });
      return reply(env, 200, { found: true, mime: r.mime, data: r.data });
    }
    if (url.pathname === "/api/launch/status" && request.method === "GET") {
      // Whether removing requests is still allowed. Directors and managers (the people who see Team Requests).
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      try { const l = await requestsLock(env); return reply(env, 200, { requestsLocked: l.locked, lockedBy: l.by, lockedAt: l.at }); }
      catch (e) { return reply(env, 500, { error: "Could not read the launch settings" }); }
    }
    if (url.pathname === "/api/launch/lock-requests" && request.method === "POST") {
      // The go live switch. Director (T1) only. One way: there is no route that turns it off.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        await launchTable(env);
        await env.DB.prepare("INSERT OR IGNORE INTO app_settings (key, value, by_name) VALUES ('launch.requests_locked', '1', ?)").bind(String(user.name || user.email || "").slice(0, 120)).run();
        const l = await requestsLock(env);
        return reply(env, 200, { requestsLocked: l.locked, lockedBy: l.by, lockedAt: l.at });
      } catch (e) {
        console.log("launch lock error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not set the lock", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/requests" && request.method === "POST") {
      // Any signed in person asking the office to change their own details.
      if (!user.name && !user.email) return reply(env, 400, { error: "No name on your sign in" });
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const section = String(body.section || "").slice(0, 40), sectionTitle = String(body.sectionTitle || "").slice(0, 80);
      if (!section || !body.values || typeof body.values !== "object") return reply(env, 400, { error: "A section and values are needed" });
      try {
        await profileTables(env);
        const id = crypto.randomUUID();
        const fields = Array.isArray(body.fields) ? body.fields.slice(0, 20).map((f) => ({ key: String(f.key || "").slice(0, 40), label: String(f.label || "").slice(0, 80), from: String(f.from === undefined ? "" : f.from).slice(0, 200), to: String(f.to === undefined ? "" : f.to).slice(0, 200) })) : [];
        await env.DB.prepare("INSERT INTO profile_requests (id, person_key, name, email, section, section_title, fields, vals, text, status) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, 'new')")
          .bind(id, profileKey(user.name || user.email), String(user.name || user.email).slice(0, 120), String(user.email || "").toLowerCase(), section, sectionTitle, JSON.stringify(fields), JSON.stringify(body.values).slice(0, 8000), String(body.text || "").slice(0, 600)).run();
        const row = await env.DB.prepare("SELECT * FROM profile_requests WHERE id = ?").bind(id).first();
        await profileAudit(env, "submitted", row, row.name, "");
        return reply(env, 200, { request: profileRequestOut(row) });
      } catch (e) {
        console.log("profile request error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not send the request", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/requests" && request.method === "GET") {
      // Directors and managers see everyone's requests. Everyone else sees only their own.
      try {
        await profileTables(env);
        const office = user.tier === "T1" || user.tier === "T2";
        const rows = office
          ? await env.DB.prepare("SELECT * FROM profile_requests ORDER BY created_at DESC LIMIT 300").all()
          : await env.DB.prepare("SELECT * FROM profile_requests WHERE person_key = ? ORDER BY created_at DESC LIMIT 100").bind(profileKey(user.name || user.email)).all();
        return reply(env, 200, { requests: (rows.results || []).map(profileRequestOut) });
      } catch (e) {
        console.log("profile requests list error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not read the requests", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    const prd = url.pathname.match(/^\/api\/profile\/requests\/([0-9a-fA-F-]{20,40})$/);
    if (prd && request.method === "DELETE") {
      // Removing a closed request from Team Requests. Directors (T1) only, only once it has been processed, and never once the go live lock is on.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        if ((await requestsLock(env)).locked) return reply(env, 403, { error: REMOVAL_LOCKED });
        await profileTables(env);
        const row = await env.DB.prepare("SELECT id, status FROM profile_requests WHERE id = ?").bind(prd[1]).first();
        if (!row) return reply(env, 200, { removed: false, alreadyGone: true });
        if (row.status !== "processed") return reply(env, 409, { error: "Only a processed request can be removed" });
        await env.DB.prepare("DELETE FROM profile_requests WHERE id = ?").bind(row.id).run();
        return reply(env, 200, { removed: true });
      } catch (e) {
        console.log("profile request delete error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not remove the request", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    const prm = url.pathname.match(/^\/api\/profile\/requests\/([0-9a-fA-F-]{20,40})\/process$/);
    if (prm && request.method === "POST") {
      // Approving a request writes the new values into the person's saved profile. Directors and managers only.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      try {
        await profileTables(env);
        const row = await env.DB.prepare("SELECT * FROM profile_requests WHERE id = ?").bind(prm[1]).first();
        if (!row) return reply(env, 404, { error: "not found" });
        if (row.status === "processed") return reply(env, 200, { request: profileRequestOut(row), alreadyDone: true });
        if (row.section === "personal" || row.section === "emergency") {
          // Director only. Preferred name and personal email are saved here. Address and mobile go to Xero first and are read back. The emergency contact is kept in our own table, because Xero Payroll AU has no emergency contact field.
          if (!isApprover(user)) return reply(env, 403, { error: "Only Tayler can process profile changes for now" });
          let vals = {};
          try { vals = JSON.parse(row.vals || "{}"); } catch (e) { /* nothing */ }
          const mail = String(row.email || "").toLowerCase();
          try {
            await dfpTables(env);
            const dx = await env.DB.prepare("SELECT xero_employee_id FROM dfp_xero WHERE lower(email) = ?").bind(mail).first();
            if (row.section === "personal") {
              if (vals.personalEmail !== undefined && vals.personalEmail && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(vals.personalEmail).trim())) throw new Error("The personal email does not look like an email address.");
              if (vals.address !== undefined || vals.mobile !== undefined) {
                const empId = dx && dx.xero_employee_id;
                if (!empId) return reply(env, 409, { error: "This person is not linked to a Xero employee record yet, so nothing was changed." });
                await xeroPushPersonal(env, empId, vals, mail);
              }
              if (vals.preferredName !== undefined) await reservoirOwnSet(env, row.email, row.name, { preferred_name: String(vals.preferredName || "").trim().slice(0, 40) || null }, user.email);
              if (vals.personalEmail !== undefined) {
                const rr = await reservoirRowFor(env, row.email, row.name, true);
                await env.DB.prepare("UPDATE dfp_reservoir SET personal_email = ?, updated_by = ?, updated_at = datetime('now') WHERE person_key = ?").bind(String(vals.personalEmail || "").replace(/\s+/g, "").slice(0, 160), user.email || "unknown", rr.person_key).run();
              }
            } else {
              const clean = (x, n) => String(x === undefined || x === null ? "" : x).replace(/\s+/g, " ").trim().slice(0, n);
              const map = { emName: ["em_name", 120], emRelationship: ["em_relationship", 80], emMobile: ["em_mobile", 40], emEmail: ["em_email", 160] };
              const patch = {};
              for (const [k, [col, n]] of Object.entries(map)) if (vals[k] !== undefined) patch[col] = k === "emMobile" ? normPhone(clean(vals[k], n)) : clean(vals[k], n);
              if (Object.keys(patch).length) await reservoirOwnSet(env, row.email, row.name, { ...patch, em_saved_at: new Date().toISOString() }, user.email);
            }
          } catch (e) {
            const msg = String(e && e.message || e).slice(0, 400);
            await profileAudit(env, "failed", row, user.email || "unknown", msg);
            return reply(env, 502, { error: msg });
          }
          await env.DB.prepare("UPDATE profile_requests SET status = 'processed', processed_by = ?, processed_at = datetime('now'), seen_at = NULL WHERE id = ?").bind(user.email || "unknown", row.id).run();
          const doneRow = await env.DB.prepare("SELECT * FROM profile_requests WHERE id = ?").bind(row.id).first();
          await profileAudit(env, "processed", doneRow, user.name || user.email || "unknown", vals.address !== undefined || vals.mobile !== undefined ? "Written to Xero and read back" : "Saved");
          return reply(env, 200, { request: profileRequestOut(doneRow), xero: vals.address !== undefined || vals.mobile !== undefined });
        }
        if (row.section === "bank" || row.section === "super") {
          // Bank and super go to Xero first. Only a director can press this, and nothing is marked processed unless Xero took the change.
          if (!isApprover(user)) return reply(env, 403, { error: "Only Tayler can process bank and super changes for now" });
          let vals = {};
          try { vals = JSON.parse(row.vals || "{}"); } catch (e) { /* nothing */ }
          const dx = await env.DB.prepare("SELECT xero_employee_id FROM dfp_xero WHERE email = ?").bind(String(row.email || "").toLowerCase()).first();
          const empId = dx && dx.xero_employee_id;
          if (!empId) return reply(env, 409, { error: "This person is not linked to a Xero employee record yet, so nothing was changed." });
          let pushed;
          try { pushed = row.section === "bank" ? await xeroPushBank(env, empId, vals) : await xeroPullSuper(env, empId, vals); }
          catch (e) {
            const msg = String(e && e.message || e).slice(0, 400);
            await profileAudit(env, "failed", row, user.email || "unknown", msg);
            return reply(env, 502, { error: row.section === "super" ? msg : `Xero did not take the change. ${msg}` });
          }
          if (row.section === "bank") {
            await env.DB.prepare("UPDATE dfp_xero SET bank_account_name = ?, bsb = ?, account_number = ?, updated_by = ?, updated_at = datetime('now') WHERE email = ?").bind(pushed.accountName, pushed.bsb, pushed.accountNumber, user.email || "unknown", String(row.email || "").toLowerCase()).run();
          } else {
            const sets = ["super_member_number = ?"], binds = [pushed.member];
            if (pushed.fundName) { sets.push("super_fund_name = ?"); binds.push(pushed.fundName); }
            if (pushed.usi) { sets.push("super_usi = ?"); binds.push(pushed.usi); }
            await env.DB.prepare(`UPDATE xero_payroll SET ${sets.join(", ")} WHERE employee_id = ?`).bind(...binds, empId).run();
          }
          await env.DB.prepare("UPDATE profile_requests SET status = 'processed', processed_by = ?, processed_at = datetime('now'), seen_at = NULL WHERE id = ?").bind(user.email || "unknown", row.id).run();
          const doneRow = await env.DB.prepare("SELECT * FROM profile_requests WHERE id = ?").bind(row.id).first();
          await profileAudit(env, "processed", doneRow, user.name || user.email || "unknown", row.section === "super" ? "Pulled from Xero" : "Written to Xero");
          return reply(env, 200, { request: profileRequestOut(doneRow), xero: true });
        }
        let values = {};
        try { values = JSON.parse(row.vals || "{}"); } catch (e) { /* nothing to apply */ }
        const keep = {};
        for (const k of Object.keys(PROFILE_COLS)) if (Object.prototype.hasOwnProperty.call(values, k)) keep[k] = values[k];
        if (Object.prototype.hasOwnProperty.call(values, "training")) keep.training = values.training;
        let saved = null;
        if (Object.keys(keep).length) saved = await profileWrite(env, { name: row.name, email: row.email }, keep, user.email || "unknown");
        await env.DB.prepare("UPDATE profile_requests SET status = 'processed', processed_by = ?, processed_at = datetime('now') WHERE id = ?").bind(user.email || "unknown", row.id).run();
        const done = await env.DB.prepare("SELECT * FROM profile_requests WHERE id = ?").bind(row.id).first();
        return reply(env, 200, { request: profileRequestOut(done), profile: saved, durable: Object.keys(keep) });
      } catch (e) {
        const msg = String(e && e.message || e);
        console.log("profile process error", msg);
        return reply(env, /^bad /.test(msg) ? 400 : 500, { error: /^bad /.test(msg) ? "That request holds a value that is not valid" : "Could not apply the request", detail: msg.slice(0, 200) });
      }
    }
    const psn = url.pathname.match(/^\/api\/profile\/requests\/([0-9a-fA-F-]{20,40})\/seen$/);
    if (psn && request.method === "POST") {
      // The person has now seen the "details updated" banner. It will not show again.
      try {
        await profileTables(env);
        await env.DB.prepare("UPDATE profile_requests SET seen_at = datetime('now') WHERE id = ? AND person_key = ? AND seen_at IS NULL").bind(psn[1], profileKey(user.name || user.email)).run();
        return reply(env, 200, { ok: true });
      } catch (e) {
        return reply(env, 500, { error: "Could not record that", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/profile/audit" && request.method === "GET") {
      // Director only. The permanent record of every change request, newest first.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        await profileTables(env);
        const rows = await env.DB.prepare("SELECT a.*, r.fields AS req_fields FROM profile_audit a LEFT JOIN profile_requests r ON r.id = a.request_id ORDER BY a.id DESC LIMIT 300").all();
        return reply(env, 200, { entries: (rows.results || []).map((r) => { let fields = []; try { fields = JSON.parse(r.req_fields || r.fields || "[]"); } catch (e) { try { fields = JSON.parse(r.fields || "[]"); } catch (e2) { /* none */ } } return { id: r.id, requestId: r.request_id, at: r.at, event: r.event, person: r.person, section: r.section, sectionTitle: r.section_title, fields, by: r.by_name, note: r.note || "" }; }) });
      } catch (e) {
        return reply(env, 500, { error: "Could not read the change record", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/access/tiers" && request.method === "GET") {
      // Director only (T1). Each person's real access level from Entra, keyed by email. Used by Login As and View As so they start at the person's true level.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        return reply(env, 200, { ok: true, tiers: await accessTiers(env), worker: WORKER_BUILD });
      } catch (e) {
        return reply(env, 200, { ok: false, error: String(e && e.message || e).slice(0, 400), worker: WORKER_BUILD });
      }
    }
    if (url.pathname === "/api/xero/me" && request.method === "GET") {
      // The signed in person's own record, read from DFP's own tables (dfp_xero, with the personal email from dfp_reservoir). Nothing is read from Xero. Any tier, but only ever their own details.
      try {
        const email = String(user.email || "").trim().toLowerCase();
        if (!email) return reply(env, 200, { found: false, reason: "No email on your sign in" });
        await reservoirTables(env);
        const row = await env.DB.prepare(`${PERSON_SQL} WHERE x.email = ?`).bind(email).first();
        if (!row || !row.name) return reply(env, 200, { found: false, reason: NOT_IN_RESERVOIR });
        const out = dfpPersonOut(row);
        delete out.id; delete out.status;
        const pe = await personalEmailFor(env, email);
        if (pe) Object.assign(out, pe);
        return reply(env, 200, { found: true, employee: out });
      } catch (e) {
        console.log("me error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/xero/payslips" && request.method === "GET") {
      // The signed in person's own pay slips, read from the stored copy in Reservoir's database. Nothing is read from Xero. Any tier, only their own.
      try {
        const r = await payslipListFor(env, String(user.email || "").trim().toLowerCase());
        return reply(env, 200, r);
      } catch (e) {
        console.log("payslips error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    const xs = url.pathname.match(/^\/api\/xero\/payslips\/([0-9a-fA-F-]{20,40})$/);
    if (xs && request.method === "GET") {
      // One stored pay slip. Year to date is added up from the stored slips, and the person's own bank details are added from dfp_xero. Nothing is read from Xero. Any tier, only their own.
      try {
        const email = String(user.email || "").trim().toLowerCase();
        if (!email) return reply(env, 404, { error: "No email on your sign in" });
        const out = await payslipOneFor(env, email, xs[1]);
        if (!out) return reply(env, 404, { error: "That pay slip is not in Reservoir yet" });
        return reply(env, 200, { payslip: out });
      } catch (e) {
        console.log("payslip error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    // Director only (T1): another person's stored pay slips, for the Team Pay Slips screen. The person is named by their Xero employee id. Read only, nothing is read from Xero.
    const tf = url.pathname.match(/^\/api\/dfp\/payslips\/for\/([0-9a-fA-F-]{20,40})(?:\/([0-9a-fA-F-]{20,40}))?$/);
    if (tf && request.method === "GET") {
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        await dfpTables(env);
        const who = await env.DB.prepare("SELECT email FROM xero_payroll WHERE employee_id = ?").bind(tf[1]).first();
        const email = who && who.email ? String(who.email).trim().toLowerCase() : "";
        if (!email) return reply(env, 200, tf[2] ? { error: "No email on record for that person" } : { found: false, reason: "That person has no work email in Reservoir yet" });
        if (!tf[2]) return reply(env, 200, await payslipListFor(env, email));
        const out = await payslipOneFor(env, email, tf[2]);
        if (!out) return reply(env, 404, { error: "That pay slip is not in Reservoir yet" });
        return reply(env, 200, { payslip: out });
      } catch (e) {
        console.log("team payslip error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/dfp/payslips/pull" && request.method === "POST") {
      // Director only (T1). Reads the next few pay runs from Xero and stores the pay slips of current employees. Run it again until remaining is 0. Run it weekly to add new pay runs.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        let body = {};
        try { body = await request.json(); } catch (e) { body = {}; }
        const r = await payslipPull(env, { runs: body && body.runs, replace: !!(body && body.replace) });
        if (r.fail) return reply(env, r.fail.status || 502, { error: r.fail.error, detail: r.fail.detail || "", worker: WORKER_BUILD });
        return reply(env, 200, { worker: WORKER_BUILD, ...r });
      } catch (e) {
        console.log("payslip pull error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/dfp/payslips/filed" && request.method === "POST") {
      // Director only (T1). Marks the most recent pay run that is not yet filed as filed, so its slips move out of Next Pay and read Paid. Changes no pay figure.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        await payslipTables(env);
        const run = await env.DB.prepare("SELECT run_id, payment_date FROM payslip_runs WHERE filed = 0 ORDER BY payment_date DESC LIMIT 1").first();
        if (!run) return reply(env, 200, { filed: false, reason: "There is no pay run waiting to be filed.", worker: WORKER_BUILD });
        await env.DB.prepare("UPDATE payslip_runs SET filed = 1, filed_at = datetime('now'), filed_by = ? WHERE run_id = ?").bind(user.name || user.email || "", run.run_id).run();
        // The one pull after a pay run is filed: the timesheets Xero now holds as processed are banked here as permanent history. A processed timesheet already stored is never overwritten again. A failure here does not undo the filing.
        let timesheets = { ok: false };
        try { const ts = await timesheetSync(env); timesheets = { ok: !!ts.ok, saved: ts.saved || 0, error: ts.ok ? "" : String(ts.error || "").slice(0, 200) }; } catch (e) { timesheets = { ok: false, error: String(e && e.message || e).slice(0, 200) }; }
        return reply(env, 200, { filed: true, paymentDate: run.payment_date, timesheets, worker: WORKER_BUILD });
      } catch (e) {
        console.log("payslip filed error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/dfp/payslips/status" && request.method === "GET") {
      // Director only (T1). What is stored, from the database alone. No Xero call.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      await payslipTables(env);
      const runs = await env.DB.prepare("SELECT COUNT(*) c, MIN(payment_date) a, MAX(payment_date) b FROM payslip_runs WHERE rules >= ?").bind(PAYSLIP_RULES).first();
      const slips = await env.DB.prepare("SELECT COUNT(*) c, COALESCE(SUM(missing),0) m, COALESCE(SUM(CASE WHEN rules >= ? THEN 1 ELSE 0 END),0) r FROM payslip_store").bind(PAYSLIP_RULES).first();
      const per = ((await env.DB.prepare("SELECT name, COUNT(*) c, MIN(payment_date) a, MAX(payment_date) b FROM payslip_store GROUP BY employee_id ORDER BY name").all()).results) || [];
      const waiting = ((await env.DB.prepare("SELECT r.payment_date, r.built_at, (SELECT COUNT(*) FROM payslip_store s WHERE s.run_id = r.run_id) AS slips FROM payslip_runs r WHERE r.filed = 0 ORDER BY r.payment_date DESC").all()).results) || [];
      return reply(env, 200, { worker: WORKER_BUILD, unfiled: waiting.map((x) => x.payment_date), unfiledRuns: waiting.map((x) => ({ paymentDate: x.payment_date, builtAt: x.built_at, slips: x.slips })), from: PAYSLIP_FROM, runsStored: runs.c, firstPaid: runs.a, lastPaid: runs.b, slipsStored: slips.c, slipsWithGaps: slips.m, slipsOnCurrentRules: slips.r, people: per.map((x) => ({ name: x.name, slips: x.c, first: x.a, last: x.b })) });
    }
    if (url.pathname === "/api/xero/pay" && request.method === "GET") {
      // The signed in person's own pay details, read from xero_payroll (the snapshot the one time fill and the Saturday sweep keep). Nothing is read from Xero. Any tier, only their own record.
      try {
        const email = String(user.email || "").trim().toLowerCase();
        if (!email) return reply(env, 200, { found: false, reason: "No email on your sign in" });
        await dfpTables(env);
        const row = await env.DB.prepare("SELECT * FROM xero_payroll WHERE email = ?").bind(email).first();
        if (!row) return reply(env, 200, { found: false, reason: NOT_IN_RESERVOIR });
        const dm = await leaveDayMap(env);
        return reply(env, 200, { ...dfpPayOut(row), dayHours: leaveDayFrom(dm, email, row.name || user.name) });
      } catch (e) {
        console.log("pay error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/xero/staffpay" && request.method === "GET") {
      // The pay details of one employee, chosen by Xero employee id, read from xero_payroll. Director only (T1).
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      const id = (url.searchParams.get("id") || "").trim();
      if (!/^[0-9a-fA-F-]{20,40}$/.test(id)) return reply(env, 400, { error: "bad employee id" });
      try {
        await dfpTables(env);
        const row = await env.DB.prepare("SELECT * FROM xero_payroll WHERE employee_id = ?").bind(id).first();
        if (!row) return reply(env, 200, { found: false, reason: "No pay details saved for that person yet" });
        return reply(env, 200, dfpPayOut(row));
      } catch (e) {
        console.log("staffpay error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/xero/connect" && request.method === "GET") {
      // Returns the Xero consent address. Director only (T1).
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      if (!env.XERO_CLIENT_SECRET) return reply(env, 503, { error: "xero not set up" });
      const q = new URLSearchParams({ response_type: "code", client_id: xeroId(env), redirect_uri: xeroRedirect(env), scope: env.XERO_SCOPES || XERO_SCOPES_DEFAULT, state: await xeroState(env, user.name) });
      return reply(env, 200, { url: `https://login.xero.com/identity/connect/authorize?${q}` });
    }
    if (url.pathname === "/api/xero/status" && request.method === "GET") {
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      await xeroTable(env);
      const row = await env.DB.prepare("SELECT tenant_name, connected_by, updated_at FROM xero_auth WHERE id = 1").first();
      let calls;
      try { calls = await xeroCallsToday(env); } catch (e) { calls = null; }
      return reply(env, 200, row ? { connected: true, organisation: row.tenant_name, connectedBy: row.connected_by, updated: row.updated_at, calls } : { connected: false, calls });
    }
    if (url.pathname === "/api/xero/staff" && request.method === "GET") {
      // Every person, read from DFP's own tables (dfp_xero, with personal email from dfp_reservoir). Nothing is read from Xero. Director only (T1).
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        await reservoirTables(env);
        const people = ((await env.DB.prepare(`${PERSON_SQL} WHERE x.name IS NOT NULL AND x.name <> ''`).all()).results) || [];
        const pes = ((await env.DB.prepare("SELECT email, personal_email FROM dfp_reservoir WHERE email IS NOT NULL AND personal_email IS NOT NULL AND personal_email <> ''").all()).results) || [];
        const peBy = new Map(pes.map((x) => [String(x.email).toLowerCase(), x.personal_email]));
        const rows = people.map((r) => { const o = dfpPersonOut(r); const pe = peBy.get(String(r.email || "").toLowerCase()); if (pe) o.personalEmail = pe; return o; });
        rows.sort((a, b) => (a.status === "ACTIVE" ? 0 : 1) - (b.status === "ACTIVE" ? 0 : 1) || a.name.localeCompare(b.name));
        return reply(env, 200, { employees: rows });
      } catch (e) {
        console.log("staff error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/staffdocs/me" && request.method === "GET") {
      // The signed in person's own employment documents. Any tier, only ever their own.
      if (!user.email) return reply(env, 400, { error: "No email on your sign in" });
      try {
        await staffDocsTable(env);
        const rows = await env.DB.prepare(`SELECT ${STAFF_DOC_COLS} FROM staff_docs WHERE person_email = ? AND removed = 0 ORDER BY uploaded_at DESC, rowid DESC`).bind(user.email.toLowerCase()).all();
        return reply(env, 200, { docs: (rows.results || []).map(staffDocOut) });
      } catch (e) {
        return reply(env, 500, { error: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/names" && request.method === "GET") {
      // Anyone signed in. The people who have a preferred name, so the app can show it in place of the legal first name on screens that are not official documents.
      try {
        await reservoirTables(env);
        const rows = ((await env.DB.prepare("SELECT x.name AS name, r.preferred_name AS preferred_name FROM dfp_xero x JOIN dfp_reservoir r ON lower(r.email) = lower(x.email) WHERE r.preferred_name IS NOT NULL AND TRIM(r.preferred_name) <> '' AND x.name IS NOT NULL").all()).results) || [];
        return reply(env, 200, { names: rows.map((r) => ({ name: r.name, preferred: String(r.preferred_name).trim() })) });
      } catch (e) {
        console.log("names error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not read the preferred names" });
      }
    }
    if (url.pathname === "/api/xero/staff/birthtime" && request.method === "POST") {
      // Sets (or clears) the time of day a person was born, as 24 hour "HH:MM", matched on work email. The director can set anyone's, everyone else only their own.
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const email = txt(body.email, 200).toLowerCase();
      if (user.tier !== "T1" && (!email || email !== String(user.email || "").toLowerCase())) return reply(env, 403, { error: "forbidden" });
      const when = txt(body.birthTime, 5);
      if (!email) return reply(env, 400, { error: "No person given" });
      if (when && !/^([01]\d|2[0-3]):[0-5]\d$/.test(when)) return reply(env, 400, { error: "That is not a time of day" });
      try {
        await reservoirTables(env);
        const row = await env.DB.prepare("SELECT email, name FROM dfp_xero WHERE lower(email) = ?").bind(email).first();
        if (!row) return reply(env, 404, { error: "That person is not in the staff list" });
        await reservoirOwnSet(env, row.email, row.name, { birth_time: when || null }, (user.email || "").toLowerCase());
        return reply(env, 200, { ok: true, birthTime: when });
      } catch (e) {
        console.log("birth time error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not save the time of birth" });
      }
    }
    if (url.pathname === "/api/xero/staff/preferred" && request.method === "POST") {
      // Sets (or clears) one person's preferred name, matched on work email. The director can set anyone's, everyone else only their own. Never creates a person.
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const email = txt(body.email, 200).toLowerCase();
      if (user.tier !== "T1" && (!email || email !== String(user.email || "").toLowerCase())) return reply(env, 403, { error: "forbidden" });
      const preferred = txt(body.preferredName, 40).trim();
      if (!email) return reply(env, 400, { error: "No person given" });
      try {
        await reservoirTables(env);
        const row = await env.DB.prepare("SELECT email, name FROM dfp_xero WHERE lower(email) = ?").bind(email).first();
        if (!row) return reply(env, 404, { error: "That person is not in the staff list" });
        await reservoirOwnSet(env, row.email, row.name, { preferred_name: preferred || null }, (user.email || "").toLowerCase());
        return reply(env, 200, { ok: true, preferredName: preferred });
      } catch (e) {
        console.log("preferred name error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not save the preferred name" });
      }
    }
    if (url.pathname === "/api/dfp/names" && request.method === "POST") {
      // Director only. Brings each person's legal name from Xero into Reservoir's own tables, for people already there. One Xero call (the employee list). Matched on the Xero employee ID, then on work email. Nobody is ever added.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        await dfpTables(env);
        await reservoirTables(env);
        const res = await xeroFetch(env, "payroll.xro/1.0/Employees");
        if (!res.ok) return reply(env, 502, { error: "Xero refused the request", status: res.status, detail: await xeroWhy(res) });
        const list = (await res.json()).Employees || [];
        const out = { checked: 0, changed: [], middleSet: [], unchanged: 0, notInReservoir: [], worker: WORKER_BUILD };
        const sameLetters = (n) => String(n || "").toLowerCase().replace(/[^a-z0-9]/g, "");
        const everyone = ((await env.DB.prepare("SELECT email, name, middle_names FROM dfp_xero").all()).results) || [];
        for (const e of list) {
          const nameNow = [e.FirstName, e.LastName].map((x) => String(x || "").trim()).filter(Boolean).join(" ");
          if (!nameNow) continue;
          out.checked++;
          const mail = String(e.Email || "").trim().toLowerCase();
          let row = await env.DB.prepare("SELECT email, name, middle_names FROM dfp_xero WHERE xero_employee_id = ?").bind(e.EmployeeID).first();
          if (!row && mail) row = await env.DB.prepare("SELECT email, name, middle_names FROM dfp_xero WHERE lower(email) = ?").bind(mail).first();
          // Last resort: the same name once punctuation is ignored (Rochow-OToole and Rochow-O'Toole), so a corrected spelling still finds the person.
          if (!row) { const key = sameLetters(nameNow); const hits = everyone.filter((x) => x.name && sameLetters(x.name) === key); if (hits.length === 1) row = hits[0]; }
          if (!row) { out.notInReservoir.push(nameNow); continue; }
          // Middle names. The employee list may leave them out, so ask for that one person when it does. Never cleared on a failed read.
          let mid = e.MiddleNames;
          if (mid === undefined && e.EmployeeID) {
            try { const one = await xeroFetch(env, `payroll.xro/1.0/Employees/${e.EmployeeID}`); if (one.ok) { const rec = ((await one.json()).Employees || [])[0]; if (rec && rec.MiddleNames !== undefined) mid = rec.MiddleNames; else if (rec) mid = ""; } } catch (e3) { /* leave as it was */ }
          }
          if (mid !== undefined) {
            const midNow = String(mid || "").trim().replace(/\s+/g, " ");
            if (midNow !== String(row.middle_names || "").trim()) {
              await env.DB.prepare("UPDATE dfp_xero SET middle_names = ?, updated_at = datetime('now'), updated_by = ? WHERE email = ?").bind(midNow || null, (user.email || "").toLowerCase(), row.email).run();
              out.middleSet.push({ name: row.name || nameNow, middle: midNow });
            }
          }
          if (row.name === nameNow) { out.unchanged++; continue; }
          await env.DB.prepare("UPDATE dfp_xero SET name = ?, xero_employee_id = COALESCE(xero_employee_id, ?), updated_at = datetime('now'), updated_by = ? WHERE email = ?").bind(nameNow, e.EmployeeID, (user.email || "").toLowerCase(), row.email).run();
          await env.DB.prepare("UPDATE xero_payroll SET name = ? WHERE employee_id = ?").bind(nameNow, e.EmployeeID).run();
          for (const t of ["xero_timesheets", "xero_leave"]) { try { await env.DB.prepare(`UPDATE ${t} SET name = ? WHERE employee_id = ?`).bind(nameNow, e.EmployeeID).run(); } catch (e2) { /* that table is not made yet */ } }
          await env.DB.prepare("UPDATE dfp_reservoir SET name = ? WHERE lower(email) = ?").bind(nameNow, String(row.email).toLowerCase()).run();
          out.changed.push({ from: row.name || "", to: nameNow });
        }
        return reply(env, 200, out);
      } catch (e) {
        console.log("names refresh error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/staffdocs/scan" && request.method === "POST") {
      // Director only. Reads a contract PDF for the person's name, the commencement date and the signed date. The text itself is never sent back.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      const email = txt(url.searchParams.get("email"), 200).toLowerCase();
      if (!email) return reply(env, 400, { error: "No person given" });
      try {
        const bytes = await request.arrayBuffer();
        if (!bytes.byteLength) return reply(env, 400, { error: "empty file" });
        if (bytes.byteLength > MAX_UPLOAD) return reply(env, 413, { error: "file too large" });
        await dfpTables(env);
        const person = await env.DB.prepare("SELECT name FROM dfp_xero WHERE lower(email) = ?").bind(email).first();
        const head = new Uint8Array(bytes.slice(0, 1024));
        const isPdf = Array.from(head.slice(0, 1024)).map((c) => String.fromCharCode(c)).join("").includes("%PDF");
        if (!isPdf) return reply(env, 200, { readable: false, nameFound: false, commenced: "", signed: "", why: "Only PDF files can be read. Enter the dates by hand." });
        let facts;
        try { const got = await pdfText(new Uint8Array(bytes));
          facts = pdfContractFacts(got.text, person && person.name || ""); }
        catch (e) { console.log("contract scan error", String(e && e.message || e).slice(0, 200)); facts = { readable: false, nameFound: false, commenced: "", signed: "" }; }
        return reply(env, 200, { readable: !!facts.readable, nameFound: !!facts.nameFound, commenced: facts.commenced || "", signed: facts.signed || "" });
      } catch (e) {
        console.log("staff doc scan error", String(e && e.message || e));
        return reply(env, 502, { error: "Could not read the file" });
      }
    }
    if (url.pathname === "/api/staffdocs" && (request.method === "GET" || request.method === "POST")) {
      // Director only. GET lists one person's documents. POST files a document: ?email=&slot=&name=, the file is the request body.
      // Contract: &signed=YYYY-MM-DD&commenced=YYYY-MM-DD. Additional: &kind=cv|details|tfn|super|other&label=&docDate=YYYY-MM-DD.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      const email = txt(url.searchParams.get("email"), 200).toLowerCase();
      if (!email) return reply(env, 400, { error: "No person given" });
      try {
        await staffDocsTable(env);
        if (request.method === "GET") {
          const rows = await env.DB.prepare(`SELECT ${STAFF_DOC_COLS} FROM staff_docs WHERE person_email = ? AND removed = 0 ORDER BY uploaded_at DESC, rowid DESC`).bind(email).all();
          return reply(env, 200, { docs: (rows.results || []).map(staffDocOut) });
        }
        const slot = String(url.searchParams.get("slot") || "");
        if (!STAFF_DOC_SLOTS[slot]) return reply(env, 400, { error: "Unknown document type" });
        const signedRaw = url.searchParams.get("signed"), commRaw = url.searchParams.get("commenced"), dateRaw = url.searchParams.get("docDate");
        const signed = isoDay(signedRaw), commenced = isoDay(commRaw), docDate = isoDay(dateRaw);
        if (signedRaw && !signed) return reply(env, 400, { error: "The signed date is not a valid date" });
        if (commRaw && !commenced) return reply(env, 400, { error: "The commencement date is not a valid date" });
        if (dateRaw && !docDate) return reply(env, 400, { error: "The document date is not a valid date" });
        let title = slot === "contract" ? "Employment Contract" : STAFF_DOC_SLOTS[slot];
        if (slot === "additional") {
          const kind = String(url.searchParams.get("kind") || "other");
          title = kind === "other" ? cleanName(url.searchParams.get("label")) : STAFF_EXTRA_TYPES[kind];
          if (!title) return reply(env, 400, { error: kind === "other" ? "Name the document" : "Unknown document type" });
        } else if (slot === "tfn" || slot === "super") title = STAFF_EXTRA_TYPES[slot];
        const bytes = await request.arrayBuffer();
        if (!bytes.byteLength) return reply(env, 400, { error: "empty file" });
        if (bytes.byteLength > MAX_UPLOAD) return reply(env, 413, { error: "file too large" });
        await reservoirTables(env);
        const person = await env.DB.prepare("SELECT x.name AS name, r.preferred_name AS preferred_name FROM dfp_xero x LEFT JOIN dfp_reservoir r ON lower(r.email) = lower(x.email) WHERE lower(x.email) = ?").bind(email).first();
        if (!person || !person.name) return reply(env, 404, { error: "That person is not in the staff list" });
        const parts = String(person.name).trim().split(/\s+/);
        const first = parts[0], last = parts.slice(1).join(" ");
        const sr = await staffRoot(env);
        if (!sr.root) return reply(env, 502, { error: "The staff folders could not be found", detail: sr.why || "" });
        const folders = await driveChildren(env, sr.root.driveBase, sr.root.id);
        const known = String(person.preferred_name || "").trim();
        let mf = matchStaffFolder(folders, first, last);
        if (!mf.folder && known) mf = matchStaffFolder(folders, known, last);
        if (!mf.folder) return reply(env, 404, { error: "No staff folder was found for that person", detail: mf.why || "" });
        // 20261019 Employment Contract - Chad.pdf (the signed date, then the title, a spaced dash, the first name) and CV - Chad.pdf for the extras.
        const ext = fileExt(url.searchParams.get("name"));
        const callName = known || first;
        const stem = slot === "contract" ? `${signed ? signed.replace(/-/g, "") + " " : ""}${title} - ${callName}` : `${title} - ${callName}`;
        const fileName = cleanName(stem + ext) || "document" + ext;
        const mime = request.headers.get("Content-Type") || "application/octet-stream";
        const item = await uploadUnderItem(env, sr.root.driveBase, mf.folder.id, ["Employment Documents"], fileName, bytes, mime);
        // One file for the contract (and the old fixed slots): a new upload replaces the last. Additional documents add up.
        if (slot !== "additional") await env.DB.prepare("UPDATE staff_docs SET removed = 1 WHERE person_email = ? AND slot = ? AND removed = 0").bind(email, slot).run();
        const id = crypto.randomUUID();
        const filedAt = melbourneDay() + " " + new Intl.DateTimeFormat("en-GB", { timeZone: "Australia/Melbourne", hour: "2-digit", minute: "2-digit", second: "2-digit", hour12: false }).format(new Date()).replace(/^24/, "00");
        const shownDate = slot === "contract" ? signed : docDate;
        await env.DB.prepare("INSERT INTO staff_docs (id, person_email, slot, drive_base, item_id, file_name, mime, uploaded_by, uploaded_at, doc_date, label) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)").bind(id, email, slot, sr.root.driveBase, item.id, item.name || fileName, mime, (user.email || "").toLowerCase(), filedAt, shownDate || null, slot === "additional" ? title : null).run();
        // The commencement date goes on the Reservoir profile only. Xero is not called.
        let commencedSaved = false;
        if (slot === "contract" && commenced) {
          await env.DB.prepare("UPDATE dfp_xero SET start_date = ?, updated_at = datetime('now'), updated_by = ? WHERE lower(email) = ?").bind(commenced, (user.email || "").toLowerCase(), email).run();
          commencedSaved = true;
        }
        const row = await env.DB.prepare(`SELECT ${STAFF_DOC_COLS} FROM staff_docs WHERE id = ?`).bind(id).first();
        return reply(env, 200, { doc: staffDocOut(row), commencedSaved });
      } catch (e) {
        console.log("staff doc error", String(e && e.message || e));
        return reply(env, e.status || 502, { error: e.say || "Could not file the document", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    const sdel = url.pathname.match(/^\/api\/staffdocs\/([0-9a-fA-F-]{20,40})$/);
    if (sdel && request.method === "DELETE") {
      // Director only. Takes a document off the person's record and removes its file from the staff folder (SharePoint keeps it in the recycle bin).
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        await staffDocsTable(env);
        const row = await env.DB.prepare("SELECT id, drive_base, item_id FROM staff_docs WHERE id = ? AND removed = 0").bind(sdel[1]).first();
        if (!row) return reply(env, 404, { error: "document not found" });
        let fileRemoved = false;
        try {
          const del = await graphFetch(env, `${row.drive_base}/items/${row.item_id}`, { method: "DELETE" });
          fileRemoved = del.ok || del.status === 404;
        } catch (e) { console.log("staff doc file delete error", String(e && e.message || e).slice(0, 200)); }
        await env.DB.prepare("UPDATE staff_docs SET removed = 1 WHERE id = ?").bind(row.id).run();
        return reply(env, 200, { ok: true, fileRemoved });
      } catch (e) {
        console.log("staff doc delete error", String(e && e.message || e));
        return reply(env, 502, { error: "Could not delete the document" });
      }
    }
    const sdm = url.pathname.match(/^\/api\/staffdocs\/([0-9a-fA-F-]{20,40})\/file$/);
    if (sdm && request.method === "GET") {
      // Opens one document. The person it belongs to, or the director.
      try {
        await staffDocsTable(env);
        const row = await env.DB.prepare("SELECT id, person_email, drive_base, item_id, file_name, mime FROM staff_docs WHERE id = ? AND removed = 0").bind(sdm[1]).first();
        if (!row) return reply(env, 404, { error: "document not found" });
        if (user.tier !== "T1" && String(user.email || "").toLowerCase() !== String(row.person_email).toLowerCase()) return reply(env, 403, { error: "forbidden" });
        const dl = await graphFetch(env, `${row.drive_base}/items/${row.item_id}/content`, { redirect: "manual" });
        const loc = dl.headers.get("Location");
        const file = loc ? await fetch(loc) : dl;
        if (!file.ok) return reply(env, 502, { error: "download failed" });
        return new Response(file.body, { status: 200, headers: { ...cors(env), "Content-Type": row.mime || "application/octet-stream", "Content-Disposition": `inline; filename*=UTF-8''${encodeURIComponent(row.file_name || "document")}`, "Cache-Control": "no-store" } });
      } catch (e) {
        console.log("staff doc open error", String(e && e.message || e));
        return reply(env, 502, { error: "Could not open the document" });
      }
    }
    if (url.pathname === "/api/ppe/me" && request.method === "GET") {
      // The signed in person's own PPE bag and orders. Any tier.
      try { return reply(env, 200, { items: PPE_ITEMS, ...(await ppeMe(env, user)) }); }
      catch (e) { console.log("ppe read error", String(e && e.message || e)); return reply(env, 500, { error: "Could not read the PPE details" }); }
    }
    if (url.pathname === "/api/ppe/ack" && request.method === "POST") {
      // The bag acknowledgement. The person ticks what is physically in the bag. Anything unticked that DFP supplies becomes an order to the office.
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const ticked = new Set(Array.isArray(body.ticked) ? body.ticked.map(String) : []);
      const sizes = body.sizes && typeof body.sizes === "object" ? body.sizes : {};
      const expiry = String(body.firstAidExpiry || "").trim();
      if (expiry && !/^\d{4}-\d{2}-\d{2}$/.test(expiry)) return reply(env, 400, { error: "That expiry date is not valid" });
      try {
        await ppeTables(env);
        const who = ppePerson(user);
        if (!who.key) return reply(env, 400, { error: "No name on your sign in" });
        const have = {};
        PPE_ITEMS.filter((x) => x.kind === "stock").forEach((x) => { if (ticked.has(x.id)) have[x.id] = true; });
        await env.DB.prepare("INSERT INTO ppe_state (person_key, name, email, have, first_aid_expiry, acknowledged_at, updated_at) VALUES (?, ?, ?, ?, ?, datetime('now'), datetime('now')) ON CONFLICT(person_key) DO UPDATE SET name = excluded.name, email = excluded.email, have = excluded.have, first_aid_expiry = excluded.first_aid_expiry, acknowledged_at = COALESCE(ppe_state.acknowledged_at, excluded.acknowledged_at), updated_at = excluded.updated_at")
          .bind(who.key, who.name, who.email, JSON.stringify(have), expiry || null).run();
        const mine = ((await env.DB.prepare("SELECT * FROM ppe_orders WHERE person_key = ?").bind(who.key).all()).results) || [];
        const openFor = (id) => mine.some((o) => ["requested", "approved", "provided"].includes(o.status) && ppeJson(o.items, []).some((x) => x.id === id));
        const missing = PPE_ITEMS.filter((x) => x.kind === "stock" && !have[x.id] && !openFor(x.id));
        let made = [];
        if (missing.length) {
          const list = missing.map((x) => { const sz = String(sizes[x.id] || "").toUpperCase(); return { id: x.id, label: x.label, kind: "stock", ...(x.sizes ? { size: x.sizes.includes(sz) ? sz : "" } : {}) }; });
          made = await ppeCreateOrders(env, who, list, "bag check");
        }
        return reply(env, 200, { saved: true, ordersMade: made.length, ...(await ppeMe(env, user)) });
      } catch (e) {
        console.log("ppe ack error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not save the PPE acknowledgement", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/ppe/orders" && request.method === "POST") {
      // A replenishment order from the signed in person. Any tier.
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const items = ppeCleanItems(body.items);
      if (!items) return reply(env, 400, { error: "Those items are not valid" });
      if (items.some((x) => { const it = ppeItem(x.id); return it && it.sizes && !x.size; })) return reply(env, 400, { error: "Choose a size" });
      try {
        await ppeTables(env);
        const who = ppePerson(user);
        if (!who.key) return reply(env, 400, { error: "No name on your sign in" });
        const made = await ppeCreateOrders(env, who, items, "order");
        return reply(env, 200, { ...(await ppeMe(env, user)), made });
      } catch (e) {
        console.log("ppe order error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not send the order", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/ppe/orders" && request.method === "GET") {
      // Every PPE order, for the office. Directors and managers only.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      try {
        await ppeTables(env);
        const rows = ((await env.DB.prepare("SELECT * FROM ppe_orders ORDER BY created_at DESC, rowid DESC LIMIT 300").all()).results) || [];
        return reply(env, 200, { orders: rows.map(ppeOrderOut) });
      } catch (e) { console.log("ppe list error", String(e && e.message || e)); return reply(env, 500, { error: "Could not read the PPE orders" }); }
    }
    const pod = url.pathname.match(/^\/api\/ppe\/orders\/([0-9a-fA-F-]{20,40})$/);
    if (pod && request.method === "DELETE") {
      // Removing a PPE order. The person who made it can delete it while it is requested, received or declined (not while the office is working on it). Directors (T1) can also remove closed orders from Team Requests, until the go live lock is on.
      try {
        await ppeTables(env);
        const row = await env.DB.prepare("SELECT id, status, person_key FROM ppe_orders WHERE id = ?").bind(pod[1]).first();
        if (!row) return reply(env, 200, { removed: false, alreadyGone: true, ...(await ppeMe(env, user)) });
        const owner = row.person_key === ppePerson(user).key;
        if (owner) {
          if (row.status !== "requested" && row.status !== "received" && row.status !== "declined") return reply(env, 409, { error: "The office is already working on that order, so it can not be deleted yet" });
        } else {
          if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
          if ((await requestsLock(env)).locked) return reply(env, 403, { error: REMOVAL_LOCKED });
          if (row.status !== "received" && row.status !== "declined") return reply(env, 409, { error: "Only a received or declined order can be removed" });
        }
        await env.DB.prepare("DELETE FROM ppe_orders WHERE id = ?").bind(row.id).run();
        return reply(env, 200, { removed: true, ...(await ppeMe(env, user)) });
      } catch (e) {
        console.log("ppe order delete error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not remove the order", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    const pom = url.pathname.match(/^\/api\/ppe\/orders\/([0-9a-fA-F-]{20,40})\/(status|receipt)$/);
    if (pom && request.method === "POST") {
      // status: the office approves, declines or marks provided; the owner marks received. receipt: the owner uploads the receipt for a bought item.
      try {
        await ppeTables(env);
        const row = await env.DB.prepare("SELECT * FROM ppe_orders WHERE id = ?").bind(pom[1]).first();
        if (!row) return reply(env, 404, { error: "not found" });
        const who = ppePerson(user);
        const office = user.tier === "T1" || user.tier === "T2";
        const owner = row.person_key === who.key;
        const items = ppeJson(row.items, []);
        const buy = items.some((x) => x.kind === "buy");
        const by = user.email || user.name || "unknown";
        const markReceived = async () => {
          await env.DB.prepare("UPDATE ppe_orders SET status = 'received', received_at = datetime('now') WHERE id = ?").bind(row.id).run();
          if (!buy) {
            const st = await env.DB.prepare("SELECT have FROM ppe_state WHERE person_key = ?").bind(row.person_key).first();
            const have = ppeJson(st && st.have, {});
            items.forEach((x) => { have[x.id] = true; });
            await env.DB.prepare("UPDATE ppe_state SET have = ?, updated_at = datetime('now') WHERE person_key = ?").bind(JSON.stringify(have), row.person_key).run();
          }
        };
        if (pom[2] === "receipt") {
          if (!owner) return reply(env, 403, { error: "forbidden" });
          if (!buy || row.status !== "approved") return reply(env, 409, { error: "This order is not waiting for a receipt" });
          const fileName = cleanName(url.searchParams.get("name")) || "receipt.pdf";
          const bytes = await request.arrayBuffer();
          if (!bytes.byteLength) return reply(env, 400, { error: "empty file" });
          if (bytes.byteLength > MAX_UPLOAD) return reply(env, 413, { error: "file too large" });
          const stamp = `${(row.po_number || "PPE").replace(/[^A-Za-z0-9-]/g, "")}-${profileKey(row.name) || "staff"}-${fileName}`;
          let item;
          try { item = await uploadToSharePoint(env, ["Receipts", "PPE"], stamp, bytes, request.headers.get("Content-Type"), false); }
          catch (e) { console.log("ppe receipt upload failed", String(e && e.message || e)); return reply(env, 502, { error: "The receipt could not be uploaded" }); }
          await env.DB.prepare("UPDATE ppe_orders SET receipt_item = ? WHERE id = ?").bind(item.id || "", row.id).run();
          await markReceived();
        } else {
          let body;
          try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
          const next = String(body.status || "");
          if (next === "received") {
            if (!owner || buy || row.status !== "provided") return reply(env, 409, { error: "This order cannot be marked received" });
            await markReceived();
          } else if (next === "approved" || next === "declined") {
            if (!office) return reply(env, 403, { error: "forbidden" });
            if (row.status !== "requested") return reply(env, 409, { error: "This order has already been dealt with" });
            if (next === "approved") {
              let po = "";
              if (buy) { const hi = (await env.DB.prepare("SELECT MAX(CAST(SUBSTR(po_number, 5) AS INTEGER)) AS m FROM ppe_orders WHERE po_number LIKE 'PPE-%'").first()).m || 1000; po = "PPE-" + String(Number(hi) + 1); }
              await env.DB.prepare("UPDATE ppe_orders SET status = 'approved', approved_by = ?, approved_at = datetime('now'), po_number = ? WHERE id = ?").bind(by, po || null, row.id).run();
            } else {
              await env.DB.prepare("UPDATE ppe_orders SET status = 'declined', declined_by = ?, declined_at = datetime('now') WHERE id = ?").bind(by, row.id).run();
            }
          } else if (next === "provided") {
            if (!office) return reply(env, 403, { error: "forbidden" });
            if (buy || row.status !== "approved") return reply(env, 409, { error: "This order cannot be marked provided" });
            await env.DB.prepare("UPDATE ppe_orders SET status = 'provided', provided_by = ?, provided_at = datetime('now') WHERE id = ?").bind(by, row.id).run();
          } else return reply(env, 400, { error: "That status is not valid" });
        }
        const done = await env.DB.prepare("SELECT * FROM ppe_orders WHERE id = ?").bind(row.id).first();
        return reply(env, 200, { order: ppeOrderOut(done) });
      } catch (e) {
        console.log("ppe status error", String(e && e.message || e));
        return reply(env, 500, { error: "Could not update the order", detail: String(e && e.message || e).slice(0, 200) });
      }
    }
    if (url.pathname === "/api/dfp/populate" && (request.method === "GET" || request.method === "POST")) {
      // One time seeding of DFP's tables from what Xero already holds. Director only (T1). After this, DFP's data is the source and Xero is fed from it.
      // GET is a preview: it reads Xero and reports what would be saved, and writes nothing to the new tables. POST saves.
      // dfp_xero gets DFP's six fields for each current employee, only when the person has no row yet. dfp_reservoir gets a row for each person who does not have one (the old tables' data is merged in by reservoirTables). xero_payroll rows are replaced with Xero's current values.
      // The reply carries counts, names and the names of empty fields. It never carries a personal value, bank details included.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        const save = request.method === "POST";
        const got = await dfpReadFromXero(env);
        if (got.fail) {
          console.log("dfp populate failed", got.fail.detail);
          return reply(env, 502, { error: "Xero refused the request", status: got.fail.status, detail: got.fail.detail });
        }
        const out = { saved: save, worker: WORKER_BUILD, employeesRead: got.people.length, xeroCalls: got.calls, dfpXeroAdded: 0, dfpXeroAlreadyThere: 0, dfpReservoirAdded: 0, dfpReservoirAlreadyThere: 0, xeroPayrollSaved: 0, noWorkEmail: [], people: [] };
        if (save) { await dfpTables(env); await reservoirTables(env); }
        for (const p of got.people) {
          const i = p.xero, y = p.payroll;
          const entry = { name: i.name, missing: dfpGaps(p), leaveBalancesFound: y.balanceCount, allowancesFound: y.allowanceCount };
          if (save) {
            if (i.email) {
              const bk = i.bank || {};
              const bankIn = { name: String(bk.name || "").trim(), bsb: String(bk.bsb || "").trim(), number: String(bk.number || "").trim() };
              const existing = await env.DB.prepare("SELECT email, bank_account_name, bsb, account_number FROM dfp_xero WHERE email = ?").bind(i.email).first();
              const blank = (v) => v === null || v === undefined || String(v).trim() === "";
              if (existing) {
                // Only blank fields are filled (a row made by a saved emergency contact has none of these yet). Anything already saved is kept, bank details included.
                await env.DB.prepare("UPDATE dfp_xero SET xero_employee_id = COALESCE(xero_employee_id, ?), name = COALESCE(name, ?), address = COALESCE(address, ?), dob = COALESCE(dob, ?), mobile = COALESCE(mobile, ?), start_date = COALESCE(start_date, ?), job_title = COALESCE(job_title, ?), bank_account_name = CASE WHEN bank_account_name IS NULL OR TRIM(bank_account_name) = '' THEN COALESCE(NULLIF(?, ''), bank_account_name) ELSE bank_account_name END, bsb = CASE WHEN bsb IS NULL OR TRIM(bsb) = '' THEN COALESCE(NULLIF(?, ''), bsb) ELSE bsb END, account_number = CASE WHEN account_number IS NULL OR TRIM(account_number) = '' THEN COALESCE(NULLIF(?, ''), account_number) ELSE account_number END WHERE email = ?").bind(i.xeroId, i.name, i.address, i.dob, normPhone(i.mobile), i.startDate, i.jobTitle, bankIn.name, bankIn.bsb, bankIn.number, i.email).run();
                out.dfpXeroAlreadyThere++; entry.dfpXero = "kept as it was, blanks filled";
                const filled = (blank(existing.bank_account_name) && bankIn.name ? 1 : 0) + (blank(existing.bsb) && bankIn.bsb ? 1 : 0) + (blank(existing.account_number) && bankIn.number ? 1 : 0);
                if (filled) { out.bankDetailsFilled = (out.bankDetailsFilled || 0) + 1; entry.bankDetails = filled === 3 ? "filled from Xero" : "partly filled from Xero"; }
                else if (!blank(existing.bank_account_name) || !blank(existing.bsb) || !blank(existing.account_number)) entry.bankDetails = "already saved, kept";
                else entry.bankDetails = "none in Xero";
              }
              else {
                await env.DB.prepare("INSERT OR IGNORE INTO dfp_xero (email, xero_employee_id, name, address, dob, mobile, start_date, job_title, bank_account_name, bsb, account_number, status, source, updated_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?, NULLIF(?, ''), NULLIF(?, ''), NULLIF(?, ''), 'ACTIVE', 'xero one time population', ?)").bind(i.email, i.xeroId, i.name, i.address, i.dob, normPhone(i.mobile), i.startDate, i.jobTitle, bankIn.name, bankIn.bsb, bankIn.number, user.name || user.email || "").run();
                out.dfpXeroAdded++; entry.dfpXero = "added";
                if (bankIn.name || bankIn.bsb || bankIn.number) { out.bankDetailsFilled = (out.bankDetailsFilled || 0) + 1; entry.bankDetails = "filled from Xero"; }
                else entry.bankDetails = "none in Xero";
              }
              // An emergency contact Xero already holds is copied in, but only for someone DFP has no saved contact for. A contact saved through Reservoir is never touched.
              const em = i.em || {};
              if (em.name || em.relationship || em.mobile || em.email) {
                const rr = await reservoirRowFor(env, i.email, i.name, true);
                const r2 = await env.DB.prepare("UPDATE dfp_reservoir SET em_name = ?, em_relationship = ?, em_mobile = ?, em_email = ?, em_saved_at = datetime('now') WHERE person_key = ? AND em_saved_at IS NULL").bind(em.name || "", em.relationship || "", normPhone(em.mobile || ""), em.email || "", rr.person_key).run();
                if (r2 && r2.meta && r2.meta.changes) { out.emergencySeeded = (out.emergencySeeded || 0) + 1; entry.emergencyContact = "copied from Xero"; }
              }
            } else { entry.dfpXero = "not saved, no work email in Xero"; }
            let superPct = y.superPct === undefined ? null : y.superPct, superRun = y.superRunDate || null;
            if (superPct === null) {
              const before = await env.DB.prepare("SELECT super_pct, super_run_date FROM xero_payroll WHERE employee_id = ?").bind(y.employeeId).first();
              if (before && before.super_pct !== null && before.super_pct !== undefined) { superPct = before.super_pct; superRun = before.super_run_date || null; }
            }
            const had = await reservoirRowFor(env, i.email, i.name, false);
            if (had) { out.dfpReservoirAlreadyThere++; entry.dfpReservoir = "already there"; }
            else { await reservoirRowFor(env, i.email, i.name, true); out.dfpReservoirAdded++; entry.dfpReservoir = "added"; }
            await env.DB.prepare("INSERT OR REPLACE INTO xero_payroll (employee_id, email, name, status, classification, employment_basis, rate_ordinary, rate_ot15, rate_ot2, has_overtime, rates_note, leave_loading, super_fund_name, super_member_number, super_usi, leave_balances, allowances, super_amount_period, last_pay_run_id, synced_at, source, annual_salary, week_hours, super_pct, super_run_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, NULL, NULL, datetime('now'), 'xero one time population', ?, ?, ?, ?)").bind(y.employeeId, y.email, y.name, y.status, y.classification, y.basis, y.ordinary, y.ot15, y.ot2, y.hasOvertime, y.ratesNote, y.leaveLoading, y.fundName, y.memberNumber, y.usi, y.balances, y.allowances, y.annual === null || y.annual === undefined ? null : y.annual, y.weekHours === null || y.weekHours === undefined ? null : y.weekHours, superPct, superRun).run();
            out.xeroPayrollSaved++;
          }
          if (!i.email) out.noWorkEmail.push(i.name);
          out.people.push(entry);
        }
        out.people.sort((a, b) => a.name.localeCompare(b.name));
        return reply(env, 200, out);
      } catch (e) {
        console.log("dfp populate error", String(e && e.message || e));
        return reply(env, 502, { error: String(e && e.message || e).slice(0, 300) });
      }
    }
    if (url.pathname === "/api/xero/test" && request.method === "GET") {
      // Reads the organisation name and employee count to prove the link and the payroll scopes work.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        const org = await xeroFetch(env, "api.xro/2.0/Organisation");
        const emp = await xeroFetch(env, "payroll.xro/1.0/Employees");
        let fields = null;
        let emergencyShape = null;
        try {
          const first = emp.ok ? ((await emp.clone().json()).Employees || [])[0] : null;
          if (first) {
            const one = await xeroFetch(env, `payroll.xro/1.0/Employees/${first.EmployeeID}`);
            const rec = one.ok ? ((await one.json()).Employees || [])[0] : null;
            if (rec) {
              fields = Object.keys(rec);
              emergencyShape = Object.keys(rec).filter((k) => /emergency/i.test(k)).map((k) => { const v = rec[k]; const o = Array.isArray(v) ? v[0] : v; return k + (o && typeof o === "object" ? " {" + Object.keys(o).join(", ") + "}" : ""); });
            }
          }
        } catch (e) { fields = null; }
        const orgData = org.ok ? await org.json() : null;
        const empData = emp.ok ? await emp.json() : null;
        return reply(env, 200, { accounting: org.status, organisation: orgData && orgData.Organisations && orgData.Organisations[0] ? orgData.Organisations[0].Name : null, payroll: emp.status, employees: empData && empData.Employees ? empData.Employees.length : null, fields, emergencyShape });
      } catch (e) { return reply(env, 502, { error: String(e.message || e).slice(0, 200) }); }
    }

    if (url.pathname === "/api/address" && request.method === "GET") {
      // Australian address suggestions from Geoscape. The key is a Worker secret, never in the app.
      if (!env.GEOSCAPE_KEY) return reply(env, 503, { error: "address service not set up" });
      const q = (url.searchParams.get("q") || "").trim().slice(0, 50);
      if (q.length < 3) return reply(env, 200, { options: [] });
      try {
        const res = await fetch(`https://api.psma.com.au/v1/predictive/address?query=${encodeURIComponent(q)}&maxNumberOfResults=5&stateTerritory=ALL`, { headers: { Authorization: env.GEOSCAPE_KEY, Accept: "application/json" } });
        if (!res.ok) return reply(env, 502, { error: "address service", detail: String(res.status) });
        const data = await res.json();
        const list = data.suggest || data.suggestions || [];
        return reply(env, 200, { options: list.map((x) => x.address).filter(Boolean) });
      } catch (e) { return reply(env, 502, { error: "address service" }); }
    }

    const um = url.pathname.match(/^\/api\/updates(\/seen)?$/);
    if (um) {
      // The T2 Updates monitor is for T1 only.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      if (request.method === "GET" && !um[1]) {
        const { results } = await env.DB.prepare("SELECT id, created_at, user_name, vehicle_id, vehicle_label, summary, detail, seen FROM t2_updates ORDER BY id DESC LIMIT 100").all();
        return reply(env, 200, results);
      }
      if (request.method === "PUT" && um[1]) {
        await env.DB.prepare("UPDATE t2_updates SET seen = 1 WHERE seen = 0").run();
        return reply(env, 200, { ok: true });
      }
      return reply(env, 405, { error: "method not allowed" });
    }

    const sm = url.pathname.match(/^\/api\/secrets(?:\/([A-Za-z0-9._-]{1,100}))?$/);
    if (sm) {
      // Company passwords and IDs. Only T1 and T2 can read or change them, and the Director ID is T1 only.
      if (user.tier === "T3") return reply(env, 403, { error: "forbidden" });
      const skey = sm[1];
      const allowed = (k) => k !== "company.directorId" || user.tier === "T1";
      if (request.method === "GET" && !skey) {
        await settingsTable(env);
        const { results } = await env.DB.prepare("SELECT key, value FROM app_settings WHERE key LIKE 'secret.%'").all();
        const out = {};
        results.forEach((r) => { const k = r.key.slice(7); if (allowed(k)) out[k] = r.value; });
        return reply(env, 200, out);
      }
      if (request.method === "PUT" && skey) {
        if (!allowed(skey)) return reply(env, 403, { error: "forbidden" });
        let body;
        try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
        const value = String(body.value ?? "").slice(0, 500);
        await settingsTable(env);
        await env.DB.prepare("INSERT INTO app_settings (key, value, by_name) VALUES (?, ?, ?) ON CONFLICT(key) DO UPDATE SET value = excluded.value, at = datetime('now'), by_name = excluded.by_name").bind("secret." + skey, value, user.name).run();
        return reply(env, 200, { ok: true });
      }
      return reply(env, 405, { error: "method not allowed" });
    }


    if (url.pathname === "/api/contacts" && request.method === "GET") {
      // Xero contact names, searched as the office types. Reads the saved list only, never Xero. T1 and T2 only.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      await xeroContactsTable(env);
      const q = String(url.searchParams.get("q") || "").trim().toLowerCase().slice(0, 60);
      if (q.length < 2) return reply(env, 200, { items: [] });
      const esc = q.replace(/[\\%_]/g, (m) => "\\" + m);
      const rows = ((await env.DB.prepare("SELECT contact_id, name FROM xero_contacts WHERE name_lc LIKE ? ESCAPE '\\' ORDER BY (name_lc LIKE ? ESCAPE '\\') DESC, name_lc LIMIT 8").bind(`%${esc}%`, `${esc}%`).all()).results) || [];
      return reply(env, 200, { items: rows.map((r) => ({ id: r.contact_id, name: r.name })) });
    }

    const rm = url.pathname.match(/^\/api\/fleet\/([A-Za-z0-9_-]{1,64})\/repairs$/);
    if (rm) {
      // Repairs and maintenance records on a vehicle. The office adds them (T1 and T2); nothing here edits or removes a record.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      await env.DB.prepare("CREATE TABLE IF NOT EXISTS fleet_repairs (id TEXT PRIMARY KEY, vehicle_id TEXT NOT NULL, repair_date TEXT NOT NULL, contact_id TEXT, contact_name TEXT, notes TEXT, docs TEXT, created_by TEXT, created_at TEXT DEFAULT (datetime('now')))").run();
      await env.DB.prepare("CREATE INDEX IF NOT EXISTS fleet_repairs_vehicle ON fleet_repairs (vehicle_id, repair_date)").run();
      const out = (r) => { let docs = []; try { docs = JSON.parse(r.docs || "[]"); } catch (e) { docs = []; } return { id: r.id, date: r.repair_date, contactId: r.contact_id || "", contactName: r.contact_name || "", notes: r.notes || "", docs, createdBy: r.created_by || "", createdAt: r.created_at || "" }; };
      if (request.method === "GET") {
        const rows = ((await env.DB.prepare("SELECT * FROM fleet_repairs WHERE vehicle_id = ? ORDER BY repair_date DESC, created_at DESC LIMIT 200").bind(rm[1]).all()).results) || [];
        return reply(env, 200, { items: rows.map(out) });
      }
      if (request.method === "POST") {
        const veh = await env.DB.prepare("SELECT rego, name FROM fleet WHERE id = ?").bind(rm[1]).first();
        if (!veh) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
        let body;
        try { body = await request.json(); } catch (e) { return reply(env, 400, { error: "bad json" }); }
        const date = String(body.date || "").slice(0, 10);
        if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) return reply(env, 400, { error: "A date is needed." });
        const contactName = String(body.contactName || "").trim().slice(0, 200);
        if (!contactName) return reply(env, 400, { error: "A contact is needed." });
        const docs = (Array.isArray(body.docs) ? body.docs : []).slice(0, 10).map((d) => ({ id: String(d.id || "").slice(0, 64), name: String(d.name || "").slice(0, 200), itemId: String(d.itemId || "").slice(0, 200), kind: d.kind === "invoice" ? "invoice" : "extra" }));
        const id = `r${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
        await env.DB.prepare("INSERT INTO fleet_repairs (id, vehicle_id, repair_date, contact_id, contact_name, notes, docs, created_by) VALUES (?, ?, ?, ?, ?, ?, ?, ?)").bind(id, rm[1], date, String(body.contactId || "").slice(0, 64), contactName, String(body.notes || "").slice(0, 2000), JSON.stringify(docs), user.name || "").run();
        if (user.tier === "T2") { try { await logUpdate(env, user, rm[1], veh.rego || veh.name || rm[1], "Repair Added", { date, url: "", notes: `${contactName}. ${String(body.notes || "").slice(0, 300)}` }); } catch (e) { /* the record is already saved */ } }
        const row = await env.DB.prepare("SELECT * FROM fleet_repairs WHERE id = ?").bind(id).first();
        return reply(env, 200, { item: out(row) });
      }
      return reply(env, 405, { error: "method not allowed" });
    }

    const tm = url.pathname.match(/^\/api\/fleet\/([A-Za-z0-9_-]{1,64})\/tools$/);
    if (tm) {
      // The Vehicle Tools check. Whoever takes in a vehicle ticks off what is on it; each save is kept, so the next swap can be compared with the last.
      await env.DB.prepare("CREATE TABLE IF NOT EXISTS vehicle_tool_items (id INTEGER PRIMARY KEY AUTOINCREMENT, grp TEXT NOT NULL, item TEXT NOT NULL, spec TEXT, sort INTEGER, active INTEGER DEFAULT 1)").run();
      await env.DB.prepare("CREATE TABLE IF NOT EXISTS vehicle_tool_checks (id TEXT PRIMARY KEY, vehicle_id TEXT NOT NULL, checked_by TEXT, checked_at TEXT DEFAULT (datetime('now')), missing_count INTEGER, notes TEXT, results TEXT)").run();
      await env.DB.prepare("CREATE INDEX IF NOT EXISTS vehicle_tool_checks_v ON vehicle_tool_checks (vehicle_id, checked_at)").run();
      await env.DB.prepare("DELETE FROM vehicle_tool_items WHERE grp = 'Vehicle Supply'").run(); // supplies are covered by the PPE screen
      const have = await env.DB.prepare("SELECT COUNT(*) AS n FROM vehicle_tool_items").first();
      if (!have || !have.n) {
        const seed = [["Company Tool", "Extension ladder", ""], ["Company Tool", "A-frame ladders", "Small and medium"], ["Company Tool", "Extractor fans", "2 units"], ["Company Tool", "SDS drill-bit set", "5 mm, 6.5 mm, 8 mm and 25 mm"], ["Company Tool", "Impact drill-bit set", "Phillips head, hex bits 5/16 and 3/8, and spade bits"], ["Company Tool", "Hole-saw kit", ""], ["Company Tool", "Gas manometer", ""], ["Company Tool", "Bucket pump", ""], ["Company Tool", "Water-pressure gauge and thermometer", ""], ["Company Tool", "String lines", ""], ["Company Tool", "Ratchet straps", ""], ["Company Tool", "30 m tape measure", ""], ["Company Tool", "B-press tool", ""], ["Company Tool", "Shovels", ""], ["Company Tool", "Probe", ""], ["Company Tool", "Broom, dustpan and brush", ""]];
        for (let i = 0; i < seed.length; i++) await env.DB.prepare("INSERT INTO vehicle_tool_items (grp, item, spec, sort) VALUES (?, ?, ?, ?)").bind(seed[i][0], seed[i][1], seed[i][2], i).run();
      }
      const lastOf = async () => {
        const r = await env.DB.prepare("SELECT * FROM vehicle_tool_checks WHERE vehicle_id = ? ORDER BY checked_at DESC, rowid DESC LIMIT 1").bind(tm[1]).first();
        if (!r) return null;
        let results = []; try { results = JSON.parse(r.results || "[]"); } catch (e) { results = []; }
        return { id: r.id, at: String(r.checked_at || "").slice(0, 10), by: r.checked_by || "", missing: results.filter((x) => !x.present).map((x) => x.item), missingCount: r.missing_count || 0, notes: r.notes || "" };
      };
      if (request.method === "GET") {
        const rows = ((await env.DB.prepare("SELECT id, grp, item, spec FROM vehicle_tool_items WHERE active = 1 ORDER BY sort, id").all()).results) || [];
        return reply(env, 200, { items: rows.map((r) => ({ id: r.id, group: r.grp, item: r.item, spec: r.spec || "" })), last: await lastOf() });
      }
      if (request.method === "POST") {
        const veh = await env.DB.prepare("SELECT rego, name FROM fleet WHERE id = ?").bind(tm[1]).first();
        if (!veh) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
        let body;
        try { body = await request.json(); } catch (e) { return reply(env, 400, { error: "bad json" }); }
        const rows = ((await env.DB.prepare("SELECT id, grp, item FROM vehicle_tool_items WHERE active = 1 ORDER BY sort, id").all()).results) || [];
        const present = new Set((Array.isArray(body.present) ? body.present : []).map((x) => Number(x)));
        const results = rows.map((r) => ({ id: r.id, group: r.grp, item: r.item, present: present.has(r.id) }));
        const missing = results.filter((x) => !x.present);
        const id = `t${Date.now().toString(36)}${Math.random().toString(36).slice(2, 7)}`;
        const notes = String(body.notes || "").slice(0, 2000);
        await env.DB.prepare("INSERT INTO vehicle_tool_checks (id, vehicle_id, checked_by, missing_count, notes, results) VALUES (?, ?, ?, ?, ?, ?)").bind(id, tm[1], user.name || "", missing.length, notes, JSON.stringify(results)).run();
        if (missing.length) {
          try {
            const tiers = await accessTiers(env);
            const label = veh.rego || veh.name || tm[1];
            for (const e of Object.keys(tiers).filter((k) => tiers[k] === "T1")) await notify(env, { to: e, kind: "toolcheck", title: `${label} Tools Missing`, body: `${missing.length} missing: ${missing.map((x) => x.item).join(", ").slice(0, 300)}`, link: "", by: user.name || "Reservoir" });
          } catch (e) { /* the check is already saved */ }
        }
        return reply(env, 200, { last: await lastOf() });
      }
      return reply(env, 405, { error: "method not allowed" });
    }

    const dm = url.pathname.match(/^\/api\/fleet\/([A-Za-z0-9_-]{1,64})\/docs$/);
    if (dm && request.method === "POST") {
      // Upload a document into the vehicle's Service folder, or its Finance or Repairs folder when folder=Finance or folder=Repairs. T1 and T2 only.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      const row = await env.DB.prepare("SELECT grp, rego, name FROM fleet WHERE id = ?").bind(dm[1]).first();
      if (!row) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
      const label = cleanName(row.rego || row.name || dm[1]);
      if (!label) return reply(env, 400, { error: "vehicle needs a rego or name" });
      const fileName = cleanName(url.searchParams.get("name")) || "document";
      const bytes = await request.arrayBuffer();
      if (!bytes.byteLength) return reply(env, 400, { error: "empty file" });
      if (bytes.byteLength > MAX_UPLOAD) return reply(env, 413, { error: "file too large" });
      try {
        const item = await uploadToSharePoint(env, ["Fleet Management", row.grp === "Machinery" ? "Machinery" : "Vehicles", label, ({ Finance: "Finance", Repairs: "Repairs" })[url.searchParams.get("folder")] || "Service"], fileName, bytes, request.headers.get("Content-Type"), url.searchParams.get("replace") === "1");
        if (user.tier === "T2") {
          try { await logUpdate(env, user, dm[1], row.rego || row.name || dm[1], "Document Uploaded", { date: new Date().toISOString().slice(0, 10), url: "", notes: item.name || fileName }); } catch (e) { /* upload already succeeded */ }
        }
        return reply(env, 200, { id: item.id, name: item.name || fileName });
      } catch (e) {
        return reply(env, 502, { error: "upload failed" });
      }
    }

    const fm = url.pathname.match(/^\/api\/docs\/([A-Za-z0-9!_-]{1,200})$/);
    if (fm && request.method === "GET") {
      // Open or download a stored document. T1 and T2 only.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      try {
        const metaRes = await graphFetch(env, `/sites/${siteId(env)}/drive/items/${fm[1]}?select=name,file,folder`);
        if (!metaRes.ok) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
        const meta = await metaRes.json();
        if (!meta.file) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
        const dl = await graphFetch(env, `/sites/${siteId(env)}/drive/items/${fm[1]}/content`, { redirect: "manual" });
        const loc = dl.headers.get("Location");
        const file = loc ? await fetch(loc) : dl;
        if (!file.ok) return reply(env, 502, { error: "download failed" });
        const type = (meta.file && meta.file.mimeType) || "application/octet-stream";
        const disp = url.searchParams.get("download") ? "attachment" : "inline";
        return new Response(file.body, { status: 200, headers: { ...cors(env), "Content-Type": type, "Content-Disposition": `${disp}; filename*=UTF-8''${encodeURIComponent(meta.name)}`, "Access-Control-Expose-Headers": "Content-Disposition", "Cache-Control": "no-store" } });
      } catch (e) {
        return reply(env, 502, { error: "download failed" });
      }
    }

    const fd = url.pathname.match(/^\/api\/docs\/([A-Za-z0-9!_-]{1,200})$/);
    if (fd && request.method === "DELETE") {
      // Delete a stored document. T1 only.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      try {
        const meta = await driveMeta(env, fd[1]);
        if (!meta) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
        const del = await graphFetch(env, `/sites/${siteId(env)}/drive/items/${fd[1]}`, { method: "DELETE" });
        if (!del.ok && del.status !== 404) return reply(env, 502, { error: "delete failed" });
        return reply(env, 200, { ok: true });
      } catch (e) {
        return reply(env, 502, { error: "delete failed" });
      }
    }

    const xm = url.pathname.match(/^\/api\/(?:docs|receipts)\/([A-Za-z0-9!_-]{1,200})\/xero$/);
    if (xm && request.method === "POST") {
      // Email a stored PDF to the Xero bills inbox. Reimbursements can only be released by T1.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      try {
        const meta = await driveMeta(env, xm[1]);
        if (!meta) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
        if (meta.folder.startsWith("Receipts/Reimbursements") && user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
        await sendToXero(env, user, meta);
        return reply(env, 200, { ok: true });
      } catch (e) {
        return reply(env, e.message === "too large" ? 413 : 502, { error: "send failed", detail: String(e.message || "").slice(0, 200) });
      }
    }

    // The receipt history. Every receipt sent from Reservoir is kept here, so it is still there after the app is closed and the office can see reimbursements from other phones. The files themselves are in SharePoint. The office (T1) sees every receipt, T2 sees their own. Only T1 marks a reimbursement paid or declined.
    const rdel = url.pathname.match(/^\/api\/receipt-records\/([A-Za-z0-9_-]{1,64})$/);
    if (rdel && request.method === "DELETE") {
      // Remove a receipt for good: the SharePoint file and the saved record. T1 only.
      if (user.tier !== "T1") return reply(env, 403, { error: "forbidden" });
      await receiptRecordsTable(env);
      const row = await env.DB.prepare("SELECT item_id FROM receipt_records WHERE id = ?").bind(rdel[1]).first();
      if (row && row.item_id) {
        try { await graphFetch(env, `/sites/${siteId(env)}/drive/items/${row.item_id}`, { method: "DELETE" }); } catch (e) { /* the record still goes */ }
      }
      await env.DB.prepare("DELETE FROM receipt_records WHERE id = ?").bind(rdel[1]).run();
      return reply(env, 200, { ok: true });
    }

    const rrm = url.pathname.match(/^\/api\/receipt-records(?:\/([A-Za-z0-9_-]{1,64})\/status)?$/);
    if (rrm) {
      // Anyone signed in can send a reimbursement and see their own. Only T1 and T2 handle company receipts, and only T1 processes a reimbursement.
      await receiptRecordsTable(env);
      const first = String(user.name || "").split(" ")[0];
      const mail = String(user.email || "").toLowerCase();
      const today = new Date().toISOString().slice(0, 10);
      if (request.method === "GET" && !rrm[1]) {
        const rows = user.tier === "T1"
          ? await env.DB.prepare("SELECT * FROM receipt_records ORDER BY created_at DESC LIMIT 300").all()
          : await env.DB.prepare("SELECT * FROM receipt_records WHERE by_email = ? ORDER BY created_at DESC LIMIT 300").bind(mail).all();
        return reply(env, 200, { receipts: (rows.results || []).map(receiptRecordOut) });
      }
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      if (request.method === "POST" && !rrm[1]) {
        const id = String(body.id || "");
        if (!/^[A-Za-z0-9_-]{1,64}$/.test(id)) return reply(env, 400, { error: "bad id" });
        const type = body.receiptType === "reimbursement" ? "reimbursement" : "company";
        if (type === "company" && user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
        const status = ["sent", "submitted"].includes(body.status) ? body.status : (type === "company" ? "sent" : "submitted");
        const view = typeof body.view === "string" && body.view.startsWith("data:image/") && body.view.length < 400000 ? body.view : "";
        const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || "")) ? body.date : today;
        // A director (never while viewing as someone) can add a reimbursement on behalf of a team member. It is filed under that person, so it shows in their own history, and it can go in already processed.
        let byName = first, byMail = mail, rowStatus = status, doneOn = "";
        const forMail = String(body.forEmail || "").trim().toLowerCase();
        if (forMail && forMail !== mail) {
          if (user.tier !== "T1" || user.viewAs || type !== "reimbursement") return reply(env, 403, { error: "forbidden" });
          const person = await env.DB.prepare("SELECT name FROM dfp_xero WHERE lower(email) = ?").bind(forMail).first();
          if (!person || !person.name) return reply(env, 400, { error: "That team member was not found" });
          byName = String(person.name).split(" ")[0];
          byMail = forMail;
        }
        if (type === "reimbursement" && body.status === "paid" && user.tier === "T1") { rowStatus = "paid"; doneOn = date; }
        await env.DB.prepare("INSERT OR IGNORE INTO receipt_records (id, by_name, by_email, type, amount, note, item_id, file_name, status, view, sent_date, done_date) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)")
          .bind(id, byName, byMail, type, type === "reimbursement" ? txt(body.amount, 12) : "", txt(body.text, 500), txt(body.itemId, 200), txt(body.fileName, 200), rowStatus, view, date, doneOn).run();
        if (type === "reimbursement" && rowStatus === "paid") await notify(env, { to: byMail, kind: "receipt", title: "Reimbursement request processed", body: "Your reimbursement has been paid. Open Receipts to see it.", link: "receipts", by: user.name || mail });
        return reply(env, 200, { ok: true });
      }
      if (request.method === "POST" && rrm[1]) {
        if (!isApprover(user)) return reply(env, 403, { error: "forbidden" });
        if (!["paid", "declined"].includes(body.status)) return reply(env, 400, { error: "bad status" });
        const date = /^\d{4}-\d{2}-\d{2}$/.test(String(body.date || "")) ? body.date : today;
        await env.DB.prepare("UPDATE receipt_records SET status = ?, done_date = ? WHERE id = ?").bind(body.status, date, rrm[1]).run();
        // The person being reimbursed is told it is processed, even when the director who processed it is the one who logged it.
        if (body.status === "paid") {
          const rr = await env.DB.prepare("SELECT by_email, type FROM receipt_records WHERE id = ?").bind(rrm[1]).first();
          if (rr && rr.by_email && rr.type === "reimbursement") await notify(env, { to: rr.by_email, kind: "receipt", title: "Reimbursement request processed", body: "Your reimbursement has been paid. Open Receipts to see it.", link: "receipts", by: user.name || mail });
        }
        return reply(env, 200, { ok: true });
      }
      return reply(env, 405, { error: "method not allowed" });
    }

    if (url.pathname === "/api/receipts" && request.method === "POST") {
      // Upload a receipt photo or PDF. Anyone signed in can upload a reimbursement. Company receipts are T1 and T2.
      const fileName = cleanName(url.searchParams.get("name")) || "receipt.pdf";
      const kind = url.searchParams.get("kind") === "reimbursement" ? "Reimbursements" : "Company";
      if (kind === "Company" && user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      const bytes = await request.arrayBuffer();
      if (!bytes.byteLength) return reply(env, 400, { error: "empty file" });
      if (bytes.byteLength > MAX_UPLOAD) return reply(env, 413, { error: "file too large" });
      try {
        const item = await uploadToSharePoint(env, ["Receipts", kind], fileName, bytes, request.headers.get("Content-Type"), false);
        return reply(env, 200, { itemId: item.id, name: item.name || fileName });
      } catch (e) {
        return reply(env, 502, { error: "upload failed" });
      }
    }

    if (url.pathname === "/api/fleet-addresses" && request.method === "GET") {
      // Name, home address and emergency contact, for the fleet screens and the Operations team screen. T1 and T2.
      if (user.tier !== "T1" && user.tier !== "T2") return reply(env, 403, { error: "forbidden" });
      try {
        const rows = ((await env.DB.prepare("SELECT x.name, x.address, x.mobile, r.personal_email, r.preferred_name, r.em_name, r.em_relationship, r.em_mobile, r.em_email, r.em_saved_at FROM dfp_xero x LEFT JOIN dfp_reservoir r ON lower(r.email) = lower(x.email) WHERE x.name IS NOT NULL AND x.name <> ''").all()).results) || [];
        return reply(env, 200, { people: rows.map((r) => ({ name: r.name, address: r.address || "", mobile: r.mobile || "", personalEmail: r.personal_email || "", preferredName: String(r.preferred_name || "").trim(), ...(r.em_saved_at ? { emName: r.em_name || "", emRelationship: r.em_relationship || "", emMobile: r.em_mobile || "", emEmail: r.em_email || "" } : {}) })) });
      } catch (e) {
        return reply(env, 502, { error: "could not read addresses" });
      }
    }

    const m = url.pathname.match(/^\/api\/fleet(?:\/([A-Za-z0-9_-]{1,64}))?$/);
    if (!m) return reply(env, 404, { error: "not found", worker: WORKER_BUILD });
    const id = m[1];
    const canEdit = user.tier === "T1" || user.tier === "T2";

    await fleetCols(env);
    if (request.method === "GET" && !id) {
      const { results } = await env.DB.prepare("SELECT * FROM fleet ORDER BY grp, sort_order, id").all();
      return reply(env, 200, results.map(rowToItem));
    }

    if (request.method === "PUT" && id) {
      let body;
      try { body = await request.json(); } catch { return reply(env, 400, { error: "bad json" }); }
      const existing = await env.DB.prepare("SELECT * FROM fleet WHERE id = ?").bind(id).first();
      // Saves are changes, never whole records. The app sends only the fields, requests and service entries that changed, plus the ids it removed. Each is applied on its own to what is stored now, so requests, service history and the vehicle details can never write over one another, and one device can never write an old copy over another device's work.
      if (body && body.delta === true) {
        const arr = (v) => Array.isArray(v) ? v : [];
        const applyList = (storedJson, upsert, remove) => {
          let stored;
          try { stored = JSON.parse(storedJson || "[]"); } catch { stored = []; }
          if (!Array.isArray(stored)) stored = [];
          const up = arr(upsert).filter((x) => x && x.id !== undefined && x.id !== null);
          const gone = new Set(arr(remove).map(String));
          const byId = new Map(up.map((x) => [String(x.id), x]));
          const used = new Set();
          const kept = stored.filter((x) => !(x && gone.has(String(x.id)))).map((x) => {
            const k = x && String(x.id);
            if (byId.has(k)) { used.add(k); return byId.get(k); }
            return x;
          });
          return [...up.filter((x) => !used.has(String(x.id))), ...kept];
        };
        const touchRepairs = arr(body.repairsUpsert).length + arr(body.repairsRemove).length > 0;
        const touchHistory = canEdit && arr(body.historyUpsert).length + arr(body.historyRemove).length > 0;
        if (!existing && !canEdit) return reply(env, 403, { error: "forbidden" });
        const grp = body.group === "Machinery" ? "Machinery" : "Vehicles";
        const sent = body.fields && typeof body.fields === "object" ? body.fields : {};
        const names = canEdit ? FIELDS.filter((f) => Object.prototype.hasOwnProperty.call(sent, f)) : [];
        const clean = (f) => f === "servicePhone" ? normPhone(String(sent[f] ?? "").slice(0, 500)) : String(sent[f] ?? "").slice(0, fleetLimit(f));
        const newRepairs = JSON.stringify(applyList(existing ? existing.repairs : "[]", body.repairsUpsert, body.repairsRemove));
        if (newRepairs.length > 900000) return reply(env, 413, { error: "too large" });
        const newHistory = JSON.stringify(applyList(existing ? existing.serviceHistory : "[]", canEdit ? body.historyUpsert : [], canEdit ? body.historyRemove : []).slice(0, 500));
        if (newHistory.length > 400000) return reply(env, 413, { error: "too large" });
        if (!existing) {
          const next = await env.DB.prepare("SELECT COALESCE(MAX(sort_order), 0) + 1 AS n FROM fleet WHERE grp = ?").bind(grp).first();
          await env.DB.prepare(`INSERT INTO fleet (id, grp, ${FIELDS.join(", ")}, repairs, serviceHistory, sort_order, updated_by) VALUES (?, ?, ${FIELDS.map(() => "?").join(", ")}, ?, ?, ?, ?)`).bind(id, grp, ...FIELDS.map((f) => clean(f)), newRepairs, newHistory, next.n, user.name).run();
          return reply(env, 200, { ok: true });
        }
        const sets = [];
        const binds = [];
        if (canEdit && body.group && grp !== existing.grp) { sets.push("grp = ?"); binds.push(grp); }
        names.forEach((f) => { sets.push(`${f} = ?`); binds.push(clean(f)); });
        if (touchRepairs) { sets.push("repairs = ?"); binds.push(newRepairs); }
        if (touchHistory) { sets.push("serviceHistory = ?"); binds.push(newHistory); }
        if (!sets.length) return reply(env, 200, { ok: true });
        await env.DB.prepare(`UPDATE fleet SET ${sets.join(", ")}, updated_at = datetime('now'), updated_by = ? WHERE id = ?`).bind(...binds, user.name, id).run();
        if (canEdit && names.includes("rego") && env.GRAPH_CLIENT_SECRET && existing.rego && String(sent.rego || "") && String(sent.rego) !== String(existing.rego)) {
          try {
            const grpName = existing.grp === "Machinery" ? "Machinery" : "Vehicles";
            const oldPath = pathOf(["Fleet Management", grpName, cleanName(existing.rego)]);
            await graphFetch(env, `/sites/${siteId(env)}/drive/root:/${oldPath}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: cleanName(sent.rego) }) });
          } catch (e) { /* folder may not exist yet */ }
        }
        if (user.tier === "T2") {
          try {
            const label = sent.rego || existing.rego || existing.name || id;
            let oldHistory = [];
            try { oldHistory = JSON.parse(existing.serviceHistory || "[]"); } catch { oldHistory = []; }
            const known = new Set(oldHistory.map((e) => String(e.id)));
            for (const e of arr(body.historyUpsert).filter((e) => e && !known.has(String(e.id)))) {
              await logUpdate(env, user, id, label, "Service Added", { date: String(e.date || "").slice(0, 10), url: String(e.url || "").slice(0, 500), notes: String(e.notes || "").slice(0, 1000) });
            }
            for (const f of names) {
              if (f === "lastService" || f === "nextService") continue;
              const to = clean(f);
              if (String(existing[f] ?? "") !== to) await logUpdate(env, user, id, label, `${LABELS[f] || f} Changed`, { from: String(existing[f] ?? "").slice(0, 300), to: to.slice(0, 300) });
            }
          } catch (e) { /* the change itself has already been saved */ }
        }
        return reply(env, 200, { ok: true });
      }
      // Whole record saves from an older copy of the app are refused on a vehicle that already exists, because that is the save that wiped service entries.
      if (existing) return reply(env, 409, { error: "This copy of the app is out of date. Close it and open it again to update." });
      const incomingRepairs = Array.isArray(body.repairs) ? body.repairs : [];
      const repairs = JSON.stringify(existing ? mergeById(incomingRepairs, existing.repairs, body.knownRepairs) : incomingRepairs);
      if (repairs.length > 900000) return reply(env, 413, { error: "too large" });

      if (!canEdit) {
        // Team members can only add vehicle reports to a vehicle that already exists.
        if (!existing) return reply(env, 403, { error: "forbidden" });
        await env.DB.prepare("UPDATE fleet SET repairs = ?, updated_at = datetime('now'), updated_by = ? WHERE id = ?").bind(repairs, user.name, id).run();
        return reply(env, 200, { ok: true });
      }

      const incomingHistory = Array.isArray(body.serviceHistory) ? body.serviceHistory.slice(0, 500) : [];
      const history = JSON.stringify((existing ? mergeById(incomingHistory, existing.serviceHistory, body.knownHistory) : incomingHistory).slice(0, 500));
      if (history.length > 400000) return reply(env, 413, { error: "too large" });
      const changes = [];
      const grp = body.group === "Machinery" ? "Machinery" : "Vehicles";
      const vals = FIELDS.map((f) => f === "servicePhone" ? normPhone(String(body[f] ?? "").slice(0, 500)) : String(body[f] ?? "").slice(0, fleetLimit(f)));
      if (existing) {
        await env.DB.prepare(`UPDATE fleet SET grp = ?, ${FIELDS.map((f) => `${f} = ?`).join(", ")}, repairs = ?, serviceHistory = ?, updated_at = datetime('now'), updated_by = ? WHERE id = ?`).bind(grp, ...vals, repairs, history, user.name, id).run();
        if (env.GRAPH_CLIENT_SECRET && existing.rego && String(body.rego || "") && String(body.rego) !== String(existing.rego)) {
          try {
            const grpName = existing.grp === "Machinery" ? "Machinery" : "Vehicles";
            const oldPath = pathOf(["Fleet Management", grpName, cleanName(existing.rego)]);
            await graphFetch(env, `/sites/${siteId(env)}/drive/root:/${oldPath}`, { method: "PATCH", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ name: cleanName(body.rego) }) });
          } catch (e) { /* folder may not exist yet */ }
        }
        if (user.tier === "T2") {
          try {
            const label = body.rego || existing.rego || existing.name || id;
            let oldHistory = [];
            try { oldHistory = JSON.parse(existing.serviceHistory || "[]"); } catch { oldHistory = []; }
            const known = new Set(oldHistory.map((e) => e.id));
            for (const e of JSON.parse(history).filter((e) => e && !known.has(e.id))) {
              await logUpdate(env, user, id, label, "Service Added", { date: String(e.date || "").slice(0, 10), url: String(e.url || "").slice(0, 500), notes: String(e.notes || "").slice(0, 1000) });
            }
            FIELDS.forEach((f, i) => { if (f !== "lastService" && f !== "nextService" && String(existing[f] ?? "") !== vals[i]) changes.push([f, String(existing[f] ?? ""), vals[i]]); });
            for (const [f, from, to] of changes) await logUpdate(env, user, id, label, `${LABELS[f] || f} Changed`, { from: from.slice(0, 300), to: to.slice(0, 300) });
          } catch (e) { /* the change itself has already been saved */ }
        }
      } else {
        const next = await env.DB.prepare("SELECT COALESCE(MAX(sort_order), 0) + 1 AS n FROM fleet WHERE grp = ?").bind(grp).first();
        await env.DB.prepare(`INSERT INTO fleet (id, grp, ${FIELDS.join(", ")}, repairs, serviceHistory, sort_order, updated_by) VALUES (?, ?, ${FIELDS.map(() => "?").join(", ")}, ?, ?, ?, ?)`).bind(id, grp, ...vals, repairs, history, next.n, user.name).run();
      }
      return reply(env, 200, { ok: true });
    }

    if (request.method === "DELETE" && id) {
      if (!canEdit) return reply(env, 403, { error: "forbidden" });
      await env.DB.prepare("DELETE FROM fleet WHERE id = ?").bind(id).run();
      return reply(env, 200, { ok: true });
    }

    return reply(env, 405, { error: "method not allowed" });
  }
};

globalThis.w169_BOTTOM = "Reservoir-Worker-2026.10.28-w169";
// Reservoir-Worker-2026.10.28-w169
