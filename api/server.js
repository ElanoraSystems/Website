// Contact form backend: receives the website form and emails it through Microsoft Graph.
// No dependencies; needs Node 18+ for the built-in fetch.
const http = require("http");

const {
  TENANT_ID, CLIENT_ID, CLIENT_SECRET, MAIL_FROM, MAIL_TO, DRY_RUN,
  ALLOWED_ORIGIN = "https://elanorasystems.com,https://www.elanorasystems.com",
  PORT = 10000,
} = process.env;
const ORIGINS = ALLOWED_ORIGIN.split(",").map((o) => o.trim());

// ponytail: in-memory limit per IP, resets on restart; move to a shared store if abuse appears
const hits = new Map();
function limited(ip) {
  const now = Date.now(), recent = (hits.get(ip) || []).filter((t) => now - t < 10 * 60 * 1000);
  recent.push(now); hits.set(ip, recent);
  return recent.length > 5;
}

let cached = { value: "", exp: 0 };
async function token() {
  if (cached.value && Date.now() < cached.exp) return cached.value;
  const r = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body: new URLSearchParams({ client_id: CLIENT_ID, client_secret: CLIENT_SECRET, scope: "https://graph.microsoft.com/.default", grant_type: "client_credentials" }),
  });
  if (!r.ok) throw new Error(`token ${r.status}`);
  const j = await r.json();
  cached = { value: j.access_token, exp: Date.now() + (j.expires_in - 120) * 1000 };
  return cached.value;
}

async function sendMail({ name, email, message }) {
  const clean = (s) => s.replace(/[\r\n]+/g, " ");
  const payload = {
    message: {
      subject: `Website enquiry from ${clean(name)}`,
      body: { contentType: "Text", content: `Name: ${clean(name)}\nEmail: ${clean(email)}\n\n${message}` },
      toRecipients: [{ emailAddress: { address: MAIL_TO } }],
      replyTo: [{ emailAddress: { address: email, name: clean(name) } }],
    },
    saveToSentItems: false,
  };
  if (DRY_RUN) { console.log("DRY_RUN mail:", JSON.stringify(payload)); return; }
  const r = await fetch(`https://graph.microsoft.com/v1.0/users/${encodeURIComponent(MAIL_FROM)}/sendMail`, {
    method: "POST",
    headers: { Authorization: `Bearer ${await token()}`, "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });
  if (!r.ok) throw new Error(`sendMail ${r.status} ${await r.text()}`);
}

function reply(res, code, origin, body = {}) {
  const h = { "Content-Type": "application/json" };
  if (ORIGINS.includes(origin)) { h["Access-Control-Allow-Origin"] = origin; h["Vary"] = "Origin"; }
  res.writeHead(code, h); res.end(JSON.stringify(body));
}

http.createServer((req, res) => {
  const origin = req.headers.origin || "";
  if (req.method === "GET" && req.url === "/") return reply(res, 200, origin, { ok: true });
  if (req.method === "OPTIONS") {
    res.writeHead(204, { "Access-Control-Allow-Origin": ORIGINS.includes(origin) ? origin : "", "Access-Control-Allow-Methods": "POST", "Access-Control-Allow-Headers": "Content-Type", "Vary": "Origin" });
    return res.end();
  }
  if (req.method !== "POST" || req.url !== "/contact") return reply(res, 404, origin, { error: "Not found" });
  if (!ORIGINS.includes(origin)) return reply(res, 403, origin, { error: "Forbidden" });
  const ip = (req.headers["x-forwarded-for"] || req.socket.remoteAddress || "").split(",")[0].trim();
  if (limited(ip)) return reply(res, 429, origin, { error: "Too many messages. Please try again later." });

  let raw = "";
  req.on("data", (c) => { raw += c; if (raw.length > 10000) req.destroy(); });
  req.on("end", async () => {
    let d;
    try { d = JSON.parse(raw); } catch { return reply(res, 400, origin, { error: "Invalid request" }); }
    const name = String(d.name || "").trim(), email = String(d.email || "").trim(), message = String(d.message || "").trim();
    if (d.website) return reply(res, 200, origin, { ok: true }); // spam trap filled: pretend success
    if (!name || name.length > 100 || !/^\S+@\S+\.\S+$/.test(email) || email.length > 200 || !message || message.length > 5000)
      return reply(res, 400, origin, { error: "Please check your name, email and message." });
    try { await sendMail({ name, email, message }); reply(res, 200, origin, { ok: true }); }
    catch (e) { console.error(e.message); reply(res, 502, origin, { error: "Message could not be sent." }); }
  });
}).listen(PORT, () => console.log(`contact api on ${PORT}`));
