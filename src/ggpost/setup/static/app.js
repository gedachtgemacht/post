"use strict";

/* post. setup. No framework, no external requests. */

const params = new URLSearchParams(location.search);
const TOKEN = params.get("t") || "";

const S = {
  state: null,
  outlook: null,      // { name, picked, accounts:[{email, kind, picked}], status, message, note }
  imap: [],           // [{ email, label, imap, smtp, picked, hint, help_url, supported, has_password, status, message, provider }]
  notes: [],          // notes about detected programs without accounts (new Outlook, Apple Mail)
  apps: [],
  step: 1,
  saved: false,
  registered: [],
};

const $ = (sel, root = document) => root.querySelector(sel);
const el = (tag, attrs = {}, ...children) => {
  const node = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) {
    if (v === false || v === null || v === undefined) continue;
    if (k === "class") node.className = v;
    else if (k === "text") node.textContent = v;
    else if (k.startsWith("on")) node.addEventListener(k.slice(2), v);
    else node.setAttribute(k, v === true ? "" : v);
  }
  for (const c of children.flat()) if (c !== null && c !== undefined && c !== false) node.append(c);
  return node;
};

function announce(text) {
  const node = $("#announce");
  node.textContent = "";
  setTimeout(() => { node.textContent = text; }, 30);
}

async function api(path, body) {
  const opts = { headers: { "X-Post-Token": TOKEN } };
  if (body !== undefined) {
    opts.method = "POST";
    opts.headers["Content-Type"] = "application/json";
    opts.body = JSON.stringify(body);
  }
  let res;
  try {
    res = await fetch(path, opts);
  } catch (e) {
    throw new Error("Setup is no longer reachable. Was post. closed?");
  }
  let data = {};
  try { data = await res.json(); } catch (e) { /* leer */ }
  if (!res.ok) throw new Error(data.error || `Error ${res.status}`);
  return data;
}

function setStatus(node, text, kind = "") {
  node.textContent = text || "";
  node.className = node.className.replace(/\b(ok|err|busy)\b/g, "").trim();
  if (!node.classList.contains("status")) node.classList.add("status");
  if (kind) node.classList.add(kind);
}

function serverLine(s) {
  if (!s || !s.imap) return "No server settings yet.";
  const sec = { ssl: "SSL/TLS", starttls: "STARTTLS", plain: "unencrypted" }[s.imap.security] || s.imap.security;
  let line = `Posteingang ${s.imap.host}:${s.imap.port} ${sec}`;
  if (s.smtp) line += `, Postausgang ${s.smtp.host}:${s.smtp.port}`;
  return line;
}

const isWin = () => S.state && S.state.platform === "win32";
const vaultName = () => (!S.state ? "in your system keychain"
  : S.state.platform === "win32" ? "in the Windows Credential Manager"
  : S.state.platform === "darwin" ? "in the macOS Keychain" : "in your system keychain");

/* ── Modell aufbauen ─────────────────────────────── */

function buildModel(state) {
  const cfg = state.config;
  const det = state.detection;
  const cfgOutlook = cfg.sources.find((s) => s.kind === "outlook");
  const cfgImap = cfg.sources.filter((s) => s.kind === "imap");

  const oc = det.clients.find((c) => c.id === "outlook_classic");
  if (oc || cfgOutlook) {
    const wanted = cfgOutlook && cfgOutlook.outlook_accounts ? cfgOutlook.outlook_accounts.map((a) => a.toLowerCase()) : null;
    const accounts = (oc ? oc.accounts : []).map((a) => ({
      email: a.email,
      kind: a.kind,
      picked: wanted ? wanted.includes(a.email.toLowerCase()) : true,
    }));
    S.outlook = {
      name: oc ? oc.name : "Outlook (klassisch)",
      picked: cfg.exists ? Boolean(cfgOutlook) : Boolean(oc && oc.in_use && accounts.length),
      accounts,
      note: oc ? oc.note : "",
      status: "",
      message: "",
    };
  }

  for (const s of cfgImap) {
    S.imap.push({ ...s, picked: true, status: s.has_password ? "saved" : "", message: "" });
  }
  for (const c of det.clients) {
    if (c.id === "outlook_classic") continue;
    if (c.note) S.notes.push({ name: c.name, note: c.note });
    for (const a of c.accounts) {
      if (S.imap.some((x) => x.email.toLowerCase() === a.email.toLowerCase())) continue;
      S.imap.push({
        email: a.email,
        label: a.name || "",
        imap: a.imap,
        smtp: a.smtp,
        hint: a.hint || "",
        help_url: a.help_url || "",
        supported: a.supported !== false,
        provider: c.name,
        picked: !cfg.exists && c.in_use && a.supported !== false,
        status: "",
        message: "",
        needsDiscovery: !a.imap && a.supported !== false,
      });
    }
  }

  S.apps = state.apps.map((a) => ({ ...a, picked: a.installed && !a.registered }));
  $("#perm-modify").checked = cfg.permissions.modify;
  $("#perm-send").checked = cfg.permissions.send;
  $("#lede").textContent = `Connects your mailbox to your AI. post. itself runs on this computer only, and passwords live ${vaultName()}. Whatever your AI reads from your mail, the AI app sends to its provider, exactly as if you had pasted it into the chat. With a local model everything stays in house.`;
}

/* ── Darstellung ─────────────────────────────────── */

function pickedMailboxes() {
  const out = [];
  if (S.outlook && S.outlook.picked) out.push(...S.outlook.accounts.filter((a) => a.picked).map((a) => a.email));
  if (S.outlook && S.outlook.picked && !S.outlook.accounts.length) out.push("Outlook");
  out.push(...S.imap.filter((s) => s.picked).map((s) => s.email));
  return out;
}

function connectedApps() {
  return S.apps.filter((a) => a.registered || S.registered.includes(a.id));
}

function renderRoute() {
  const mails = pickedMailboxes();
  const mailText = !mails.length ? "not chosen yet" : mails.length === 1 ? mails[0] : `${mails.length} mailboxes`;
  $("#route-mail").textContent = mailText;
  $("#route-mail").title = mails.join(", ");
  $("#stop-mail").classList.toggle("chosen", mails.length > 0);
  $("#stop-mail").classList.toggle("live", S.saved);
  $("#leg-left").classList.toggle("live", S.saved);

  const apps = connectedApps();
  $("#route-ai").textContent = !apps.length ? "not chosen yet" : apps.length === 1 ? apps[0].name : `${apps.length} AI apps`;
  $("#route-ai").title = apps.map((a) => a.name).join(", ");
  $("#stop-ai").classList.toggle("chosen", apps.length > 0);
  const done = S.step === 4 && S.saved && apps.length > 0;
  $("#stop-ai").classList.toggle("live", done);
  $("#leg-right").classList.toggle("live", done);
}

function renderSteps(focus = false) {
  for (let i = 1; i <= 4; i++) {
    const node = $(`#step-${i}`);
    node.classList.toggle("open", S.step === i);
    node.classList.toggle("done", i < S.step || (i === 4 && S.step === 4));
    node.classList.toggle("todo", i > S.step);
    if (S.step === i) node.setAttribute("aria-current", "step"); else node.removeAttribute("aria-current");
    const edit = node.querySelector("[data-edit]");
    if (edit) edit.hidden = !(i < S.step && S.step < 4);
  }
  const mails = pickedMailboxes();
  $("#sum-1").textContent = S.step > 1 ? (mails.length > 2 ? `${mails.length} mailboxes` : mails.join(", ")) : "";
  const perms = ["Read", "Drafts"];
  if ($("#perm-modify").checked) perms.push("Sort and delete");
  if ($("#perm-send").checked) perms.push("Send");
  $("#sum-2").textContent = S.step > 2 ? perms.join(", ") : "";
  const apps = connectedApps().map((a) => a.name);
  $("#sum-3").textContent = S.step > 3 ? (apps.join(", ") || "skipped") : "";
  renderRoute();
  if (focus) $(`#step-${S.step} h2`).focus();
}

function goto(step) {
  S.step = step;
  renderSteps(true);
  if (step === 1) renderSources();
  if (step === 3) renderApps();
  if (step === 4) renderPrompts();
}

function renderSources() {
  const root = $("#sources");
  root.replaceChildren();

  if (S.outlook) {
    const o = S.outlook;
    const list = el("ul", { class: "rows" });
    const ok = o.status === "ok";
    list.append(el("li", { class: "row" + (o.picked ? " picked" : "") },
      el("label", { class: "check" },
        el("input", { type: "checkbox", class: "pick", checked: o.picked, onchange: (e) => { o.picked = e.target.checked; renderSources(); renderRoute(); } }),
        el("span", { class: "box", "aria-hidden": "true" }),
        el("span", { text: o.name })),
      ok ? el("span", { class: "tag ok", text: "verbunden" }) : el("span")));
    if (o.picked) {
      for (const a of o.accounts) {
        list.append(el("li", { class: "row sub" },
          el("label", { class: "check", title: a.email },
            el("input", { type: "checkbox", checked: a.picked, onchange: (e) => { a.picked = e.target.checked; renderRoute(); } }),
            el("span", { class: "box", "aria-hidden": "true" }),
            el("span", { class: "addr", text: a.email })),
          el("span", { class: "tag", text: a.kind || "Account" })));
      }
      const status = el("span", { class: "status" });
      setStatus(status, o.message, ok ? "ok" : o.status === "err" ? "err" : o.status === "busy" ? "busy" : "");
      list.append(el("li", { class: "row sub" },
        el("div", { class: "pw-line" },
          el("button", { type: "button", class: "btn small", onclick: testOutlook, text: "Test connection", disabled: o.status === "busy" }),
          status)));
    }
    root.append(el("section", { class: "source" },
      el("div", { class: "source-head" }, el("h3", { text: "Outlook on this computer" })),
      el("p", { class: "source-note", text: o.note || "post. drives your Outlook directly. No password needed, every kind of account works." }),
      list));
  }

  for (const n of S.notes) {
    root.append(el("section", { class: "source" },
      el("div", { class: "source-head" }, el("h3", { text: n.name })),
      el("p", { class: "source-note", text: n.note })));
  }

  if (S.imap.length) {
    const list = el("ul", { class: "rows" });
    S.imap.forEach((s) => list.append(imapRow(s)));
    root.append(el("section", { class: "source" },
      el("div", { class: "source-head" }, el("h3", { text: "Other mailboxes" })),
      el("p", { class: "source-note", text: `Enter the password once and test it. Afterwards it only lives ${vaultName()}.` }),
      list));
  }

  if (!S.outlook && !S.imap.length && !S.notes.length) {
    root.append(el("p", { class: "source-note", text: "No mail program on this computer can be used by post. directly. Add your mailbox by mail address below." }));
  }
}

function imapRow(s) {
  const node = $("#tpl-imap").content.firstElementChild.cloneNode(true);
  const unsupported = s.supported === false;
  node.classList.toggle("picked", s.picked);
  node.classList.toggle("off", unsupported);
  node.classList.toggle("unsupported", unsupported);
  node.dataset.email = s.email;
  const pick = $(".pick", node);
  pick.checked = s.picked && !unsupported;
  pick.disabled = unsupported;
  pick.addEventListener("change", () => { s.picked = pick.checked; node.classList.toggle("picked", s.picked); renderRoute(); });
  $(".addr", node).textContent = s.email;
  $(".check", node).title = s.email;
  const tag = $(".tag", node);
  const tagText = unsupported ? "not possible" : s.status === "ok" ? "connected" : s.status === "saved" ? "saved" : (s.provider || "");
  tag.textContent = tagText;
  tag.hidden = !tagText;
  tag.classList.toggle("ok", s.status === "ok" || s.status === "saved");
  $(".server", node).textContent = s.needsDiscovery ? "Looking up server settings ..." : serverLine(s);
  const hint = $(".hint", node);
  if (s.hint) {
    hint.hidden = false;
    hint.textContent = s.hint + " ";
    if (s.help_url) hint.append(el("a", { href: s.help_url, target: "_blank", rel: "noopener noreferrer", text: "Open the setting" }));
  }
  const remove = $(".remove", node);
  remove.hidden = !(unsupported || s.added);
  remove.setAttribute("aria-label", `${s.email} entfernen`);
  remove.addEventListener("click", () => {
    S.imap = S.imap.filter((x) => x !== s);
    renderSources();
    renderRoute();
    announce(`${s.email} entfernt.`);
    $("#add-email").focus();
  });
  if (unsupported) {
    $(".pw-line", node).hidden = true;
    $(".manual", node).hidden = true;
    $(".sig", node).hidden = true;
    return node;
  }

  const pw = $(".pw", node);
  const pwId = "pw-" + s.email.replace(/[^a-z0-9]/gi, "-");
  const statusId = pwId + "-status";
  pw.id = pwId;
  $(".pw-label", node).setAttribute("for", pwId);
  $(".pw-label", node).textContent = `Password for ${s.email}`;
  pw.value = s.password || "";
  pw.placeholder = s.has_password ? "Saved, retype to change it" : "Password";
  pw.addEventListener("input", () => { s.password = pw.value; if (s.status === "ok") s.status = ""; });
  pw.addEventListener("keydown", (e) => { if (e.key === "Enter") { e.preventDefault(); testImap(s, node); } });
  const status = $(".row-status", node);
  status.id = statusId;
  pw.setAttribute("aria-describedby", statusId);
  if (s.status === "err") pw.setAttribute("aria-invalid", "true");
  $(".test", node).addEventListener("click", () => testImap(s, node));
  $(".test", node).disabled = s.status === "busy";
  setStatus(status, s.message, s.status === "ok" || s.status === "saved" ? "ok" : s.status === "err" ? "err" : s.status === "busy" ? "busy" : "");

  const sig = $(".sig-text", node);
  const sigId = pwId + "-sig";
  sig.id = sigId;
  $(".sig-label", node).setAttribute("for", sigId);
  sig.value = s.signature || "";
  if (s.signature) $(".sig", node).open = true;
  sig.addEventListener("input", () => { s.signature = sig.value; });

  const m = s.imap || {};
  $(".m-host", node).value = m.host || "";
  $(".m-port", node).value = m.port || "";
  $(".m-sec", node).value = m.security || "ssl";
  $(".m-user", node).value = m.username || s.email;
  $(".m-shost", node).value = (s.smtp && s.smtp.host) || "";
  $(".m-sport", node).value = (s.smtp && s.smtp.port) || "";
  if (s.openManual) $(".manual", node).open = true;
  const syncManual = () => {
    const host = $(".m-host", node).value.trim();
    if (!host) return;
    const sec = $(".m-sec", node).value;
    s.imap = { host, port: parseInt($(".m-port", node).value, 10) || (sec === "ssl" ? 993 : 143), security: sec, username: $(".m-user", node).value.trim() || s.email };
    const shost = $(".m-shost", node).value.trim();
    if (shost) {
      const sport = parseInt($(".m-sport", node).value, 10) || 465;
      s.smtp = { host: shost, port: sport, security: sport === 465 ? "ssl" : "starttls", username: s.imap.username };
    }
    s.needsDiscovery = false;
    s.status = "";
    $(".server", node).textContent = serverLine(s);
  };
  node.querySelectorAll(".manual input, .manual select").forEach((i) => i.addEventListener("change", syncManual));
  return node;
}

function renderApps() {
  const root = $("#apps");
  root.replaceChildren();
  for (const a of S.apps) {
    const done = a.registered || S.registered.includes(a.id);
    const tagText = !a.installed ? "not found" : done ? (a.detail === "installed as an extension" ? "as an extension" : "registered") : "found";
    root.append(el("li", { class: "row" + (a.installed ? "" : " off") },
      el("label", { class: "check" },
        el("input", { type: "checkbox", checked: a.picked || done, disabled: !a.installed || done, onchange: (e) => { a.picked = e.target.checked; updateAppButton(); } }),
        el("span", { class: "box", "aria-hidden": "true" }),
        el("span", { text: a.name })),
      el("span", { class: "tag" + (done ? " ok" : ""), text: tagText }),
      a.message ? el("p", { class: "row-detail status " + (a.ok ? "ok" : "err"), text: a.message }) : null));
  }
  updateAppButton();
}

function updateAppButton() {
  const picked = S.apps.filter((a) => a.picked && a.installed && !(a.registered || S.registered.includes(a.id)));
  const btn = $("#next-3");
  btn.textContent = !picked.length ? "Continue" : picked.length === 1 ? `Register with ${picked[0].name}` : `Register with ${picked.length} apps`;
  $("#skip-3").hidden = !picked.length;
}

function restartText() {
  const ids = connectedApps().map((a) => a.id);
  if (S.skipped || !ids.length) {
    return "You did not register any AI app. In Claude Desktop a double-click on the post.mcpb file works too, in other apps the command post serve. After that you can ask, for example:";
  }
  const parts = [];
  if (ids.includes("claude_desktop")) {
    parts.push(isWin()
      ? "Quit Claude Desktop completely: right-click the Claude icon in the notification area at the bottom right, choose Quit, then open it again."
      : "Quit Claude Desktop completely and open it again.");
  }
  if (ids.includes("claude_code") || ids.includes("codex")) parts.push("Start a new session in Claude Code or Codex.");
  return parts.join(" ") + " Then you can ask, for example:";
}

const EXAMPLES = [
  "Which mailboxes can you see?",
  "Summarise my unread mail from today.",
  "What did the last customer who wrote to me want? Draft a reply and open it in Outlook.",
  "Find every mail about quotes from the last 14 days and list the open points.",
];

function promptList() {
  const list = el("ul", { class: "prompts" });
  for (const text of EXAMPLES) {
    const btn = el("button", { type: "button", class: "btn small ghost", text: "Kopieren", "aria-label": `Kopieren: ${text}` });
    btn.addEventListener("click", async () => {
      try { await navigator.clipboard.writeText(text); btn.textContent = "Kopiert"; announce("Kopiert."); }
      catch (e) { btn.textContent = "Please select it"; }
      setTimeout(() => (btn.textContent = "Kopieren"), 1600);
    });
    list.append(el("li", {}, el("q", { text }), btn));
  }
  return list;
}

function renderPrompts() {
  $("#prompts").replaceWith(Object.assign(promptList(), { id: "prompts" }));
  let text = restartText();
  if (S.state && S.state.mode === "embedded") {
    text = "New mailboxes work right away. Changed permissions apply after you restart your AI app. Try for example:";
  }
  $("#done-text").textContent = (S.installed ? S.installed + " " : "") + text;
}

/* ── Aktionen ────────────────────────────────────── */

async function testOutlook() {
  const o = S.outlook;
  o.status = "busy";
  o.message = "Connecting to Outlook. The first time, Outlook starts in the background, which can take a moment. If Outlook asks for permission for a program, please allow it.";
  renderSources();
  try {
    const res = await api("/api/test-outlook", {});
    const unpicked = new Set(o.accounts.filter((a) => !a.picked).map((a) => a.email.toLowerCase()));
    o.accounts = res.accounts.map((a) => ({ email: a.email, kind: (a.detail || "").replace("Outlook, ", ""), picked: !unpicked.has(a.email.toLowerCase()) }));
    o.status = "ok";
    o.message = res.message;
    announce("Outlook verbunden. " + res.message);
  } catch (e) {
    o.status = "err";
    o.message = e.message;
    announce("Outlook: " + e.message);
  }
  renderSources();
  renderRoute();
  return o.status === "ok";
}

function rowFor(s) {
  return [...document.querySelectorAll(".imap-row")].find((r) => r.dataset.email === s.email);
}

async function testImap(s, node) {
  node = node && node.isConnected ? node : rowFor(s);
  if (!s.imap) { s.status = "err"; s.message = "Please fill in the server settings first."; s.openManual = true; refreshRow(s, node); return false; }
  if (!s.password && !s.has_password) { s.status = "err"; s.message = "Please enter the password."; refreshRow(s, node); return false; }
  s.status = "busy";
  s.message = "Logging in ...";
  refreshRow(s, node);
  try {
    const res = await api("/api/test", { email: s.email, imap: s.imap, password: s.password || "" });
    s.status = "ok";
    s.message = res.message;
    announce(`${s.email} verbunden. ${res.message}`);
  } catch (e) {
    s.status = "err";
    s.message = e.message;
    announce(`${s.email}: ${e.message}`);
  }
  refreshRow(s, rowFor(s));
  return s.status === "ok";
}

function refreshRow(s, node) {
  if (!node || !node.isConnected) { renderSources(); return; }
  const fresh = imapRow(s);
  node.replaceWith(fresh);
  if (s.status === "err" && $(".pw", fresh)) $(".pw", fresh).focus();
}

async function discover(email, statusNode, quiet = false) {
  if (!quiet) setStatus(statusNode, "Suche Servereinstellungen …", "busy");
  try {
    const d = await api("/api/discover", { email });
    const existing = S.imap.find((x) => x.email.toLowerCase() === email.toLowerCase());
    const entry = existing || { email, label: "", picked: true, status: "", message: "", added: true };
    const supported = d.supported !== false;
    const provider = d.provider && d.provider !== email.split("@")[1] ? d.provider : "";
    Object.assign(entry, {
      imap: d.imap, smtp: d.smtp, hint: d.hint, help_url: d.help_url, supported,
      provider: provider || entry.provider || "", has_password: d.has_password, needsDiscovery: false,
      picked: supported && Boolean(d.imap) && (existing ? existing.picked || !quiet : true),
    });
    if (!existing) S.imap.push(entry);
    if (supported && !d.imap) {
      entry.picked = true;
      entry.status = "err";
      entry.message = "No server settings found. Please fill them in under Server settings below.";
      entry.openManual = true;
    }
    if (!quiet) {
      const msg = !supported ? "This mailbox does not work over IMAP; the details are above." : d.imap ? `Found: ${provider || email.split("@")[1]}` : "";
      setStatus(statusNode, msg, supported ? "ok" : "err");
      $("#add-email").value = "";
    }
    renderSources();
    renderRoute();
    const row = rowFor(entry);
    if (row && !quiet) {
      const pw = $(".pw", row);
      if (pw && supported) pw.focus();
    }
  } catch (e) {
    setStatus(statusNode, e.message, "err");
  }
}

async function nextFromSources() {
  const status = $("#status-1");
  const btn = $("#next-1");
  const pickedImap = S.imap.filter((s) => s.picked && s.supported !== false);
  const outlookPicked = S.outlook && S.outlook.picked;
  if (!outlookPicked && !pickedImap.length) {
    setStatus(status, "Please select at least one mailbox.", "err");
    return;
  }
  btn.disabled = true;
  setStatus(status, "Checking connections ...", "busy");
  let ok = true;
  let msg = "";
  if (outlookPicked) {
    if (S.outlook.status !== "ok" && isWin()) ok = (await testOutlook()) && ok;
    if (S.outlook.status === "ok" && !S.outlook.accounts.some((a) => a.picked)) {
      ok = false;
      msg = "Please select at least one Outlook account.";
    }
  }
  for (const s of pickedImap) {
    if (s.status === "ok" || (s.status === "saved" && !s.password)) continue;
    ok = (await testImap(s, rowFor(s))) && ok;
  }
  btn.disabled = false;
  if (!ok) {
    setStatus(status, msg || "At least one connection does not work yet. What is missing is shown next to the mailbox.", "err");
    return;
  }
  setStatus(status, "");
  goto(2);
}

async function savePermissions() {
  const status = $("#status-2");
  const btn = $("#next-2");
  const sources = [];
  const passwords = {};
  if (S.outlook && S.outlook.picked) {
    const all = S.outlook.accounts;
    const picked = all.filter((a) => a.picked).map((a) => a.email);
    sources.push({ kind: "outlook", outlook_accounts: picked.length === all.length ? null : picked });
  }
  for (const s of S.imap.filter((x) => x.picked && x.supported !== false)) {
    sources.push({ kind: "imap", email: s.email, label: s.label || "", imap: s.imap, smtp: s.smtp, signature: s.signature || "" });
    if (s.password) passwords[s.email] = s.password;
  }
  btn.disabled = true;
  setStatus(status, "Speichere …", "busy");
  try {
    await api("/api/save", {
      sources,
      passwords,
      permissions: { modify: $("#perm-modify").checked, send: $("#perm-send").checked },
    });
    S.imap.forEach((s) => { if (s.picked) { s.has_password = true; s.password = ""; } });
    S.saved = true;
    setStatus(status, "");
    announce("Gespeichert.");
    goto(3);
  } catch (e) {
    setStatus(status, e.message, "err");
  } finally {
    btn.disabled = false;
  }
}

async function registerApps(skip) {
  const status = $("#status-3");
  S.skipped = Boolean(skip) && !connectedApps().length;
  const ids = skip ? [] : S.apps.filter((a) => a.picked && a.installed && !(a.registered || S.registered.includes(a.id))).map((a) => a.id);
  if (ids.length) {
    setStatus(status, "Registering ...", "busy");
    try {
      const res = await api("/api/register", { apps: ids });
      let allOk = true;
      for (const r of res.results) {
        const app = S.apps.find((a) => a.id === r.app);
        if (app) { app.message = r.message; app.ok = r.ok; app.picked = false; }
        if (r.ok) S.registered.push(r.app); else allOk = false;
      }
      renderApps();
      setStatus(status, res.installed || "", res.installed ? "ok" : "");
      if (res.installed) S.installed = res.installed;
      if (!allOk) { setStatus(status, "Not everything worked, see above.", "err"); renderRoute(); return; }
    } catch (e) {
      setStatus(status, e.message, "err");
      return;
    }
  }
  goto(4);
}

async function finish() {
  try { await api("/api/finish", {}); } catch (e) { /* the server may already be gone */ }
  const closed = el("section", { class: "closed" },
    el("h2", { tabindex: "-1", text: "Alles eingerichtet." }),
    el("p", { class: "lede", text: restartText() }),
    promptList(),
    el("p", { class: "help", text: "You can close this window now. You can start setup again at any time by asking your AI to set up post., or with the command post setup." }));
  document.querySelector(".steps").replaceWith(closed);
  $("h2", closed).focus();
  renderRoute();
}

/* ── Start ───────────────────────────────────────── */

async function init() {
  $("#add-form").addEventListener("submit", (e) => {
    e.preventDefault();
    const email = $("#add-email").value.trim();
    if (email) discover(email, $("#add-status"));
  });
  $("#next-1").addEventListener("click", nextFromSources);
  $("#next-2").addEventListener("click", savePermissions);
  $("#next-3").addEventListener("click", () => registerApps(false));
  $("#skip-3").addEventListener("click", () => registerApps(true));
  $("#finish").addEventListener("click", finish);
  document.querySelectorAll("[data-edit]").forEach((b) => b.addEventListener("click", () => goto(Number(b.dataset.edit))));
  $("#perm-modify").addEventListener("change", () => renderSteps());
  $("#perm-send").addEventListener("change", () => renderSteps());

  renderSteps();
  try {
    S.state = await api("/api/state");
  } catch (e) {
    setStatus($("#scan"), e.message, "err");
    return;
  }
  $("#version").textContent = S.state.version;
  buildModel(S.state);
  const found = (S.outlook ? 1 : 0) + S.imap.length;
  $("#scan").textContent = found
    ? "This is what post. found on this computer. Choose what your AI may see."
    : "";
  renderSources();
  renderSteps();
  for (const s of S.imap.filter((x) => x.needsDiscovery)) {
    discover(s.email, el("span"), true);
  }
}

init();
