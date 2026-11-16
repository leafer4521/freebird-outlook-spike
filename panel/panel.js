// FreeBird for Outlook - the panel. Talks only to the bridge on this PC (bridge\Bridge.cs), never to any server:
// the e-mail goes from Outlook straight to the PC.
"use strict";

const BRIDGE = window.FBO_BRIDGE || "http://127.0.0.1:8791";      // FBO_BRIDGE: addin\test\harness.html (a test helper on another port)
const KEY_NAME = "fbo.bridgeKey";
const SET_NAME = "fbo.settings";
const TEXT_LIMIT = 20000;           // characters of the body sent for "is this client named in the e-mail?"
const POLL_MS = 1000, POLL_MAX_MS = 10 * 60 * 1000;
const FILE_MAX = 25 * 1024 * 1024;  // FreeBird attaches files up to 25 MB (attach-file)

// Settings (Ben, 2026-11-16): fewer clicks by default - "Attach" files at once (the click is the approval; FreeBird's plan is still
// shown when FreeBird warns, e.g. the same filing was already done today) - and the e-mail goes on the client unless the person
// turns on choosing a policy.
const DEFAULTS = { showPlan: false, choosePolicy: false };

const S = {};                       // the current e-mail and what has been chosen for it
let runId = 0;                      // a newer e-mail (ItemChanged) cancels the views of an older one

// ------------------------------------------------------------------ small helpers

function h(tag, attrs, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === null || v === undefined || v === false) continue;
    if (k === "class") e.className = v;
    else if (k.startsWith("on")) e.addEventListener(k.slice(2), v);
    else if (k === "value") e.value = v;
    else if (k === "checked") e.checked = !!v;
    else e.setAttribute(k, v === true ? "" : v);
  }
  for (const c of kids.flat()) if (c !== null && c !== undefined && c !== false) e.append(c instanceof Node ? c : String(c));
  return e;
}
function show(...nodes) { const app = document.getElementById("app"); app.replaceChildren(...nodes.flat(2).filter(n => n !== null && n !== undefined && n !== false)); window.scrollTo(0, 0); }
function busy(text, sub) { S.view = "busy"; show(h("div", { class: "center" }, h("div", { class: "spin", "aria-hidden": "true" }), h("p", {}, text), sub ? h("p", { class: "muted small" }, sub) : null)); }
function note(kind, ...kids) { return h("div", { class: "note " + kind }, ...kids); }
function getKey() { try { return localStorage.getItem(KEY_NAME) || S.key || ""; } catch (e) { return S.key || ""; } }
function setKey(k) { S.key = k; try { localStorage.setItem(KEY_NAME, k); } catch (e) { } }
function dropKey() { S.key = ""; try { localStorage.removeItem(KEY_NAME); } catch (e) { } }
let memSettings = null;
function settings() {
  if (memSettings) return memSettings;
  let s = {}; try { s = JSON.parse(localStorage.getItem(SET_NAME) || "{}") || {}; } catch (e) { }
  return memSettings = Object.assign({}, DEFAULTS, s);
}
function setSetting(k, v) { settings()[k] = v; try { localStorage.setItem(SET_NAME, JSON.stringify(memSettings)); } catch (e) { } }
function kb(n) { return n >= 1048576 ? (n / 1048576).toFixed(1) + " MB" : Math.max(1, Math.round(n / 1024)) + " KB"; }
function plural(n, one, many) { return n + " " + (n === 1 ? one : many); }
function b64Text(s) { const b = new TextEncoder().encode(s); let t = ""; for (let i = 0; i < b.length; i += 0x8000) t += String.fromCharCode.apply(null, b.subarray(i, i + 0x8000)); return btoa(t); }

class BridgeDown extends Error { }
class NeedsPair extends Error { }

async function api(path, body) {
  let r;
  try {
    r = await fetch(BRIDGE + path, {
      method: body === undefined ? "GET" : "POST",
      headers: Object.assign({ "Content-Type": "application/json" }, getKey() ? { "X-FBO-Key": getKey() } : {}),
      body: body === undefined ? undefined : JSON.stringify(body)
    });
  } catch (e) { throw new BridgeDown(String(e)); }
  let j = null; try { j = await r.json(); } catch (e) { }
  if (r.status === 401 || (j && j.code === "pair")) throw new NeedsPair();
  if (!j) throw new Error("The FreeBird helper gave an empty answer (HTTP " + r.status + ").");
  return j;
}

function officeAsync(fn) {
  return new Promise((resolve, reject) => fn((res) => res.status === Office.AsyncResultStatus.Succeeded ? resolve(res.value) : reject(new Error(res.error && res.error.message || "Outlook could not do that."))));
}

// ------------------------------------------------------------------ start

Office.onReady(() => {
  try {
    const t = Office.context.officeTheme;
    if (t && t.bodyBackgroundColor) {
      const c = t.bodyBackgroundColor.replace("#", ""); const n = parseInt(c.length === 3 ? c.split("").map(x => x + x).join("") : c, 16);
      const lum = ((n >> 16) & 255) * 0.299 + ((n >> 8) & 255) * 0.587 + (n & 255) * 0.114;
      document.documentElement.setAttribute("data-theme", lum < 128 ? "dark" : "light");
    }
  } catch (e) { }
  try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => start()); } catch (e) { }
  start();
});

// The e-mail's attachments: files (photos, PDFs ...) and attached e-mails can be handed over (Mailbox 1.8); cloud links cannot.
// Pictures inside the message body (logos, signatures - or a pasted photo) are listed but not ticked.
function readAttachments(item) {
  let canRead = false;
  try { canRead = Office.context.requirements.isSetSupported("Mailbox", "1.8") && typeof item.getAttachmentContentAsync === "function"; } catch (e) { }
  return (item.attachments || []).map(a => {
    const type = String(a.attachmentType || "file").toLowerCase();
    let name = a.name || "attachment";
    if (type === "item" && !/\.(eml|msg)$/i.test(name)) name += ".eml";
    let why = "";
    if (type === "cloud") why = "a cloud link, not a file";
    else if (!canRead) why = "this Outlook can't hand it over";
    else if (a.size > FILE_MAX) why = "over HawkSoft's 25 MB";
    return { id: a.id, name, size: a.size || 0, inline: !!a.isInline, type, why };
  });
}

async function start() {
  const my = ++runId;
  for (const k of Object.keys(S)) if (k !== "key") delete S[k];
  const item = Office.context.mailbox.item;
  if (!item) { show(h("p", { class: "muted" }, "Open an e-mail to file it in HawkSoft.")); return; }
  S.from = item.from && item.from.emailAddress || "";
  S.fromName = item.from && item.from.displayName || "";
  S.subject = item.subject || "";
  S.extra = {};
  S.pols = {};
  S.atts = readAttachments(item);
  S.pick = new Set(S.atts.filter(a => !a.why && !a.inline).map(a => a.id));
  busy("Checking FreeBird…");
  try {
    S.text = (await officeAsync(cb => item.body.getAsync(Office.CoercionType.Text, cb)) || "").slice(0, TEXT_LIMIT);
  } catch (e) { S.text = ""; }
  if (my !== runId) return;
  await checkAndMatch(my);
}

async function checkAndMatch(my) {
  let st;
  try { st = await api("/status", {}); }
  catch (e) { if (my !== runId) return; return e instanceof NeedsPair ? viewPair() : e instanceof BridgeDown ? viewBridgeDown() : viewError(e.message, () => start()); }
  if (my !== runId) return;
  document.getElementById("testBadge").hidden = !st.testOnly;
  S.testOnly = st.testOnly; S.testPrefix = st.testPrefix;
  if (!st.freebird || !st.freebird.up) return viewFreeBirdDown(st.freebird && st.freebird.message);
  if (st.freebird.state && st.freebird.state !== "ok") S.health = st.freebird.message;
  await match({ from: S.from }, my);
}

async function match(q, my) {
  busy(q.query ? "Searching HawkSoft for “" + q.query + "”…" : "Looking up " + (S.from || "the sender") + " in HawkSoft…", "FreeBird checks HawkSoft's client list. Nothing is changed.");
  let m;
  try { m = await api("/match", Object.assign({ fromName: S.fromName, subject: S.subject, text: S.text }, q)); }
  catch (e) { if (my !== runId) return; return handle(e, () => match(q, my)); }
  if (my !== runId) return;
  if (!m.ok) return viewError(m.error, () => match(q, my));
  S.match = m; S.showList = false;
  choose(m.suggested ? m.matches.find(x => x.id === m.suggested) : null);
  viewChoose();
}

function choose(c) { if (!S.chosen || !c || S.chosen.id !== c.id) S.policy = ""; S.chosen = c; S.extra = {}; }

function handle(e, retry) {
  if (e instanceof NeedsPair) return viewPair();
  if (e instanceof BridgeDown) return viewBridgeDown();
  return viewError(e.message, retry);
}

// ------------------------------------------------------------------ views

function mailCard() {
  return h("div", { class: "mail" },
    h("div", { class: "from" }, S.fromName && S.fromName !== S.from ? S.fromName + " · " + S.from : (S.from || "Unknown sender")),
    h("div", { class: "subj" }, S.subject || "(no subject)"));
}

function picked() { return S.atts.filter(a => S.pick.has(a.id)); }
function polLabel(p) { return [p.tab || p.policy, p.number, p.effective && p.expiration ? p.effective + "–" + p.expiration : ""].filter(Boolean).join(" · "); }
function whereText() {
  const p = S.policy && (S.pols[S.chosen.id] || []).find(x => x.policy === S.policy);
  return S.chosen.name + (p ? ", policy " + (p.tab || p.policy) : "");
}

function viewChoose() {
  S.view = "choose";
  const m = S.match, set = settings();
  const desc = h("input", { type: "text", id: "desc", maxlength: "100", value: S.desc !== undefined ? S.desc : (S.subject || ("E-mail from " + (S.fromName || S.from))).slice(0, 100) });
  const count = h("div", { class: "count" });
  const upd = () => { count.textContent = desc.value.length + " / 100"; };
  desc.addEventListener("input", upd); upd();
  const noteBox = h("textarea", { id: "note" }); noteBox.value = S.note !== undefined ? S.note : "E-mail from " + (S.fromName || S.from) + (S.subject ? ": " + S.subject : "") + ".";
  const keep = () => { S.desc = desc.value; S.note = noteBox.value; };
  const redraw = () => { keep(); viewChoose(); };

  let listPart;
  if (!m.matches.length) {
    listPart = [note("warn", h("p", {}, m.typed ? "No client in HawkSoft matches “" + m.query + "”." : "No client in HawkSoft has the address " + S.from + "."), h("p", { class: "small" }, "Search by the client's name instead:")), searchBox()];
  } else {
    const total = m.total || m.matches.length;
    const intro = total === 1
      ? (m.typed ? "This client matches your search:" : "This client has the sender's address:")
      : total + " clients match" + (m.typed ? " your search" : " the sender") + ". " + (m.suggested ? "FreeBird picked the one the e-mail names. Check it, or choose another." : "Which one is this e-mail about?");
    // once a client is chosen the list folds to that one, so the description and note are right below it (live 2026-11-16: a long list hid them)
    const folded = S.chosen && !S.showList;
    const list = folded ? [S.chosen] : m.matches;
    const cards = list.map(c => {
      const radio = h("input", { type: "radio", name: "client", value: c.id, checked: S.chosen && S.chosen.id === c.id, disabled: c.blocked });
      const card = h("label", { class: "client" + (S.chosen && S.chosen.id === c.id ? " on" : "") + (c.blocked ? " off" : "") },
        radio,
        h("div", {},
          h("div", { class: "name" }, c.name,
            c.id === m.suggested ? h("span", { class: "badge pick" }, "Picked") : null,
            c.named ? h("span", { class: "badge named" }, "Named in e-mail") : null),
          h("div", { class: "meta" }, [c.type, c.status, c.address].filter(Boolean).join(" · ") + " · ID " + c.id),
          c.blocked ? h("div", { class: "meta" }, "Test mode: only “" + S.testPrefix + "” clients") : null));
      radio.addEventListener("change", () => { choose(c); S.showList = false; redraw(); });
      return card;
    });
    listPart = folded
      ? [h("div", { class: "clients" }, cards),
         h("button", { class: "link", onclick: () => { keep(); S.showList = true; viewChoose(); } }, total > 1 ? "Change client (" + total + " matched)" : "Change client")]
      : [h("p", {}, intro), m.suggested && m.why ? h("p", { class: "muted small" }, "Why: " + m.why + ".") : null,
         m.more ? note("warn", h("p", {}, "Showing " + m.matches.length + " of " + total + ". Type more of the name to narrow it down:")) : null,
         m.more ? searchBox(m.typed ? m.query : "") : null,
         h("div", { class: "clients", role: "radiogroup" }, cards),
         m.more ? null : h("button", { class: "link", onclick: (ev) => { ev.target.replaceWith(searchBox()); } }, "Not listed? Search HawkSoft by name")];
  }

  // File on: the client (default) or one of its policies - only with the setting on
  let polPart = null;
  if (set.choosePolicy && S.chosen) {
    const pols = S.pols[S.chosen.id];
    if (pols === undefined) { loadPolicies(S.chosen, redraw); polPart = [h("h3", {}, "File on"), h("p", { class: "muted small" }, "Reading the client's policies…")]; }
    else if (pols.error) polPart = [h("h3", {}, "File on"), note("warn", h("p", { class: "small" }, "Couldn't read the policies, so it goes on the client: " + pols.error))];
    else {
      const sel = h("select", { id: "policy", "aria-label": "File on" },
        h("option", { value: "" }, "The client (no policy)"),
        pols.map(p => h("option", { value: p.policy }, polLabel(p) + (p.status && p.status !== "Active" ? " (" + p.status + ")" : ""))));
      sel.value = S.policy || "";
      sel.addEventListener("change", () => { S.policy = sel.value; S.extra = {}; redraw(); });
      polPart = [h("h3", {}, "File on"), sel, pols.length ? null : h("p", { class: "muted small" }, "This client has no policies in HawkSoft.")];
    }
  }

  // the attachments: ticked = filed too, each as its own attachment with its own log
  let attPart = null;
  if (S.atts.length) {
    attPart = [h("h3", {}, "Attachments"),
      h("div", { class: "atts" }, S.atts.map(a => {
        const box = h("input", { type: "checkbox", checked: S.pick.has(a.id), disabled: !!a.why });
        box.addEventListener("change", () => { if (box.checked) S.pick.add(a.id); else S.pick.delete(a.id); redraw(); });
        return h("label", { class: "att" + (a.why ? " off" : "") }, box,
          h("div", {}, h("div", { class: "name" }, a.name), h("div", { class: "meta" }, [a.size ? kb(a.size) : "", a.why || (a.inline ? "in the message body" : "")].filter(Boolean).join(" · "))));
      }))];
  }

  const n = picked().length;
  const what = "the e-mail" + (n ? " and " + plural(n, "attachment", "attachments") : "");
  const go = h("button", { class: "primary", disabled: !S.chosen || (set.choosePolicy && S.chosen && S.pols[S.chosen.id] === undefined), onclick: () => { S.desc = desc.value.trim(); S.note = noteBox.value.trim(); review(); } },
    set.showPlan ? "Review and file" : "Attach");
  const hint = !S.chosen ? "Choose the client first."
    : set.showPlan ? "You see FreeBird's plan for " + what + " before anything is filed."
    : "Pressing Attach files " + what + " on " + whereText() + " right away. FreeBird reads it back from HawkSoft.";

  show(mailCard(),
    S.health ? note("warn", h("p", {}, S.health)) : null,
    h("h3", {}, "Client"), listPart,
    polPart,
    attPart,
    h("h3", {}, "Filed as"),
    h("label", { class: "f", for: "desc" }, "E-mail's attachment description"), desc, count,
    h("label", { class: "f", for: "note" }, "Log note"), noteBox,
    h("p", { class: "muted small" }, "Logged in HawkSoft as an e-mail from the insured, category Emails."),
    h("p", { class: "small hint" }, hint),
    go,
    settingsPart(redraw));
}

function settingsPart(redraw) {
  const set = settings();
  if (!S.showSettings) return h("button", { class: "link quiet", onclick: () => { S.showSettings = true; redraw(); } }, "Settings");
  const row = (k, label, sub) => {
    const box = h("input", { type: "checkbox", checked: !!set[k] });
    box.addEventListener("change", () => { setSetting(k, box.checked); redraw(); });
    return h("label", { class: "att" }, box, h("div", {}, h("div", { class: "name" }, label), h("div", { class: "meta" }, sub)));
  };
  return h("div", { class: "settings" },
    h("h3", {}, "Settings"),
    row("choosePolicy", "Choose a policy", "Off: files go on the client. On: pick one of the client's policies."),
    row("showPlan", "Show FreeBird's plan first", "Off: Attach files at once. On: you confirm FreeBird's plan."),
    h("button", { class: "link quiet", onclick: () => { S.showSettings = false; redraw(); } }, "Close settings"));
}

async function loadPolicies(c, redraw) {
  if (S.pols[c.id] !== undefined || S.polLoading === c.id) return;
  S.polLoading = c.id; const my = runId;
  let r;
  try { r = await api("/policies", { clientId: c.id, clientName: c.name }); }
  catch (e) { if (e instanceof NeedsPair || e instanceof BridgeDown) { S.polLoading = null; return handle(e); } r = { ok: false, error: e.message }; }
  if (my !== runId) return;
  S.polLoading = null;
  S.pols[c.id] = r.ok ? (r.policies || []) : { error: r.error || "no answer" };
  if (S.view === "choose" && S.chosen && S.chosen.id === c.id) redraw();
}

function searchBox(initial) {
  const q = h("input", { type: "text", placeholder: "Client name, phone or policy number", "aria-label": "Search HawkSoft", value: initial || null });
  const run = () => { if (q.value.trim()) match({ query: q.value.trim() }, runId); };
  q.addEventListener("keydown", (e) => { if (e.key === "Enter") run(); });
  return h("div", { class: "row" }, q, h("button", { class: "secondary", onclick: run }, "Search"));
}

// The e-mail (.eml) and the ticked attachments go to the helper once; a change of the ticks stages them again.
async function stage() {
  const item = Office.context.mailbox.item, files = picked(), key = files.map(a => a.id).join("|");
  if (S.draftId && S.stagedKey === key) return true;
  busy("Reading the e-mail…", files.length ? "and " + plural(files.length, "attachment", "attachments") : null);
  const eml = await officeAsync(cb => item.getAsFileAsync(cb));
  const attachments = [];
  for (const a of files) {
    const c = await officeAsync(cb => item.getAttachmentContentAsync(a.id, cb));
    const fmt = String(c.format || "").toLowerCase();
    if (fmt === "base64") attachments.push({ name: a.name, base64: c.content });
    else if (fmt === "eml" || fmt === "icalendar") attachments.push({ name: a.name, base64: b64Text(c.content) });
  }
  const sg = await api("/stage", { emlBase64: eml, subject: S.subject, attachments });
  if (!sg.ok) { viewError(sg.error, review); return false; }
  S.draftId = sg.draftId; S.stagedKey = key; S.extra = {};
  S.taskNames = ["the e-mail"].concat(files.map(a => "“" + a.name + "”"));
  return true;
}

async function review() {
  const my = runId, quick = !settings().showPlan;
  try {
    if (!(await stage())) return;
    if (my !== runId) return;
    busy(quick ? "Filing in HawkSoft…" : "Asking FreeBird for its plan…", quick ? "Getting FreeBird ready." : "Nothing is filed yet.");
    const pl = await api("/plan", { draftId: S.draftId, clientId: S.chosen.id, clientName: S.chosen.name, policy: S.policy || "", from: S.from, fromName: S.fromName, subject: S.subject, description: S.desc, note: S.note, extra: S.extra });
    if (my !== runId) return;
    if (!pl.ok) {
      if (pl.code === "restage") { S.draftId = null; return review(); }
      return viewError(pl.error + (pl.missing ? " (" + JSON.stringify(pl.missing) + ")" : ""), review, viewChoose);
    }
    if (pl.jobIds) return filing(pl.jobIds);
    // quick: the person's "Attach" click is the approval - unless FreeBird warns (the same filing was already done today): then its plan is shown
    if (quick && !pl.earlierRun) return approve(pl);
    viewConfirm(pl);
  } catch (e) { if (my === runId) handle(e, review); }
}

async function approve(pl) {
  try {
    const ap = await api("/approve", { pendingId: pl.pendingId });
    if (!ap.ok) return ap.code === "replan" ? review() : viewError(ap.error, review, viewChoose);
    filing(ap.jobIds || [ap.jobId], ap.duplicate ? ap.note : null);
  } catch (e) { handle(e, review); }
}

function planLines(plan) {
  const l = Array.isArray(plan) ? plan : (plan ? [plan] : []);
  return l.flatMap(x => typeof x === "string" ? x.split("\n") : [JSON.stringify(x)]).map(x => x.trim()).filter(Boolean);
}

function viewConfirm(pl) {
  S.view = "confirm";
  const yes = h("button", { class: "primary", onclick: () => { yes.disabled = true; approve(pl); } }, "Yes, file it");
  // a batch plan starts "4 tasks, one approval for all of them ...:" and numbers its tasks itself
  let lines = planLines(pl.plan), intro = null;
  if (lines.length > 1 && /:$/.test(lines[0])) intro = h("p", { class: "small" }, lines.shift());
  lines = lines.map(l => l.replace(/^\d+[.)]\s+/, ""));
  show(mailCard(),
    pl.earlierRun ? note("warn", h("p", {}, pl.earlierRun)) : null,
    h("h3", {}, "FreeBird will"), intro,
    h("ol", { class: "plan" }, lines.map(l => h("li", {}, l))),
    h("div", { class: "actions" }, h("button", { class: "secondary", onclick: viewChoose }, "Back"), yes));
}

// One job per file (the e-mail first). Each is followed until it ends; FreeBird does them one after the other.
async function filing(jobIds, dupNote) {
  const my = runId, t0 = Date.now(), res = jobIds.map(() => null);
  const names = jobIds.map((_, i) => (S.taskNames && S.taskNames[i]) || ("file " + (i + 1)));
  const sub = "FreeBird is working in HawkSoft on this PC. Leave HawkSoft alone until it's done.";
  busy("Filing in HawkSoft…", sub);
  while (Date.now() - t0 < POLL_MAX_MS) {
    await new Promise(r => setTimeout(r, POLL_MS));
    if (my !== runId) return;
    for (let i = 0; i < jobIds.length; i++) {
      if (res[i]) continue;
      let j;
      try { j = await api("/job/" + encodeURIComponent(jobIds[i])); } catch (e) { if (e instanceof BridgeDown) break; return handle(e, () => filing(jobIds, dupNote)); }
      if (my !== runId) return;
      if (!j.ok) return viewError(j.error, () => filing(jobIds, dupNote));
      if (j.status === "queued" || j.status === "running") break;      // the next ones wait behind this one
      res[i] = j;
    }
    const left = res.findIndex(r => !r);
    if (left < 0) return viewResult(res, names, dupNote);
    if (jobIds.length > 1) busy("Filing " + (left + 1) + " of " + jobIds.length + " in HawkSoft…", names[left].replace(/^the/, "The") + ". " + sub);
  }
  viewError("FreeBird is still working after 10 minutes. Check FreeBird's window on this PC.", () => filing(jobIds, dupNote));
}

function viewResult(res, names, dupNote) {
  if (res.every(r => r.status === "done")) return viewDone(res, dupNote);
  const q = res.findIndex(r => r.status === "needs_input" && r.needsInput);
  if (q >= 0) return viewQuestion(res[q].needsInput, q, res, names);
  viewFailed(res, names);
}

function filedList(res, names) {
  const done = names.filter((_, i) => res[i] && res[i].status === "done");
  return done.length && done.length < res.length ? note("ok", h("p", {}, "Filed: " + done.join(", ") + ".")) : null;
}

function viewDone(res, dupNote) {
  S.view = "done";
  const all = (k) => res.flatMap(r => r[k] || []);
  const created = all("created").filter(c => c && c.kind);
  const decisions = [...new Set(all("decisions"))], needsYou = all("needsYou");
  const verified = res.every(r => r.verified);
  show(h("div", { class: "center" }, h("div", { class: "done-icon", "aria-hidden": "true" }, "✓"),
      h("h2", {}, "Filed on " + whereText()),
      h("p", { class: "muted" }, verified ? "FreeBird checked it in HawkSoft." : "FreeBird filed it but could not read all of it back. Check the client's attachments.")),
    dupNote ? note("warn", h("p", {}, dupNote)) : null,
    created.length ? [h("h3", {}, "In HawkSoft now"), h("ul", { class: "list" }, created.map(c => h("li", {}, (c.kind === "log" ? "Log note: " : c.kind === "attachment" ? "Attachment: " : c.kind + ": ") + c.text)))] : null,
    decisions.length ? [h("h3", {}, "FreeBird decided"), h("ul", { class: "list" }, decisions.map(d => h("li", {}, d)))] : null,
    needsYou.length ? note("warn", h("p", {}, "Still for you:"), h("ul", { class: "list" }, needsYou.map(d => h("li", {}, d)))) : null);
}

// FreeBird asked about one of the files. The answer goes with that file only; filed ones are not filed again (FreeBird reuses them).
function viewQuestion(ni, ix, res, names) {
  S.view = "question";
  const opts = (ni.options || []).map(o => typeof o === "string" ? { value: o, label: o } : o);
  const answer = (v) => { (S.extra[ix] = S.extra[ix] || {})[ni.param] = v; review(); };
  const free = h("input", { type: "text", placeholder: "Your answer" });
  const about = res && res.length > 1 ? "About " + names[ix] + ": " : "";
  show(mailCard(),
    note("warn", h("p", {}, "FreeBird needs one more detail." + (res && res.some(r => r.status === "done") ? "" : " Nothing was filed yet."))),
    res ? filedList(res, names) : null,
    h("h2", {}, about + ni.question),
    ni.hint ? h("p", { class: "muted small" }, ni.hint) : null,
    opts.length ? h("div", { class: "opts" }, opts.map(o => h("button", { class: "secondary", onclick: () => answer(o.value) }, o.label || o.value)))
      : [free, h("button", { class: "primary", onclick: () => { if (free.value.trim()) answer(free.value.trim()); } }, "Continue")],
    h("button", { class: "link", onclick: viewChoose }, "Back"));
}

function viewFailed(res, names) {
  S.view = "failed";
  const bad = res.map((r, i) => ({ r, i })).filter(x => x.r.status !== "done");
  const may = bad.map(x => x.r.maySaved).filter(Boolean);
  const many = res.length > 1;
  show(mailCard(),
    note("bad", h("p", {}, many ? "FreeBird could not file " + bad.map(x => names[x.i]).join(", ") + "." : "FreeBird could not file this e-mail."),
      bad.map(x => h("p", { class: "small" }, (many ? names[x.i].replace(/^the/, "The") + ": " : "FreeBird says: ") + (x.r.error || ("it ended with status " + x.r.status + "."))))),
    filedList(res, names),
    may.length ? note("warn", h("p", {}, "HawkSoft may already have saved " + may.join("; ") + ". Check the client in HawkSoft before trying again."))
      : h("p", { class: "muted small" }, many ? "Those were not saved in HawkSoft. Try again files only what is missing." : "Nothing was saved in HawkSoft."),
    h("div", { class: "actions" }, h("button", { class: "secondary", onclick: viewChoose }, "Back"), may.length ? null : h("button", { class: "primary", onclick: review }, "Try again")));
}

function viewError(msg, retry, back) {
  S.view = "error";
  show(note("bad", h("p", {}, msg || "Something went wrong.")),
    h("div", { class: "actions" }, back ? h("button", { class: "secondary", onclick: back }, "Back") : null, retry ? h("button", { class: "primary", onclick: retry }, "Try again") : null));
}

function viewBridgeDown() {
  S.view = "down";
  show(h("h2", {}, "FreeBird for Outlook isn't running"),
    h("p", {}, "This panel talks to a small FreeBird helper on this PC, and it isn't answering."),
    h("p", { class: "muted small" }, "Start “FreeBird for Outlook” on this PC, then press Try again."),
    h("button", { class: "primary", onclick: start }, "Try again"));
}

function viewFreeBirdDown(msg) {
  S.view = "down";
  show(h("h2", {}, "FreeBird isn't running"),
    h("p", {}, msg || "Open FreeBird from the Start menu, then try again."),
    h("button", { class: "primary", onclick: start }, "Try again"));
}

function viewPair() {
  S.view = "pair";
  dropKey();
  const code = h("input", { type: "text", class: "code", inputmode: "numeric", maxlength: "7", placeholder: "000 000", "aria-label": "Pairing code" });
  const msg = h("div");
  const finish = async () => {
    try {
      const r = await api("/pair/finish", { code: code.value });
      if (!r.ok) { msg.replaceChildren(note("bad", h("p", {}, r.error))); return; }
      setKey(r.key); start();
    } catch (e) { handle(e, viewPair); }
  };
  code.addEventListener("keydown", (e) => { if (e.key === "Enter") finish(); });
  const step2 = h("div", { hidden: true },
    h("label", { class: "f" }, "Type the code shown on your PC"), code, msg,
    h("button", { class: "primary", onclick: finish }, "Connect"));
  const startBtn = h("button", { class: "primary", onclick: async () => {
    try { const r = await api("/pair/start", {}); if (!r.ok) return viewError(r.error, viewPair); startBtn.hidden = true; step2.hidden = false; code.focus(); }
    catch (e) { handle(e, viewPair); }
  } }, "Show a code on my PC");
  show(h("h2", {}, "Connect Outlook to FreeBird"),
    h("p", {}, "One time only: this lets the panel hand e-mails to FreeBird on this PC."),
    h("p", { class: "muted small" }, "A code appears on your PC. Typing it here proves the panel and FreeBird are on the same computer."),
    startBtn, step2);
}
