// FreeBird for Outlook - the panel. Talks only to the bridge on this PC (bridge\Bridge.cs), never to any server:
// the e-mail goes from Outlook straight to the PC.
"use strict";

const BRIDGE = "http://127.0.0.1:8791";
const KEY_NAME = "fbo.bridgeKey";
const TEXT_LIMIT = 20000;           // characters of the body sent for "is this client named in the e-mail?"
const POLL_MS = 1000, POLL_MAX_MS = 5 * 60 * 1000;

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
    else e.setAttribute(k, v === true ? "" : v);
  }
  for (const c of kids.flat()) if (c !== null && c !== undefined && c !== false) e.append(c instanceof Node ? c : String(c));
  return e;
}
function show(...nodes) { const app = document.getElementById("app"); app.replaceChildren(...nodes.flat(2).filter(n => n !== null && n !== undefined && n !== false)); window.scrollTo(0, 0); }
function busy(text, sub) { show(h("div", { class: "center" }, h("div", { class: "spin", "aria-hidden": "true" }), h("p", {}, text), sub ? h("p", { class: "muted small" }, sub) : null)); }
function note(kind, ...kids) { return h("div", { class: "note " + kind }, ...kids); }
function getKey() { try { return localStorage.getItem(KEY_NAME) || S.key || ""; } catch (e) { return S.key || ""; } }
function setKey(k) { S.key = k; try { localStorage.setItem(KEY_NAME, k); } catch (e) { } }
function dropKey() { S.key = ""; try { localStorage.removeItem(KEY_NAME); } catch (e) { } }

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

async function start() {
  const my = ++runId;
  for (const k of Object.keys(S)) if (k !== "key") delete S[k];
  const item = Office.context.mailbox.item;
  if (!item) { show(h("p", { class: "muted" }, "Open an e-mail to file it in HawkSoft.")); return; }
  S.from = item.from && item.from.emailAddress || "";
  S.fromName = item.from && item.from.displayName || "";
  S.subject = item.subject || "";
  S.extra = {};
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
  S.match = m;
  S.chosen = m.suggested ? m.matches.find(x => x.id === m.suggested) : null;
  viewChoose();
}

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

function viewChoose() {
  const m = S.match;
  const desc = h("input", { type: "text", id: "desc", maxlength: "100", value: S.desc !== undefined ? S.desc : (S.subject || ("E-mail from " + (S.fromName || S.from))).slice(0, 100) });
  const count = h("div", { class: "count" });
  const upd = () => { count.textContent = desc.value.length + " / 100"; };
  desc.addEventListener("input", upd); upd();
  const noteBox = h("textarea", { id: "note" }); noteBox.value = S.note !== undefined ? S.note : "E-mail from " + (S.fromName || S.from) + (S.subject ? ": " + S.subject : "") + ".";
  const go = h("button", { class: "primary", disabled: !S.chosen, onclick: () => { S.desc = desc.value.trim(); S.note = noteBox.value.trim(); review(); } }, "Review and file");

  let listPart;
  if (!m.matches.length) {
    listPart = [note("warn", h("p", {}, m.typed ? "No client in HawkSoft matches “" + m.query + "”." : "No client in HawkSoft has the address " + S.from + "."), h("p", { class: "small" }, "Search by the client's name instead:")), searchBox()];
  } else {
    const intro = m.matches.length === 1
      ? (m.typed ? "This client matches your search:" : "This client has the sender's address:")
      : m.matches.length + " clients match" + (m.typed ? " your search" : " the sender") + ". " + (m.suggested ? "FreeBird picked the one the e-mail names. Check it, or choose another." : "Which one is this e-mail about?");
    const cards = m.matches.map(c => {
      const radio = h("input", { type: "radio", name: "client", value: c.id, checked: S.chosen && S.chosen.id === c.id, disabled: c.blocked });
      const card = h("label", { class: "client" + (S.chosen && S.chosen.id === c.id ? " on" : "") + (c.blocked ? " off" : "") },
        radio,
        h("div", {},
          h("div", { class: "name" }, c.name,
            c.id === m.suggested ? h("span", { class: "badge pick" }, "Picked") : null,
            c.named ? h("span", { class: "badge named" }, "Named in e-mail") : null),
          h("div", { class: "meta" }, [c.type, c.status, c.address].filter(Boolean).join(" · ") + " · ID " + c.id),
          c.blocked ? h("div", { class: "meta" }, "Test mode: only “" + S.testPrefix + "” clients") : null));
      radio.addEventListener("change", () => { S.chosen = c; S.desc = desc.value; S.note = noteBox.value; viewChoose(); });
      return card;
    });
    listPart = [h("p", {}, intro), m.suggested && m.why ? h("p", { class: "muted small" }, "Why: " + m.why + ".") : null,
      h("div", { class: "clients", role: "radiogroup" }, cards),
      h("button", { class: "link", onclick: (ev) => { ev.target.replaceWith(searchBox()); } }, "Not listed? Search HawkSoft by name")];
  }

  show(mailCard(),
    S.health ? note("warn", h("p", {}, S.health)) : null,
    h("h3", {}, "Client"), listPart,
    h("h3", {}, "Filed as"),
    h("label", { class: "f", for: "desc" }, "Attachment description"), desc, count,
    h("label", { class: "f", for: "note" }, "Log note"), noteBox,
    h("p", { class: "muted small" }, "Logged in HawkSoft as an e-mail from the insured, category Emails. You see FreeBird's plan before anything is filed."),
    go);
}

function searchBox() {
  const q = h("input", { type: "text", placeholder: "Client name, phone or policy number", "aria-label": "Search HawkSoft" });
  const run = () => { if (q.value.trim()) match({ query: q.value.trim() }, runId); };
  q.addEventListener("keydown", (e) => { if (e.key === "Enter") run(); });
  return h("div", { class: "row" }, q, h("button", { class: "secondary", onclick: run }, "Search"));
}

async function review() {
  const my = runId;
  try {
    if (!S.draftId) {
      busy("Reading the e-mail…");
      const eml = await officeAsync(cb => Office.context.mailbox.item.getAsFileAsync(cb));
      const sg = await api("/stage", { emlBase64: eml, subject: S.subject });
      if (!sg.ok) return viewError(sg.error, review);
      S.draftId = sg.draftId;
    }
    if (my !== runId) return;
    busy("Asking FreeBird for its plan…", "Nothing is filed yet.");
    const pl = await api("/plan", { draftId: S.draftId, clientId: S.chosen.id, clientName: S.chosen.name, description: S.desc, note: S.note, extra: S.extra });
    if (my !== runId) return;
    if (!pl.ok) {
      if (pl.code === "restage") { S.draftId = null; return review(); }
      return viewError(pl.error + (pl.missing ? " (" + JSON.stringify(pl.missing) + ")" : ""), review, viewChoose);
    }
    if (pl.jobId) return filing(pl.jobId);
    viewConfirm(pl);
  } catch (e) { if (my === runId) handle(e, review); }
}

function viewConfirm(pl) {
  const lines = Array.isArray(pl.plan) ? pl.plan : (pl.plan ? [pl.plan] : []);
  const yes = h("button", { class: "primary", onclick: async () => {
    yes.disabled = true;
    try {
      const ap = await api("/approve", { pendingId: pl.pendingId });
      if (!ap.ok) return ap.code === "replan" ? review() : viewError(ap.error, review, viewChoose);
      filing(ap.jobId, ap.duplicate ? ap.note : null);
    } catch (e) { handle(e, review); }
  } }, "Yes, file it");
  show(mailCard(),
    h("h3", {}, "FreeBird will"),
    h("ol", { class: "plan" }, lines.map(l => h("li", {}, typeof l === "string" ? l : JSON.stringify(l)))),
    pl.earlierRun ? note("warn", h("p", {}, pl.earlierRun)) : null,
    h("div", { class: "actions" }, h("button", { class: "secondary", onclick: viewChoose }, "Back"), yes));
}

async function filing(jobId, dupNote) {
  const my = runId, t0 = Date.now();
  busy("Filing in HawkSoft…", "FreeBird attaches the e-mail and reads it back from HawkSoft. You can keep working.");
  while (Date.now() - t0 < POLL_MAX_MS) {
    await new Promise(r => setTimeout(r, POLL_MS));
    if (my !== runId) return;
    let j;
    try { j = await api("/job/" + encodeURIComponent(jobId)); } catch (e) { if (e instanceof BridgeDown) continue; return handle(e, () => filing(jobId)); }
    if (!j.ok) return viewError(j.error, () => filing(jobId));
    if (j.status === "queued" || j.status === "running") continue;
    if (j.status === "done") return viewDone(j, dupNote);
    if (j.status === "needs_input" && j.needsInput) return viewQuestion(j.needsInput);
    return viewFailed(j);
  }
  viewError("FreeBird is still working after 5 minutes. Check FreeBird's window on this PC.", () => filing(jobId));
}

function viewDone(j, dupNote) {
  const created = (j.created || []).filter(c => c && c.kind);
  show(h("div", { class: "center" }, h("div", { class: "done-icon", "aria-hidden": "true" }, "✓"),
      h("h2", {}, "Filed on " + S.chosen.name),
      h("p", { class: "muted" }, j.verified ? "FreeBird checked it in HawkSoft." : "FreeBird filed it but could not read it back. Check the client's attachments.")),
    dupNote ? note("warn", h("p", {}, dupNote)) : null,
    created.length ? [h("h3", {}, "In HawkSoft now"), h("ul", { class: "list" }, created.map(c => h("li", {}, (c.kind === "log" ? "Log note: " : c.kind === "attachment" ? "Attachment: " : c.kind + ": ") + c.text)))] : null,
    (j.decisions || []).length ? [h("h3", {}, "FreeBird decided"), h("ul", { class: "list" }, j.decisions.map(d => h("li", {}, d)))] : null,
    (j.needsYou || []).length ? note("warn", h("p", {}, "Still for you:"), h("ul", { class: "list" }, j.needsYou.map(d => h("li", {}, d)))) : null);
}

function viewQuestion(ni) {
  const opts = (ni.options || []).map(o => typeof o === "string" ? { value: o, label: o } : o);
  const answer = (v) => { S.extra[ni.param] = v; review(); };
  const free = h("input", { type: "text", placeholder: "Your answer" });
  show(mailCard(),
    note("warn", h("p", {}, "FreeBird needs one more detail. Nothing was filed yet.")),
    h("h2", {}, ni.question),
    ni.hint ? h("p", { class: "muted small" }, ni.hint) : null,
    opts.length ? h("div", { class: "opts" }, opts.map(o => h("button", { class: "secondary", onclick: () => answer(o.value) }, o.label || o.value)))
      : [free, h("button", { class: "primary", onclick: () => { if (free.value.trim()) answer(free.value.trim()); } }, "Continue")],
    h("button", { class: "link", onclick: viewChoose }, "Back"));
}

function viewFailed(j) {
  const may = j.maySaved;
  show(mailCard(),
    note("bad", h("p", {}, "FreeBird could not file this e-mail."), h("p", {}, j.error || ("It ended with status " + j.status + "."))),
    may ? note("warn", h("p", {}, "HawkSoft may already have saved " + may + ". Check the client in HawkSoft before trying again.")) : null,
    h("div", { class: "actions" }, h("button", { class: "secondary", onclick: viewChoose }, "Back"), may ? null : h("button", { class: "primary", onclick: review }, "Try again")));
}

function viewError(msg, retry, back) {
  show(note("bad", h("p", {}, msg || "Something went wrong.")),
    h("div", { class: "actions" }, back ? h("button", { class: "secondary", onclick: back }, "Back") : null, retry ? h("button", { class: "primary", onclick: retry }, "Try again") : null));
}

function viewBridgeDown() {
  show(h("h2", {}, "FreeBird for Outlook isn't running"),
    h("p", {}, "This panel talks to a small FreeBird helper on this PC, and it isn't answering."),
    h("p", { class: "muted small" }, "Start “FreeBird for Outlook” on this PC, then press Try again."),
    h("button", { class: "primary", onclick: start }, "Try again"));
}

function viewFreeBirdDown(msg) {
  show(h("h2", {}, "FreeBird isn't running"),
    h("p", {}, msg || "Open FreeBird from the Start menu, then try again."),
    h("button", { class: "primary", onclick: start }, "Try again"));
}

function viewPair() {
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
