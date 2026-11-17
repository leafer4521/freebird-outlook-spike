// FreeBird for Outlook - the panel. Talks only to the bridge on this PC (bridge\Bridge.cs), never to any server:
// the e-mail goes from Outlook straight to the PC.
"use strict";

const PANEL_VERSION = "0.3.8";
let selDiag = "";                   // what Outlook answered to "which e-mails are selected?" (counts and errors only): shown by Settings, logged by the helper      // shown next to Settings: which panel Outlook is really running (GitHub Pages caches 10 minutes)
const BRIDGE = window.FBO_BRIDGE || "http://127.0.0.1:8791";      // FBO_BRIDGE: addin\test\harness.html (a test helper on another port)
const KEY_NAME = "fbo.bridgeKey";
const SET_NAME = "fbo.settings";
const TEXT_LIMIT = 20000;           // characters of the body sent for "is this client named in the e-mail?"
const POLL_MS = 1000, POLL_MAX_MS = 10 * 60 * 1000;
const FILE_MAX = 25 * 1024 * 1024;  // FreeBird attaches files up to 25 MB (attach-file)
const MAX_MAILS = 50;               // several e-mails selected: FreeBird files up to 50 files in one batch (one approval)

// Settings (Ben, 2026-11-16): fewer clicks by default - "Attach" files at once (the click is the approval; FreeBird's plan is still
// shown when FreeBird warns, e.g. the same filing was already done today).
// The policy is not a setting (Leif, 2026-11-16 live test): a "File on" menu shows as soon as a client is chosen, set to the client;
// a policy picked there gets the e-mail instead.
const DEFAULTS = { showPlan: false };

const S = {};                       // the current e-mail and what has been chosen for it
let runId = 0;                      // a newer e-mail (ItemChanged / SelectedItemsChanged) cancels the views of an older one
// While e-mails are being filed (pinned panel), a click on another e-mail does not pull the panel away: it is offered when the filing ends.
let filingNow = false, changedWhileFiling = false, changeTimer = 0;

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
function paint(nodes) { const app = document.getElementById("app"); app.replaceChildren(...nodes.flat(2).filter(n => n !== null && n !== undefined && n !== false)); window.scrollTo(0, 0); }
function show(...nodes) { filingNow = false; paint(nodes); }      // any view but "busy": the panel may follow another e-mail again
// busy: a spinner. #busyWait says what FreeBird does first when it is busy with another task (watchBusy, filing)
function busy(text, sub, wait) {
  S.view = "busy";
  paint([h("div", { class: "center" }, h("div", { class: "spin", "aria-hidden": "true" }), h("p", {}, text), sub ? h("p", { class: "muted small" }, sub) : null),
    h("div", { class: "note warn", id: "busyWait", hidden: !wait }, wait || "")]);
}
const sleep = (ms) => new Promise(r => setTimeout(r, ms));
// FreeBird runs one task at a time, in the order asked (a Claude task, another filing, this one). What it is doing first, in plain words:
function busyNote(bz) {
  if (!bz) return null;
  if (bz.paused) return "FreeBird is paused. It carries on once it is resumed in FreeBird's window.";
  if (bz.handedOver) return "You have HawkSoft right now. FreeBird waits until you hand it back.";
  if (bz.what) return (bz.ours ? "FreeBird is finishing another filing first" : "FreeBird is busy with another task first") + ": " + bz.what + (bz.client ? " for " + bz.client : "") +
    ". It does one task at a time, in order, so this waits its turn and then runs by itself.";
  return null;
}
// a look-up waits behind FreeBird's current task: every few seconds the panel asks the helper what that is and says so
async function watchBusy(p) {
  let stop = false;
  (async () => {
    while (!stop) {
      await sleep(2500); if (stop) break;
      let t = null; try { t = busyNote((await api("/status", {})).busy); } catch (e) { }
      const el = document.getElementById("busyWait"); if (el && !stop) { el.textContent = t || ""; el.hidden = !t; }
    }
  })();
  try { return await p; } finally { stop = true; }
}
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
// when a message arrived: tells the messages of one thread apart (they share a subject; live 2026-11-16 20:09)
function whenText(d) {
  try { d = d instanceof Date ? d : d ? new Date(d) : null; if (!d || isNaN(d)) return ""; } catch (e) { return ""; }
  return d.toLocaleString("en-US", { weekday: "short", month: "numeric", day: "numeric", hour: "numeric", minute: "2-digit" });
}
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
  try { Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, onChange); } catch (e) { }
  try { if (Office.EventType.SelectedItemsChanged) Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, onChange); } catch (e) { }
  start();
});

function onChange() {
  if (filingNow) { changedWhileFiling = true; return; }
  clearTimeout(changeTimer); changeTimer = setTimeout(start, 200);      // ItemChanged and SelectedItemsChanged can both fire for one click
}
function nextMail() { return changedWhileFiling ? h("button", { class: "secondary next", onclick: start }, "Go to the e-mail selected now") : null; }

// The e-mail's attachments: files (photos, PDFs ...) and attached e-mails can be handed over (Mailbox 1.8); cloud links cannot.
// Pictures inside the message body are ticked too - Gmail puts photos pasted into the message there (live 2026-11-16: both claim photos
// were in the body) - except ones that look like signature art: Outlook's image001.png, logos, social icons, or under 5 KB.
const SIG_NAME = /^(image\d{3}\.|outlook-|~wrl|logo|signature|sig[ _.-]|facebook|twitter|linkedin|instagram|youtube|x-logo)/i;
const SIG_MAX = 5 * 1024;
// A reply or forward quotes the earlier message, pictures in its body included (live 2026-11-16 20:13: a Gmail "Re:" carried the
// original's photo as its own in-body attachment). Those are not ticked: shown as "likely quoted", the person can tick one that is new.
const REPLY_SUBJ = /^\s*(re|fw|fwd|aw|sv|wg)\s*:/i;
function readAttachments(item, subject) {
  const isReply = REPLY_SUBJ.test(subject || "");
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
    const inline = !!a.isInline;
    const sig = inline && (SIG_NAME.test(name) || (a.size || 0) < SIG_MAX);
    return { id: a.id, name, size: a.size || 0, inline, sig, quoted: inline && isReply && !sig, type, why };
  });
}
function tickedByDefault(a) { return !a.why && !a.sig && !a.quoted && !a.seenIn; }
function attMeta(a) {
  return [a.size ? kb(a.size) : "", a.why || (a.sig ? "in the message body, looks like a signature picture"
    : a.seenIn ? "the same file is in the e-mail of " + a.seenIn
    : a.quoted ? "in the body of a reply, likely quoted from an earlier e-mail"
    : a.inline ? "in the message body" : "")].filter(Boolean).join(" · ");
}

async function start() {
  const my = ++runId; changedWhileFiling = false;
  for (const k of Object.keys(S)) if (k !== "key") delete S[k];
  // several e-mails selected (Mailbox 1.15: the new Outlook, Outlook on the web): filed together, each on its own client
  const sel = await selectedItems();
  if (my !== runId) return;
  if (sel.length > 1) return startMany(sel, my);
  const item = Office.context.mailbox.item;
  if (!item) { show(h("p", { class: "muted" }, "Open an e-mail to file it in HawkSoft.")); return; }
  S.from = item.from && item.from.emailAddress || "";
  S.fromName = item.from && item.from.displayName || "";
  S.subject = item.subject || "";
  S.when = whenText(item.dateTimeCreated);
  S.extra = {};
  S.pols = {};
  S.atts = readAttachments(item, item.subject);
  S.pick = new Set(S.atts.filter(tickedByDefault).map(a => a.id));
  busy("Checking FreeBird…");
  try {
    S.text = (await officeAsync(cb => item.body.getAsync(Office.CoercionType.Text, cb)) || "").slice(0, TEXT_LIMIT);
  } catch (e) { S.text = ""; }
  if (my !== runId) return;
  await checkAndMatch(my);
}

async function checkAndMatch(my) {
  let st;
  try { st = await api("/status", { diag: selDiag }); }
  catch (e) { if (my !== runId) return; return e instanceof NeedsPair ? viewPair() : e instanceof BridgeDown ? viewBridgeDown() : viewError(e.message, () => start()); }
  if (my !== runId) return;
  document.getElementById("testBadge").hidden = !st.testOnly;
  S.testOnly = st.testOnly; S.testPrefix = st.testPrefix;
  if (!st.freebird || !st.freebird.up) return viewFreeBirdDown(st.freebird && st.freebird.message);
  if (st.freebird.state && st.freebird.state !== "ok") S.health = st.freebird.message;
  if (S.many) return matchMany(my);
  await match({ from: S.from }, my);
}

async function match(q, my) {
  busy(q.query ? "Searching HawkSoft for “" + q.query + "”…" : "Looking up " + (S.from || "the sender") + " in HawkSoft…", "FreeBird checks HawkSoft's client list. Nothing is changed.");
  let m;
  try { m = await watchBusy(api("/match", Object.assign({ fromName: S.fromName, subject: S.subject, text: S.text }, q))); }
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
  if (S.many) {
    const on = S.many.filter(m => m.on), senders = new Set(on.map(m => (m.from || "").toLowerCase())).size;
    return h("div", { class: "mail" }, h("div", { class: "from" }, plural(on.length, "e-mail", "e-mails")), h("div", { class: "subj" }, "from " + plural(senders, "sender", "senders")));
  }
  return h("div", { class: "mail" },
    h("div", { class: "from" }, S.fromName && S.fromName !== S.from ? S.fromName + " · " + S.from : (S.from || "Unknown sender")),
    h("div", { class: "subj" }, S.subject || "(no subject)"),
    S.when ? h("div", { class: "when" }, "Received " + S.when) : null);
}

function picked() { return S.atts.filter(a => S.pick.has(a.id)); }
function polLabel(p) { return [p.tab || p.policy, p.number, p.effective && p.expiration ? p.effective + "–" + p.expiration : p.effective ? "from " + p.effective : ""].filter(Boolean).join(" · "); }
function whereText() {
  const l = S.pols[S.chosen.id], p = S.policy && Array.isArray(l) && l.find(x => x.policy === S.policy);
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

  // File on: the client (default) or one of its policies. Read once per client; until then (about a second) it is the client, and
  // Attach does not wait for it. Policies in force first, prospects and the rest below.
  let polPart = null;
  if (S.chosen) {
    const pols = S.pols[S.chosen.id];
    if (pols === undefined) loadPolicies(S.chosen, redraw);
    const list = Array.isArray(pols) ? pols : [];
    const sel = policySelect(pols, S.policy, "File on", (v) => { S.policy = v; S.extra = {}; redraw(); });
    sel.id = "policy";
    const sub = pols === undefined ? "Reading the client's policies…"
      : pols.error ? "Couldn't read the policies (" + pols.error + "), so it goes on the client."
      : !list.length ? "This client has no policies in HawkSoft, so it goes on the client."
      : S.policy ? "On this policy, not on the client's own tab." : "Leave it on the client, or pick a policy.";
    polPart = [h("h3", {}, "File on"), sel, h("p", { class: "muted small sub" }, sub)];
  }

  // the attachments: ticked = filed too, each as its own attachment with its own log
  let attPart = null;
  if (S.atts.length) {
    attPart = [h("h3", {}, "Attachments"),
      h("div", { class: "atts" }, S.atts.map(a => {
        const box = h("input", { type: "checkbox", checked: S.pick.has(a.id), disabled: !!a.why });
        box.addEventListener("change", () => { if (box.checked) S.pick.add(a.id); else S.pick.delete(a.id); redraw(); });
        return h("label", { class: "att" + (a.why ? " off" : "") }, box,
          h("div", {}, h("div", { class: "name" }, a.name), h("div", { class: "meta" }, attMeta(a))));
      }))];
  }

  const n = picked().length;
  const what = "the e-mail" + (n ? " and " + plural(n, "attachment", "attachments") : "");
  const go = h("button", { class: "primary", disabled: !S.chosen, onclick: () => { S.desc = desc.value.trim(); S.note = noteBox.value.trim(); review(); } },
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
  if (!S.showSettings) return h("button", { class: "link quiet", onclick: () => { S.showSettings = true; redraw(); } }, "Settings · panel " + PANEL_VERSION);
  const row = (k, label, sub) => {
    const box = h("input", { type: "checkbox", checked: !!set[k] });
    box.addEventListener("change", () => { setSetting(k, box.checked); redraw(); });
    return h("label", { class: "att" }, box, h("div", {}, h("div", { class: "name" }, label), h("div", { class: "meta" }, sub)));
  };
  return h("div", { class: "settings" },
    h("h3", {}, "Settings"),
    row("showPlan", "Show FreeBird's plan first", "Off: Attach files at once. On: you confirm FreeBird's plan."),
    h("button", { class: "link quiet", onclick: () => { S.showSettings = false; redraw(); } }, "Close settings"));
}

async function loadPolicies(c, redraw) {
  S.polBusy = S.polBusy || new Set();
  if (S.pols[c.id] !== undefined || S.polBusy.has(c.id)) return;
  S.polBusy.add(c.id); const my = runId;
  let r;
  try { r = await api("/policies", { clientId: c.id, clientName: c.name }); }
  catch (e) { if (e instanceof NeedsPair || e instanceof BridgeDown) { S.polBusy.delete(c.id); return handle(e); } r = { ok: false, error: e.message }; }
  if (my !== runId) return;
  S.polBusy.delete(c.id);
  S.pols[c.id] = r.ok ? (r.policies || []) : { error: r.error || "no answer" };
  if (S.many ? S.view === "many" : (S.view === "choose" && S.chosen && S.chosen.id === c.id)) redraw();
}
// "File on": the client (default) or one of its policies - policies in force first, prospects and the rest below
function policySelect(pols, value, label, onPick) {
  const list = Array.isArray(pols) ? pols : [];
  const opt = (p) => h("option", { value: p.policy }, polLabel(p));
  const inForce = list.filter(p => p.inForce), other = list.filter(p => !p.inForce);
  const sel = h("select", { "aria-label": label, disabled: pols === undefined },
    h("option", { value: "" }, "The client (no policy)"),
    inForce.length && other.length ? [h("optgroup", { label: "In force" }, inForce.map(opt)), h("optgroup", { label: "Prospects and others" }, other.map(opt))] : list.map(opt));
  sel.value = value || "";
  sel.addEventListener("change", () => onPick(sel.value));
  return sel;
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
  const attachments = await readContents(item, files);
  const sg = await api("/stage", { emlBase64: eml, subject: S.subject, attachments });
  if (!sg.ok) { viewError(sg.error, review); return false; }
  if (S.stagedKey !== undefined && S.stagedKey !== key) { S.extra = {}; S.skip = new Set(); }      // other files ticked: the answers no longer fit
  S.draftId = sg.draftId; S.stagedKey = key;
  S.taskNames = ["the e-mail"].concat(files.map(a => "“" + a.name + "”"));
  return true;
}

async function readContents(item, files) {
  const attachments = [];
  for (const a of files) {
    const c = await officeAsync(cb => item.getAttachmentContentAsync(a.id, cb));
    const fmt = String(c.format || "").toLowerCase();
    if (fmt === "base64") attachments.push({ name: a.name, base64: c.content });
    else if (fmt === "eml" || fmt === "icalendar") attachments.push({ name: a.name, base64: b64Text(c.content) });
  }
  return attachments;
}

function viewMain() { return S.many ? viewMany() : viewChoose(); }

async function review() {
  if (S.many) return reviewMany();
  const my = runId, quick = !settings().showPlan;
  filingNow = true;
  try {
    if (!(await stage())) return;
    if (my !== runId) return;
    busy(quick ? "Filing in HawkSoft…" : "Asking FreeBird for its plan…", quick ? "Getting FreeBird ready." : "Nothing is filed yet.");
    const pl = await api("/plan", { draftId: S.draftId, clientId: S.chosen.id, clientName: S.chosen.name, policy: S.policy || "", from: S.from, fromName: S.fromName, subject: S.subject, description: S.desc, note: S.note, extra: S.extra, skip: [...(S.skip || [])] });
    if (my !== runId) return;
    if (!pl.ok) {
      if (pl.code === "restage") { S.draftId = null; return review(); }
      if (pl.code === "nothing") return viewLeft();
      return viewError(pl.error + (pl.missing ? " (" + JSON.stringify(pl.missing) + ")" : ""), review, viewChoose);
    }
    afterPlan(pl, quick);
  } catch (e) { if (my === runId) handle(e, review); }
}
// FreeBird's answer to /plan, one e-mail or several
function afterPlan(pl, quick) {
  // the files in this filing (not left out), with their number in the whole list (answers and skips go by that number)
  const all = S.taskNames || [];
  S.liveIx = all.map((_, i) => i).filter(i => !(S.skip && S.skip.has(i)));
  S.liveNames = S.liveIx.map(i => all[i]);
  S.liveMail = S.taskMail ? S.liveIx.map(i => S.taskMail[i]) : null;
  if (pl.jobIds) return filing(pl.jobIds);
  // quick: the person's "Attach" click is the approval - unless FreeBird warns (the same filing was already done today): then its plan is shown
  if (quick && !pl.earlierRun) return approve(pl);
  viewConfirm(pl);
}

async function approve(pl) {
  filingNow = true;
  try {
    const ap = await api("/approve", { pendingId: pl.pendingId });
    if (!ap.ok) return ap.code === "replan" ? review() : viewError(ap.error, review, viewMain);
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
    h("div", { class: "actions" }, h("button", { class: "secondary", onclick: viewMain }, "Back"), yes));
}

// One job per file (the e-mail first). Each is followed until it ends; FreeBird does them one after the other.
async function filing(jobIds, dupNote) {
  const my = runId, t0 = Date.now(), res = jobIds.map(() => null);
  filingNow = true;
  const names = jobIds.map((_, i) => (S.liveNames && S.liveNames[i]) || ("file " + (i + 1)));
  const sub = "FreeBird is working in HawkSoft on this PC. Leave HawkSoft alone until it's done.";
  busy("Filing in HawkSoft…", sub);
  // the clock starts once FreeBird is on these files: behind a long task of someone else's (Claude, another filing) it only waits
  let waitFor = null, t1 = null;
  while (t1 === null ? Date.now() - t0 < POLL_MAX_MS * 3 : Date.now() - t1 < POLL_MAX_MS) {
    await sleep(POLL_MS);
    if (my !== runId) return;
    waitFor = null;
    for (let i = 0; i < jobIds.length; i++) {
      if (res[i]) continue;
      let j;
      try { j = await api("/job/" + encodeURIComponent(jobIds[i])); } catch (e) { if (e instanceof BridgeDown) break; return handle(e, () => filing(jobIds, dupNote)); }
      if (my !== runId) return;
      if (!j.ok) return viewError(j.error, () => filing(jobIds, dupNote));
      if (j.status === "queued" && i === res.findIndex(r => !r)) waitFor = busyNote(j.waitingFor);
      if (j.status !== "queued" && t1 === null) t1 = Date.now();
      if (j.status === "queued" || j.status === "running") break;      // the next ones wait behind this one
      res[i] = j;
    }
    const left = res.findIndex(r => !r);
    if (left < 0) return viewResult(res, names, dupNote);
    if (waitFor) busy("Waiting for FreeBird…", "Nothing is filed yet. FreeBird starts on it by itself as soon as it is free; you can leave this panel open.", waitFor);
    else if (jobIds.length > 1) busy("Filing " + (left + 1) + " of " + jobIds.length + " in HawkSoft…", names[left].replace(/^the/, "The") + ". " + sub);
    else busy("Filing in HawkSoft…", sub);
  }
  viewError(t1 === null ? "FreeBird has not started on this after 30 minutes. Check FreeBird's window on this PC." : "FreeBird is still working after 10 minutes. Check FreeBird's window on this PC.", () => filing(jobIds, dupNote));
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
  if (S.many) return viewDoneMany(res, dupNote, verified);
  show(h("div", { class: "center" }, h("div", { class: "done-icon", "aria-hidden": "true" }, "✓"),
      h("h2", {}, "Filed on " + whereText()),
      h("p", { class: "muted" }, verified ? "FreeBird checked it in HawkSoft." : "FreeBird filed it but could not read all of it back. Check the client's attachments.")),
    dupNote ? note("warn", h("p", {}, dupNote)) : null,
    S.skip && S.skip.size && S.taskNames ? h("p", { class: "muted small" }, "Left as they were (already there): " + [...S.skip].map(i => S.taskNames[i]).join(", ") + ".") : null,
    created.length ? [h("h3", {}, "In HawkSoft now"), h("ul", { class: "list" }, created.map(c => h("li", {}, (c.kind === "log" ? "Log note: " : c.kind === "attachment" ? "Attachment: " : c.kind + ": ") + c.text)))] : null,
    decisions.length ? [h("h3", {}, "FreeBird decided"), h("ul", { class: "list" }, decisions.map(d => h("li", {}, d)))] : null,
    needsYou.length ? note("warn", h("p", {}, "Still for you:"), h("ul", { class: "list" }, needsYou.map(d => h("li", {}, d)))) : null,
    nextMail());
}

// FreeBird asked about one of the files. The answer goes with that file only; filed ones are not filed again (FreeBird reuses them).
function viewQuestion(ni, ix, res, names) {
  S.view = "question";
  const opts = (ni.options || []).map(o => typeof o === "string" ? { value: o, label: o } : o);
  // the same question for several files (live 2026-11-16: 5 files of 2 e-mails were all "already there today"): asked once, the answer goes to each
  const same = res ? res.map((r, i) => i).filter(i => res[i].status === "needs_input" && res[i].needsInput && res[i].needsInput.param === ni.param) : [ix];
  const orig = (i) => (S.liveIx && S.liveIx[i] !== undefined ? S.liveIx[i] : i);
  const answer = (v) => { for (const i of same) (S.extra[orig(i)] = S.extra[orig(i)] || {})[ni.param] = v; review(); };
  const free = h("input", { type: "text", placeholder: "Your answer" });
  const many = same.length > 1, dup = ni.param === "AllowDuplicate";
  const about = res && res.length > 1 && !many ? "About " + names[ix] + ": " : "";
  let list = many ? h("ul", { class: "list" }, same.map(i => h("li", {}, names[i].replace(/^the/, "The")))) : null;
  // several files "already there today": some may really be repeats, some only share words with an earlier one (live 2026-11-16: the subject
  // "LF Landscaping" vs the earlier "Please send proof of insurance for LF Landscaping"). One tick per file, none ticked: only the ticked
  // ones are attached again, the others are left out.
  let again = null;
  if (dup && many) {
    const ticks = new Set();
    const go = h("button", { class: "primary" }, "Leave them all as they are");
    const label = () => { go.textContent = ticks.size ? "Attach the ticked " + (ticks.size === 1 ? "one" : ticks.size) + " again, leave the rest" : "Leave them all as they are"; };
    list = h("div", { class: "atts" }, same.map(i => {
      const box = h("input", { type: "checkbox" });
      box.addEventListener("change", () => { if (box.checked) ticks.add(i); else ticks.delete(i); label(); });
      return h("label", { class: "att" }, box, h("div", {}, h("div", { class: "name" }, names[i].replace(/^the/, "The")), h("div", { class: "meta" }, "Tick to attach a second copy")));
    }));
    go.addEventListener("click", () => {
      S.skip = S.skip || new Set();
      for (const i of same) { if (ticks.has(i)) (S.extra[orig(i)] = S.extra[orig(i)] || {})[ni.param] = true; else S.skip.add(orig(i)); }
      review();
    });
    again = go;
  }
  show(mailCard(),
    note("warn", h("p", {}, (dup ? "FreeBird attached nothing it already had." : "FreeBird needs one more detail.") + (res && res.some(r => r.status === "done") ? "" : " Nothing was filed yet."))),
    res ? filedList(res, names) : null,
    dup ? [h("h2", {}, many ? "These are already on the client from earlier today:" : about + "This is already on the client from earlier today."), list,
           h("p", { class: "muted small" }, "FreeBird compares names with what was attached today, not the contents, so a different file with a similar name can show up here. " + (many ? "Tick only the ones that really need a second copy." : "Attach it again only if it is really a different one.")),
           !many ? h("p", { class: "muted small" }, "FreeBird says: " + ni.question) : null]
        : [h("h2", {}, about + ni.question), many ? [h("p", { class: "small" }, "The same question for each of these; your answer goes to all of them:"), list] : null,
           ni.hint ? h("p", { class: "muted small" }, ni.hint) : null],
    again ? again : opts.length ? h("div", { class: "opts" }, opts.map(o => h("button", { class: "secondary", onclick: () => answer(o.value) }, dup ? "Attach it again anyway" : (o.label || o.value))))
      : [free, h("button", { class: "primary", onclick: () => { if (free.value.trim()) answer(free.value.trim()); } }, "Continue")],
    h("button", { class: "link", onclick: viewMain }, dup && !many ? "Back (leave it as it is)" : "Back"),
    nextMail());
}

// every file asked about was left out: nothing more to file
function viewLeft() {
  S.view = "done";
  show(h("div", { class: "center" }, h("div", { class: "done-icon", "aria-hidden": "true" }, "✓"),
      h("h2", {}, "Nothing more to file"),
      h("p", { class: "muted" }, "The files FreeBird found already there were left as they are. Anything else was filed before.")),
    nextMail());
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
    h("div", { class: "actions" }, h("button", { class: "secondary", onclick: viewMain }, "Back"), may.length ? null : h("button", { class: "primary", onclick: review }, "Try again")),
    nextMail());
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

// ------------------------------------------------------------------ several e-mails selected together (Mailbox 1.15)
// Ben's two cases: several e-mails from the SAME sender (usually one client, matched once) and from DIFFERENT senders (each its own client).
// Each e-mail is matched like a single one (sender, then "named in the e-mail"); the person checks every row, then ONE Attach click files
// all of them: one FreeBird batch, one plan, one approval. Each e-mail goes on its client (or a policy picked under it) with its subject as the
// description; a File on menu (policy) and attachment ticks per e-mail, plus one switch for all attachments.

function selectedItems() {
  return new Promise((resolve) => {
    try {
      let d = ""; try { d = Office.context.diagnostics ? Office.context.diagnostics.platform + " " + Office.context.diagnostics.version : ""; } catch (e) { }
      const sets = [13, 14, 15].filter(v => { try { return Office.context.requirements.isSetSupported("Mailbox", "1." + v); } catch (e) { return false; } }).map(v => "1." + v).join(",");
      selDiag = d + " sets " + (sets || "<1.13");
      if (typeof Office.context.mailbox.getSelectedItemsAsync !== "function") { selDiag += " no getSelectedItemsAsync"; return resolve([]); }
      if (!Office.context.requirements.isSetSupported("Mailbox", "1.15")) { selDiag += " (1.15 missing)"; }
      Office.context.mailbox.getSelectedItemsAsync((r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) { selDiag += " selected: failed " + (r.error && (r.error.code + " " + r.error.message) || ""); return resolve([]); }
        const v = Array.isArray(r.value) ? r.value : [];
        selDiag += " selected: " + v.length + " [" + v.map(x => String(x.itemType || "?")).join(",") + "]";
        resolve(Office.context.requirements.isSetSupported("Mailbox", "1.15") ? v.filter(x => !x.itemType || String(x.itemType).toLowerCase() === "message") : []);
      });
    } catch (e) { selDiag += " error " + e.message; resolve([]); }
  });
}
// One selected e-mail at a time (Outlook allows only one loaded at once): load, read, unload.
async function withItem(id, fn) {
  const it = await officeAsync(cb => Office.context.mailbox.loadItemByIdAsync(id, cb));
  try { return await fn(it); }
  finally { await new Promise(r => { try { it.unloadAsync(() => r()); } catch (e) { r(); } }); }
}
function mailFiles(m) { return S.attsOn && m.pick ? (m.atts || []).filter(a => m.pick.has(a.id)) : []; }
function mailDesc(m) { return (m.subject || ("E-mail from " + (m.fromName || m.from))).slice(0, 100); }
function mailNote(m) { return "E-mail from " + (m.fromName || m.from || "the sender") + (m.subject ? ": " + m.subject : "") + "."; }
function senderGroups() {
  const g = new Map();
  for (const m of S.many) { const k = (m.from || "?").toLowerCase(); if (!g.has(k)) g.set(k, []); g.get(k).push(m); }
  return [...g.values()];
}

async function startMany(sel, my) {
  S.many = sel.slice(0, MAX_MAILS).map((x, n) => ({ n, itemId: x.itemId, subject: x.subject || "", on: true }));
  S.tooMany = sel.length > MAX_MAILS ? sel.length : 0;
  S.attsOn = true; S.extra = {}; S.pols = {};
  for (const m of S.many) {
    busy("Reading " + plural(S.many.length, "e-mail", "e-mails") + "…", (m.n + 1) + " of " + S.many.length);
    try {
      await withItem(m.itemId, async (it) => {
        m.from = it.from && it.from.emailAddress || ""; m.fromName = it.from && it.from.displayName || ""; m.subject = it.subject || m.subject;
        m.atts = readAttachments(it, it.subject || m.subject); m.at = it.dateTimeCreated ? new Date(it.dateTimeCreated).getTime() : 0; m.when = whenText(it.dateTimeCreated);
        try { m.text = (await officeAsync(cb => it.body.getAsync(Office.CoercionType.Text, cb)) || "").slice(0, TEXT_LIMIT); } catch (e) { m.text = ""; }
      });
    } catch (e) { m.error = "Outlook could not open this e-mail: " + e.message; m.on = false; m.atts = []; }
    if (my !== runId) return;
  }
  S.many.sort((a, b) => (a.at || 0) - (b.at || 0));      // oldest first: a thread reads in order
  // a file with the same name and size as one in an earlier selected e-mail (a thread quoting itself, a forward) is not ticked again
  const seen = new Map();
  for (const m of S.many) {
    for (const a of m.atts || []) { const k = a.name.toLowerCase() + "|" + a.size; if (a.why) continue; if (seen.has(k)) a.seenIn = seen.get(k); else seen.set(k, m.when || "an earlier one"); }
    m.pick = new Set((m.atts || []).filter(tickedByDefault).map(a => a.id));
  }
  await checkAndMatch(my);
}

// one look-up per e-mail (the helper keeps a sender's matches for 2 minutes: the same sender is asked of HawkSoft once)
async function matchMany(my) {
  const list = S.many.filter(m => !m.error);
  for (let i = 0; i < list.length; i++) {
    const m = list[i];
    busy("Looking up the senders in HawkSoft…", (i + 1) + " of " + list.length + ". Nothing is changed.");
    let r;
    try { r = await watchBusy(api("/match", { from: m.from, fromName: m.fromName, subject: m.subject, text: m.text })); }
    catch (e) { if (my !== runId) return; return handle(e, () => matchMany(my)); }
    if (my !== runId) return;
    m.match = r.ok ? r : { matches: [], error: r.error };
    m.chosen = r.ok && r.suggested ? r.matches.find(x => x.id === r.suggested) || null : null;
  }
  viewMany();
}

function viewMany() {
  const keepY = S.view === "many" ? window.scrollY : 0;      // a pick halfway down the list does not jump back to the top
  S.view = "many";
  const set = settings(), groups = senderGroups();
  const on = S.many.filter(m => m.on), open = on.filter(m => !m.chosen);
  const files = on.reduce((n, m) => n + mailFiles(m).length, 0), allFiles = S.many.reduce((n, m) => n + (m.atts || []).filter(a => !a.why).length, 0);
  const tasks = on.length + files, clientsN = new Set(on.filter(m => m.chosen).map(m => m.chosen.id)).size;

  const parts = groups.map(g => {
    const head = g[0];
    const who = head.fromName && head.fromName !== head.from ? head.fromName + " · " + head.from : (head.from || "Unknown sender");
    // a client picked for one e-mail can be used for the others from the same sender that have none yet - only on the person's click
    const picked = g.find(m => m.on && m.chosen);
    const rest = picked ? g.filter(m => m.on && !m.chosen && m.match && (m.match.matches || []).some(c => c.id === picked.chosen.id && !c.blocked)) : [];
    return h("section", { class: "sender" },
      h("div", { class: "who" }, who, g.length > 1 ? h("span", { class: "muted" }, " · " + plural(g.length, "e-mail", "e-mails")) : null),
      g.map(mailRow),
      rest.length ? h("button", { class: "link", onclick: () => { rest.forEach(m => { m.chosen = (m.match.matches || []).find(c => c.id === picked.chosen.id); }); viewMany(); } },
        "Use " + picked.chosen.name + (rest.length === 1 ? " for “" + (rest[0].subject || "(no subject)") + "” too" : " for the other " + rest.length + " e-mails from this sender")) : null);
  });

  const tog = h("input", { type: "checkbox", checked: S.attsOn, disabled: !allFiles });
  tog.addEventListener("change", () => { S.attsOn = tog.checked; viewMany(); });
  const go = h("button", { class: "primary", disabled: !on.length || open.length > 0 || tasks > MAX_MAILS, onclick: () => review() },
    set.showPlan ? "Review and file" : "Attach " + plural(on.length, "e-mail", "e-mails"));
  const hint = !on.length ? "Tick at least one e-mail."
    : open.length ? "Choose the client for " + plural(open.length, "e-mail", "e-mails") + " first, or untick " + (open.length === 1 ? "it" : "them") + "."
    : tasks > MAX_MAILS ? "That is " + tasks + " files with the attachments. FreeBird files up to " + MAX_MAILS + " in one go: untick some e-mails or the attachments."
    : set.showPlan ? "You see FreeBird's plan for all of them before anything is filed."
    : "Pressing Attach files " + plural(on.length, "e-mail", "e-mails") + (files ? " and " + plural(files, "attachment", "attachments") : "") + " on " + plural(clientsN, "client", "clients") + " right away, one after the other. FreeBird reads each back from HawkSoft.";

  show(h("div", { class: "mail" }, h("div", { class: "from" }, plural(S.many.length, "e-mail", "e-mails") + " selected"), h("div", { class: "subj" }, "from " + plural(groups.length, "sender", "senders"))),
    S.tooMany ? note("warn", h("p", {}, "Only the first " + MAX_MAILS + " of the " + S.tooMany + " selected e-mails are here.")) : null,
    S.health ? note("warn", h("p", {}, S.health)) : null,
    parts,
    h("h3", {}, "Attachments"),
    h("label", { class: "att" + (allFiles ? "" : " off") }, tog, h("div", {},
      h("div", { class: "name" }, "Also file the attachments as separate files"),
      h("div", { class: "meta" }, allFiles ? "Tick them under each e-mail. They are inside the filed e-mails either way." : "None of these e-mails has attachments to file."))),
    h("p", { class: "muted small" }, "Each e-mail goes on its client, or on the policy picked under it, described by its subject, with the log note “E-mail from <sender>: <subject>.”"),
    h("p", { class: "small hint" }, hint),
    go,
    settingsPart(viewMany));
  if (keepY) window.scrollTo(0, keepY);
}

function mailRow(m) {
  const box = h("input", { type: "checkbox", checked: m.on, disabled: !!m.error, "aria-label": "File this e-mail" });
  box.addEventListener("change", () => { m.on = box.checked; viewMany(); });
  const meta = m.error || (m.when ? "Received " + m.when : "");
  // its attachments, each with a tick (signature pictures, quoted pictures and files already in an earlier selected e-mail start unticked)
  const atts = !m.error && m.on && S.attsOn && (m.atts || []).length ? h("div", { class: "matts" }, m.atts.map(a => {
    const b = h("input", { type: "checkbox", checked: m.pick.has(a.id), disabled: !!a.why });
    b.addEventListener("change", () => { if (b.checked) m.pick.add(a.id); else m.pick.delete(a.id); viewMany(); });
    return h("label", { class: "att" + (a.why ? " off" : "") }, b, h("div", {}, h("div", { class: "name" }, a.name), h("div", { class: "meta" }, attMeta(a))));
  })) : null;
  let pick = null;
  if (!m.error && m.on) {
    const ms = (m.match && m.match.matches) || [];
    const search = h("button", { class: "link small", onclick: () => { m.searching = true; viewMany(); } }, "Search by name");
    if (m.searching) pick = rowSearch(m);
    else if (!ms.length) pick = [h("div", { class: "meta warn-text" }, m.match && m.match.error ? m.match.error : m.match && m.match.typed ? "No client in HawkSoft matches “" + m.match.query + "”." : "No client in HawkSoft has this address."), search];
    else {
      const sel = h("select", { "aria-label": "Client for “" + (m.subject || "this e-mail") + "”", class: m.chosen ? null : "need" },
        m.chosen ? null : h("option", { value: "" }, ms.length > 1 ? "Choose the client (" + (m.match.total || ms.length) + " match)" : "Choose the client"),
        ms.map(c => h("option", { value: c.id, disabled: c.blocked }, c.name + (c.named ? " · named in e-mail" : "") + (c.blocked ? " (test mode: not allowed)" : ""))));
      sel.value = m.chosen ? m.chosen.id : "";
      sel.addEventListener("change", () => { m.chosen = ms.find(c => c.id === sel.value) || null; m.policy = ""; viewMany(); });
      const why = m.chosen && m.match.suggested === m.chosen.id && m.match.why ? "Picked: " + m.match.why + "." : null;
      let pol = null;
      if (m.chosen) {
        const pols = S.pols[m.chosen.id];
        if (pols === undefined) loadPolicies(m.chosen, viewMany);
        const n = Array.isArray(pols) ? pols.length : 0;
        pol = pols === undefined ? h("div", { class: "meta" }, "File on: reading the client's policies…")
          : !n ? h("div", { class: "meta" }, pols.error ? "File on the client (couldn't read the policies)." : "File on the client (it has no policies).")
          : h("div", { class: "fileon" }, h("span", { class: "meta" }, "File on"), policySelect(pols, m.policy, "File “" + (m.subject || "this e-mail") + "” on", (v) => { m.policy = v; viewMany(); }));
      }
      pick = [sel, pol, h("div", { class: "meta" }, why, why ? " " : null, search)];
    }
  }
  return h("div", { class: "mrow" + (m.on ? "" : " off") },
    h("label", { class: "mhead" }, box, h("div", {}, h("div", { class: "name" }, m.subject || "(no subject)"), meta ? h("div", { class: "meta" }, meta) : null)),
    pick ? h("div", { class: "mpick" }, pick) : null,
    atts);
}

function rowSearch(m) {
  const q = h("input", { type: "text", placeholder: "Client name, phone or policy number", "aria-label": "Search HawkSoft" });
  const run = async () => {
    const t = q.value.trim(); if (!t) return;
    const my = runId;
    busy("Searching HawkSoft for “" + t + "”…", "Nothing is changed.");
    let r;
    try { r = await watchBusy(api("/match", { query: t, fromName: m.fromName, subject: m.subject, text: m.text })); }
    catch (e) { if (my !== runId) return; return handle(e, viewMany); }
    if (my !== runId) return;
    m.searching = false;
    if (!r.ok) return viewError(r.error, null, viewMany);
    m.match = r; m.chosen = r.suggested ? r.matches.find(x => x.id === r.suggested) || null : null;
    viewMany();
  };
  q.addEventListener("keydown", (e) => { if (e.key === "Enter") run(); });
  setTimeout(() => q.focus(), 0);
  return [h("div", { class: "row" }, q, h("button", { class: "secondary", onclick: run }, "Search")),
    h("button", { class: "link small", onclick: () => { m.searching = false; viewMany(); } }, "Cancel")];
}

// Each ticked e-mail (.eml) and its attachments go to the helper once; changing the attachment switch stages them again.
async function stageMany() {
  const on = S.many.filter(m => m.on);
  for (let i = 0; i < on.length; i++) {
    const m = on[i], files = mailFiles(m), key = files.map(a => a.id).join("|");
    if (m.draftId && m.stagedKey === key) continue;
    busy("Reading the e-mails…", (i + 1) + " of " + on.length);
    const got = await withItem(m.itemId, async (it) => ({ eml: await officeAsync(cb => it.getAsFileAsync(cb)), attachments: await readContents(it, files) }));
    const sg = await api("/stage", { emlBase64: got.eml, subject: m.subject, attachments: got.attachments });
    if (!sg.ok) { viewError("“" + (m.subject || "(no subject)") + "”: " + sg.error, review, viewMany); return false; }
    if (m.stagedKey !== undefined && m.stagedKey !== key) { S.extra = {}; S.skip = new Set(); }      // other files: the answers (by file number) no longer fit
    m.draftId = sg.draftId; m.stagedKey = key;
  }
  return true;
}

async function reviewMany(restaged) {
  const my = runId, quick = !settings().showPlan;
  filingNow = true;
  try {
    if (!(await stageMany())) return;
    if (my !== runId) return;
    const on = S.many.filter(m => m.on);
    // one name per FreeBird task, in the batch's order: each e-mail, then its attachments
    S.taskNames = []; S.taskMail = [];
    for (const m of on) {
      S.taskNames.push("“" + (m.subject || "(no subject)") + "”"); S.taskMail.push(m);
      for (const a of mailFiles(m)) { S.taskNames.push("“" + a.name + "” from “" + (m.subject || "(no subject)") + "”"); S.taskMail.push(m); }
    }
    busy(quick ? "Filing in HawkSoft…" : "Asking FreeBird for its plan…", quick ? "Getting FreeBird ready." : "Nothing is filed yet.");
    const pl = await api("/plan", { extra: S.extra, skip: [...(S.skip || [])], items: on.map(m => ({ draftId: m.draftId, clientId: m.chosen.id, clientName: m.chosen.name, policy: m.policy || "", from: m.from, fromName: m.fromName, subject: m.subject, description: mailDesc(m), note: mailNote(m) })) });
    if (my !== runId) return;
    if (!pl.ok) {
      if (pl.code === "restage" && !restaged) { for (const m of on) m.draftId = null; return reviewMany(true); }
      if (pl.code === "nothing") return viewLeft();
      return viewError(pl.error + (pl.missing ? " (" + JSON.stringify(pl.missing) + ")" : ""), review, viewMany);
    }
    afterPlan(pl, quick);
  } catch (e) { if (my === runId) handle(e, review); }
}

function viewDoneMany(res, dupNote, verified) {
  S.view = "done";
  const on = S.many.filter(m => m.on);
  const clientsN = new Set(on.map(m => m.chosen.id)).size;
  const rows = on.map(m => {
    const n = (S.liveMail || S.taskMail).filter(x => x === m).length, left = S.taskMail.filter(x => x === m).length - n;
    return h("li", {}, h("strong", {}, m.subject || "(no subject)"), " → " + m.chosen.name + (m.policy ? ", policy " + m.policy : "") +
      (n === 0 ? ": left as it was (already there)" : n > 1 ? " (with " + plural(n - 1, "attachment", "attachments") + ")" : "") + (n && left ? "; " + left + " left as already there" : ""));
  });
  const needsYou = res.flatMap(r => r.needsYou || []);
  show(h("div", { class: "center" }, h("div", { class: "done-icon", "aria-hidden": "true" }, "✓"),
      h("h2", {}, "Filed " + plural(on.length, "e-mail", "e-mails") + " on " + plural(clientsN, "client", "clients")),
      h("p", { class: "muted" }, verified ? "FreeBird checked each one in HawkSoft." : "FreeBird filed them but could not read all of them back. Check the clients' attachments.")),
    dupNote ? note("warn", h("p", {}, dupNote)) : null,
    h("ul", { class: "list" }, rows),
    needsYou.length ? note("warn", h("p", {}, "Still for you:"), h("ul", { class: "list" }, [...new Set(needsYou)].map(d => h("li", {}, d)))) : null,
    nextMail());
}
