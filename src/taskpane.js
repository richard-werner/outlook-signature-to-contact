/* =========================================================================
   Signature to Contact — Outlook add-in
   Flow: read open message -> heuristic signature parse -> confirm form ->
         MSAL (Nested App Auth) token -> Graph dedup -> Graph create contact.
   ========================================================================= */

// ---- CONFIG: fill these in after you register the Azure AD app ----------
const MSAL_CONFIG = {
  auth: {
    clientId: "REPLACE-WITH-YOUR-AZURE-APP-CLIENT-ID",
    // Use your tenant ID for single-tenant, or "common" for multi-tenant:
    authority: "https://login.microsoftonline.com/REPLACE-WITH-YOUR-TENANT-ID",
  },
};
const GRAPH_SCOPES = ["Contacts.ReadWrite"];
// ------------------------------------------------------------------------

let pca = null; // MSAL nestable public client, created on Office ready

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("saveBtn").addEventListener("click", onSave);
    document.getElementById("rescanBtn").addEventListener("click", loadCurrentMessage);
    initAuth().then(loadCurrentMessage);
  }
});

async function initAuth() {
  try {
    // Nested App Auth: lets the add-in get Graph tokens without a middle-tier.
    pca = await msal.createNestablePublicClientApplication(MSAL_CONFIG);
  } catch (e) {
    // Fall back to a standard PCA (popup) if NAA isn't available on this client.
    pca = new msal.PublicClientApplication(MSAL_CONFIG);
  }
}

/* ---------------------------------------------------------------------- */
/* 1. Read the current message                                            */
/* ---------------------------------------------------------------------- */

function loadCurrentMessage() {
  setStatus("Reading the message…");
  hide("contactForm");
  hide("result");

  const item = Office.context.mailbox.item;
  if (!item || !item.from) {
    setStatus("Open a received email, then reopen this pane.");
    return;
  }

  // Sender name + email come straight from Office.js and are the most reliable
  // anchor — we don't have to guess them out of the signature text.
  const from = item.from; // { displayName, emailAddress }

  item.body.getAsync(Office.CoercionType.Text, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      setStatus("Couldn't read the message body.");
      return;
    }
    const parsed = parseSignature(res.value || "", from);
    fillForm(parsed);
    show("contactForm");
    hide("status");
  });
}

/* ---------------------------------------------------------------------- */
/* 2. Heuristic signature extraction                                      */
/*    (Swap this whole function for a call to your own backend proxy if    */
/*     you later want LLM-based parsing — keep the return shape the same.)  */
/* ---------------------------------------------------------------------- */

function parseSignature(bodyText, from) {
  // Only look at the TOP block — stop at the first quoted/forwarded marker so
  // we don't scrape signatures from earlier messages in the thread.
  const cutoffs = [
    /^\s*from:\s.+sent:/im,
    /^\s*-{2,}\s*original message\s*-{2,}/im,
    /^\s*on .+ wrote:\s*$/im,
    /^_{5,}\s*$/m,
    /^\s*from:\s.+@.+/im,
  ];
  let top = bodyText;
  for (const c of cutoffs) {
    const m = top.match(c);
    if (m && m.index > 40) top = top.slice(0, m.index);
  }

  const lines = top.split(/\r?\n/).map((l) => l.trim());

  // Email + phones by regex.
  const emailRe = /[a-z0-9._%+\-]+@[a-z0-9.\-]+\.[a-z]{2,}/i;
  const phoneRe = /(\+?\d[\d\s().\-]{7,}\d)/g;

  const contact = {
    givenName: "",
    surname: "",
    displayName: from.displayName || "",
    jobTitle: "",
    companyName: "",
    emailAddress: from.emailAddress || "",
    businessPhone: "",
    mobilePhone: "",
  };

  // Split display name into given / surname.
  if (contact.displayName && contact.displayName.indexOf("@") === -1) {
    const parts = contact.displayName.replace(/\(.*?\)/g, "").trim().split(/\s+/);
    if (parts.length >= 2) {
      contact.givenName = parts[0];
      contact.surname = parts.slice(1).join(" ");
    } else {
      contact.givenName = parts[0] || "";
    }
  }

  // Phones: classify by nearby keyword; otherwise first = business, second = mobile.
  const foundPhones = [];
  for (const line of lines) {
    const matches = line.match(phoneRe);
    if (!matches) continue;
    for (const raw of matches) {
      const num = raw.trim();
      if (num.replace(/\D/g, "").length < 7) continue;
      const lc = line.toLowerCase();
      const kind = /\b(mobile|mob|cell|m:)\b/.test(lc) ? "mobile"
                 : /\b(office|work|tel|direct|o:|t:|d:)\b/.test(lc) ? "business"
                 : "";
      foundPhones.push({ num, kind });
    }
  }
  for (const p of foundPhones) {
    if (p.kind === "mobile" && !contact.mobilePhone) contact.mobilePhone = p.num;
    if (p.kind === "business" && !contact.businessPhone) contact.businessPhone = p.num;
  }
  for (const p of foundPhones) {
    if (!p.kind) {
      if (!contact.businessPhone) contact.businessPhone = p.num;
      else if (!contact.mobilePhone) contact.mobilePhone = p.num;
    }
  }

  // Job title / company: scan the lines just after the name line.
  const nameIdx = lines.findIndex(
    (l) => contact.displayName && l && l.toLowerCase().includes(contact.displayName.toLowerCase())
  );
  const titleWords = /(manager|director|president|vp|officer|engineer|lead|head|analyst|consultant|specialist|coordinator|founder|ceo|cto|cfo|coo|owner|associate|representative|rep|executive|administrator)/i;
  const start = nameIdx >= 0 ? nameIdx + 1 : 0;
  for (let i = start; i < Math.min(lines.length, start + 6); i++) {
    const l = lines[i];
    if (!l) continue;
    if (emailRe.test(l) || phoneRe.test(l)) continue;
    if (!contact.jobTitle && titleWords.test(l) && l.length < 60) {
      contact.jobTitle = l.replace(/\s*[|,]\s*.*$/, "").trim();
      continue;
    }
    // A short line with a capital letter that isn't the title is likely the company.
    if (!contact.companyName && /[A-Z]/.test(l) && l.length < 60 && !titleWords.test(l)) {
      contact.companyName = l.replace(/\s*[|].*$/, "").trim();
    }
  }

  return contact;
}

/* ---------------------------------------------------------------------- */
/* 3. Form helpers                                                        */
/* ---------------------------------------------------------------------- */

const FIELDS = ["givenName","surname","jobTitle","companyName","emailAddress","businessPhone","mobilePhone"];

function fillForm(c) {
  FIELDS.forEach((f) => (document.getElementById(f).value = c[f] || ""));
}
function readForm() {
  const c = {};
  FIELDS.forEach((f) => (c[f] = document.getElementById(f).value.trim()));
  c.displayName = [c.givenName, c.surname].filter(Boolean).join(" ") || c.emailAddress;
  return c;
}

/* ---------------------------------------------------------------------- */
/* 4. Auth + Graph: dedup then create                                     */
/* ---------------------------------------------------------------------- */

async function getToken() {
  const request = { scopes: GRAPH_SCOPES };
  try {
    const r = await pca.acquireTokenSilent(request);
    return r.accessToken;
  } catch (_) {
    const r = await pca.acquireTokenPopup(request);
    return r.accessToken;
  }
}

async function onSave() {
  const c = readForm();
  if (!c.emailAddress) return showResult("err", "An email address is required to save a contact.");

  const btn = document.getElementById("saveBtn");
  btn.disabled = true;
  showResult("", "Saving…");

  try {
    const token = await getToken();

    // ---- dedup on email address ----
    const existing = await graph(
      token,
      "GET",
      `/me/contacts?$filter=emailAddresses/any(a:a/address eq '${encodeURIComponent(c.emailAddress)}')&$select=id,displayName`
    );
    if (existing && existing.value && existing.value.length > 0) {
      showResult("ok", `"${existing.value[0].displayName || c.emailAddress}" is already in your Contacts — not creating a duplicate.`);
      btn.disabled = false;
      return;
    }

    // ---- create ----
    const payload = {
      givenName: c.givenName || undefined,
      surname: c.surname || undefined,
      displayName: c.displayName || undefined,
      jobTitle: c.jobTitle || undefined,
      companyName: c.companyName || undefined,
      emailAddresses: [{ address: c.emailAddress, name: c.displayName }],
      businessPhones: c.businessPhone ? [c.businessPhone] : [],
      mobilePhone: c.mobilePhone || undefined,
    };
    await graph(token, "POST", "/me/contacts", payload);
    showResult("ok", `Saved ${c.displayName} to your Contacts.`);
  } catch (e) {
    showResult("err", "Couldn't save: " + (e && e.message ? e.message : e));
  } finally {
    btn.disabled = false;
  }
}

async function graph(token, method, path, body) {
  const res = await fetch("https://graph.microsoft.com/v1.0" + path, {
    method,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "application/json",
    },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (!res.ok) {
    let detail = res.status + " " + res.statusText;
    try { const j = await res.json(); if (j.error) detail = j.error.message; } catch (_) {}
    throw new Error(detail);
  }
  return res.status === 204 ? null : res.json();
}

/* ---------------------------------------------------------------------- */
/* UI helpers                                                             */
/* ---------------------------------------------------------------------- */

function setStatus(t) { const e = document.getElementById("status"); e.textContent = t; show("status"); }
function show(id) { document.getElementById(id).classList.remove("hidden"); }
function hide(id) { document.getElementById(id).classList.add("hidden"); }
function showResult(kind, msg) {
  const e = document.getElementById("result");
  e.textContent = msg;
  e.className = "result" + (kind ? " " + kind : "");
  show("result");
}
