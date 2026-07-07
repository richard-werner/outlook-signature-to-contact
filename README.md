# Signature to Contact — Outlook Add-in

A cross-platform Outlook add-in. Open a received email, click **Save sender** in
the ribbon, review the parsed contact, and save it to your Outlook Contacts.

- **Reads** the open message via Office.js (sender name/email + body text).
- **Parses** the signature with client-side heuristics (no API key, no backend).
- **Writes** the contact via Microsoft Graph (`POST /me/contacts`), after
  deduplicating on email address.
- **Auth** uses MSAL Nested App Authentication — no middle-tier server required.

Because it's a web add-in, one deployment runs on **Outlook for Windows, Mac, and
web** (and new Outlook for Windows). No VBA/COM, no per-OS build.

---

## Files

```
manifest.xml          Add-in manifest (defines the ribbon button)
src/taskpane.html     Task pane markup; loads Office.js + MSAL
src/taskpane.css      Styling
src/taskpane.js       Read -> parse -> auth -> dedup -> create
src/icon-*.png        You supply these (16/32/64/80/128 px)
```

---

## Two things only you can do

### 1. Host the `src` files over HTTPS

Add-ins must be served over HTTPS. Any static host works: Azure Static Web Apps,
Azure Blob + CDN, GitHub Pages, an internal HTTPS server, etc. For local testing,
`office-addin-dev-certs` + a localhost HTTPS server is the usual route.

Then do a find-and-replace of `https://YOUR-HOST` throughout `manifest.xml` with
your real host, and add small PNG icons at the referenced paths.

### 2. Register an Azure AD app (for the Graph call)

In **Entra admin center → App registrations → New registration**:

1. Set a **Single-page application (SPA)** redirect URI:
   `brk-multihub://YOUR-HOST` (the NAA broker scheme) and also your task pane URL.
2. Note the **Application (client) ID** and **Directory (tenant) ID**.
3. **API permissions → Microsoft Graph → Delegated → `Contacts.ReadWrite`.**
   Grant admin consent if your tenant requires it.
4. Put those two IDs into `MSAL_CONFIG` at the top of `src/taskpane.js`.
5. Generate a GUID for `<Id>` in `manifest.xml` (any GUID tool / `uuidgen`).

---

## Deploy

- **Test (sideload):** in Outlook web → Settings → *Manage add-ins* / *Get add-ins
  → My add-ins → Add a custom add-in → Add from file*, and pick `manifest.xml`.
- **Company-wide:** **Microsoft 365 admin center → Settings → Integrated apps →
  Upload custom app**, upload `manifest.xml`, and assign to a group. This is the
  Windows-and-Mac, everyone-gets-it path — the same governance gate you'd expect
  wearing your IT hat.

Each user consents to `Contacts.ReadWrite` for their own mailbox on first use.
Contacts are written to that user's own Contacts, which is why it behaves
identically across platforms.

---

## Parsing: how good, and how to upgrade

The heuristic parser anchors on the sender name/email that Office.js gives you
(reliable), then mines the top of the body for phone/title/company and stops at
the first quoted-reply marker so it doesn't scrape signatures from earlier
messages in a thread.

It handles clean **text** signatures well. Two known limits:

- **Image-only signatures** yield little — there's no text to read. An OCR step
  would be needed.
- **Unusual layouts** may mis-assign title vs. company. That's why the task pane
  shows an editable form before saving — review, then save.

**To upgrade to LLM parsing later:** replace the body of `parseSignature()` in
`taskpane.js` with a `fetch` to a small backend proxy you control (which holds
the model API key server-side and returns the same field shape). Don't call an
LLM API directly from the add-in — that would expose the key in the browser.

---

## Notes

- MSAL and Office.js load from CDN in `taskpane.html`; both run in the user's
  browser at runtime, so no build step is required for this add-in.
- Graph API and MSAL versions move; verify current guidance if something in the
  auth flow behaves unexpectedly.
