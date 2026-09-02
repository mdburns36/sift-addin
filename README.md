# Send to Sift — Outlook add-in

Ribbon buttons in **New Outlook** (and Outlook on the web) that send the open
email into Sift. No dragging, no saving to disk first. Four buttons:

- **Send attachments to Sift** — files the email's attachments into your Keep
  inbox / filing queue.
- **Send email to Sift** — records the email as a note with a link that reopens
  the live thread.
- **Summarize email as note** — an AI summary of the email as a Sift note.
- **Send everything to Sift** — attachments filed **and** a summarized note.

**Effort routing:** if the email (subject or body) carries an `SFT-<code>` for
an effort, Sift files/tags it to that effort automatically — you don't pick.
Otherwise attachments go to the review queue and notes to your default vault.

## How it works

- The buttons run a small script (`commands.js`) that reads the message with
  Office.js — inside Outlook's own login, so it gets the actual attachment bytes
  and body text.
- It POSTs to Sift's local bridge at `http://127.0.0.1:8137`, which Sift runs
  whenever it's open. The bridge adds Sift's auth token, summarizes when needed,
  and files/records everything.
- So **Sift must be running** on the same PC when you click a button. The email
  and note buttons also need **vault writes enabled** (Settings → Obsidian).

## One-time setup

### 1. Host these files over HTTPS (GitHub Pages — free, always on)

Outlook requires the add-in's web files to be served over HTTPS. GitHub Pages is
the simplest host.

The manifest and icons are already set for `https://mdburns36.github.io/sift-addin`,
so the repo must be named **`sift-addin`** under your `mdburns36` account.

1. Create a new **public** GitHub repo named exactly **`sift-addin`**.
2. Upload the contents of this `outlook-addin/` folder to the repo root
   (`manifest.xml`, `commands.html`, `commands.js`, and the `assets/` folder).
3. Repo **Settings → Pages** → Source: `Deploy from a branch` → branch `main`,
   folder `/ (root)` → Save. After a minute your site is live at
   `https://mdburns36.github.io/sift-addin/`.
4. Sanity check: open `https://mdburns36.github.io/sift-addin/manifest.xml` in a
   browser — you should see the XML. Then sideload (below).

### 2. Sideload the manifest into New Outlook

1. In New Outlook, open **Settings (gear) → General → Manage add-ins**
   (or **Get Add-ins** on the ribbon). This opens the add-ins dialog.
2. Choose **My add-ins → Add a custom add-in → Add from URL…**
3. Paste your manifest URL:
   `https://mdburns36.github.io/sift-addin/manifest.xml`
4. Confirm. The **Send attachments to Sift** button appears when you open an
   email (look on the ribbon / the `…` overflow of a read message).

## Using it

1. Make sure Sift is running.
2. Open an email that has attachments.
3. Click **Send attachments to Sift**.
4. A banner confirms how many were sent. Open Sift → filing queue to tag/file them.

## Troubleshooting

- **"Couldn't reach Sift"** — Sift isn't running, or the bridge didn't start.
  Open Sift and try again. The bridge listens on `127.0.0.1:8137`.
- **Button missing** — re-open a message; some layouts tuck it under `…` (More
  actions). Re-sideload if it never shows.
- **Nothing filed** — attachments land in the root named "Keep" (or your chosen
  library drop root). Make sure you have a Keep root in Sift.
