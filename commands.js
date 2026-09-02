/*
 * "Send to Sift" — Outlook add-in commands.
 *
 * Each ribbon button runs one function here. The function reads the open
 * message with Office.js (which works inside Outlook's auth context, so it can
 * pull attachment *bytes* — the thing a drag out of New Outlook could not do),
 * then POSTs to Sift's local bridge at http://127.0.0.1:8137. Calling http
 * localhost from this https page is allowed because localhost is a trustworthy
 * origin; the bridge injects Sift's token and forwards to the sidecar.
 *
 * Four buttons: send attachments, send email (stub note), summarize email as a
 * note, and everything at once. Sift routes to the right effort by itself when
 * the message carries an SFT- code — the add-in never knows the mapping.
 */

/* eslint-disable no-undef */

const BRIDGE = "http://127.0.0.1:8137";

Office.onReady();

/** Post to the bridge, turning transport/HTTP failures into a friendly line. */
async function postToSift(path, payload) {
  let resp;
  try {
    resp = await fetch(BRIDGE + path, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
  } catch (e) {
    throw new Error("Couldn't reach Sift — is the app running?");
  }
  if (!resp.ok) {
    let detail = "";
    try {
      const j = await resp.json();
      detail = j && j.error && j.error.message ? `: ${j.error.message}` : "";
    } catch (_) {
      /* ignore */
    }
    throw new Error(`Sift returned ${resp.status}${detail}`);
  }
  return resp.json();
}

/** Read the message body as plain text (for SFT- code detection + summaries). */
function getBodyText(item) {
  return new Promise((resolve) => {
    if (!item.body) return resolve("");
    item.body.getAsync(Office.CoercionType.Text, (res) => {
      resolve(res.status === Office.AsyncResultStatus.Succeeded ? res.value || "" : "");
    });
  });
}

/** Read one attachment's bytes (base64) from the current message. */
function getAttachmentContent(item, id) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(id, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error || new Error("attachment read failed"));
    });
  });
}

/** A deep link that reopens THIS message in New Outlook / OWA. Office.js gives
 * no web link directly, so we build the OWA read-item URL from the item id.
 * Best-effort: if there's no item id (rare), reopen may not resolve. */
function buildWebLink(item) {
  try {
    if (item.itemId) {
      return (
        "https://outlook.office365.com/owa/?ItemID=" +
        encodeURIComponent(item.itemId) +
        "&exvsurl=1&viewmodel=ReadMessageItem"
      );
    }
  } catch (_) {
    /* ignore */
  }
  return "";
}

/** The message metadata Sift records (sender, date, ids, reopen link). */
function gatherEmail(item) {
  let sender = "";
  if (item.from) {
    sender = item.from.displayName || item.from.emailAddress || "";
    if (item.from.displayName && item.from.emailAddress) {
      sender = `${item.from.displayName} <${item.from.emailAddress}>`;
    }
  }
  let received = "";
  try {
    if (item.dateTimeCreated) received = new Date(item.dateTimeCreated).toISOString();
  } catch (_) {
    /* ignore */
  }
  return {
    subject: item.subject || "",
    sender,
    received_at: received,
    internet_message_id: item.internetMessageId || "",
    conversation_id: item.conversationId || "",
    web_link: buildWebLink(item),
  };
}

/** Collect this message's file attachments as base64. */
async function gatherAttachments(item) {
  const files = (item.attachments || []).filter(
    (a) => a.attachmentType === Office.MailboxEnums.AttachmentType.File && !a.isInline,
  );
  const out = [];
  for (const a of files) {
    const content = await getAttachmentContent(item, a.id);
    out.push({ name: a.name, content_base64: content.content });
  }
  return out;
}

/** Show a status banner on the message. */
function notify(text) {
  const item = Office.context.mailbox.item;
  if (!item || !item.notificationMessages) return;
  item.notificationMessages.replaceAsync("sift-status", {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: String(text).slice(0, 150),
    persistent: false,
  });
}

function effortSuffix(res) {
  return res && res.effort_tag ? ` (${res.effort_tag.replace(/^Effort\//, "")})` : "";
}

/** Button: Send attachments to Sift. */
async function sendAttachmentsToSift(event) {
  try {
    const item = Office.context.mailbox.item;
    const files = await gatherAttachments(item);
    if (files.length === 0) {
      notify("No file attachments on this email.");
      return event.completed();
    }
    const res = await postToSift("/addin/attachments", {
      files,
      subject: item.subject || "",
      email_body: await getBodyText(item),
    });
    const where = res && res.effort_tag
      ? `filed to ${res.effort_tag.replace(/^Effort\//, "")}`
      : "in the filing queue for review";
    notify(`Sent ${files.length} attachment${files.length === 1 ? "" : "s"} to Sift — ${where}.`);
  } catch (e) {
    notify(e && e.message ? e.message : "Send to Sift failed.");
  }
  event.completed();
}

/** Button: Send email to Sift (drops it in Sift's Emails inbox to triage). */
async function sendEmailToSift(event) {
  try {
    const item = Office.context.mailbox.item;
    const payload = gatherEmail(item);
    payload.email_body = await getBodyText(item);
    const res = await postToSift("/addin/email", payload);
    const where = res && res.effort_tag
      ? `Sift's inbox and attached to ${res.effort_tag.replace(/^Effort\//, "")}`
      : "Sift's Emails inbox";
    notify(`Sent this email to ${where}.`);
  } catch (e) {
    notify(e && e.message ? e.message : "Send to Sift failed.");
  }
  event.completed();
}

/** Button: Summarize email as a note in Sift.
 *
 * The summary is an AI call that takes a few seconds — too long for Outlook to
 * hold a ribbon command open — so Sift ACKs immediately and writes the note in
 * the background. We tell the user it's on its way rather than waiting. */
async function summarizeEmailToNote(event) {
  try {
    const item = Office.context.mailbox.item;
    const payload = gatherEmail(item);
    payload.email_body = await getBodyText(item);
    await postToSift("/addin/note", payload);
    notify("Summarizing this email — the note will appear in Sift's Notes shortly.");
  } catch (e) {
    notify(e && e.message ? e.message : "Send to Sift failed.");
  }
  event.completed();
}

/** Button: Send everything — attachments filed + a summarized note. */
async function sendEverythingToSift(event) {
  try {
    const item = Office.context.mailbox.item;
    const email = gatherEmail(item);
    email.email_body = await getBodyText(item);

    const files = await gatherAttachments(item);
    const parts = [];
    if (files.length > 0) {
      await postToSift("/addin/attachments", {
        files,
        subject: email.subject,
        email_body: email.email_body,
      });
      parts.push(`${files.length} attachment${files.length === 1 ? "" : "s"} filed`);
    }

    const er = await postToSift("/addin/email", email);
    parts.push("email in inbox");

    await postToSift("/addin/note", email);
    parts.push("summarized note on its way");
    notify(`Sent to Sift — ${parts.join(", ")}${effortSuffix(er)}.`);
  } catch (e) {
    notify(e && e.message ? e.message : "Send to Sift failed.");
  }
  event.completed();
}

// Register each function so the manifest's <FunctionName> can reach it.
Office.actions.associate("sendAttachmentsToSift", sendAttachmentsToSift);
Office.actions.associate("sendEmailToSift", sendEmailToSift);
Office.actions.associate("summarizeEmailToNote", summarizeEmailToNote);
Office.actions.associate("sendEverythingToSift", sendEverythingToSift);
