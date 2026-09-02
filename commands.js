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
 * Milestone 1 ships "Send attachments". Email + summarise-to-note buttons
 * follow once this pipe is confirmed.
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

/** Read one attachment's bytes (base64) from the current message. */
function getAttachmentContent(item, id) {
  return new Promise((resolve, reject) => {
    item.getAttachmentContentAsync(id, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded) resolve(res.value);
      else reject(res.error || new Error("attachment read failed"));
    });
  });
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

/** Button: Send attachments to Sift. */
async function sendAttachmentsToSift(event) {
  try {
    const item = Office.context.mailbox.item;
    const all = item.attachments || [];
    const files = all.filter(
      (a) =>
        a.attachmentType === Office.MailboxEnums.AttachmentType.File &&
        !a.isInline,
    );
    if (files.length === 0) {
      notify("No file attachments on this email.");
      return event.completed();
    }

    const payload = { files: [] };
    for (const a of files) {
      const content = await getAttachmentContent(item, a.id);
      // File attachments come back base64-encoded (content.format === "base64").
      payload.files.push({ name: a.name, content_base64: content.content });
    }

    await postToSift("/addin/attachments", payload);
    notify(
      `Sent ${payload.files.length} attachment${
        payload.files.length === 1 ? "" : "s"
      } to Sift — review them in the filing queue.`,
    );
  } catch (e) {
    notify(e && e.message ? e.message : "Send to Sift failed.");
  }
  event.completed();
}

// Register each function so the manifest's <FunctionName> can reach it.
Office.actions.associate("sendAttachmentsToSift", sendAttachmentsToSift);
