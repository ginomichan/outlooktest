const ALLOWED_DOMAINS = ["maybank.com"];

function normalizeEmail(email) {
  return (email || "").trim().toLowerCase();
}

function isAllowedAddress(email) {
  const normalized = normalizeEmail(email);
  return ALLOWED_DOMAINS.some(domain => normalized.endsWith("@" + domain));
}

function getRecipientsAsync(field) {
  return new Promise((resolve) => {
    field.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        resolve([]);
      }
    });
  });
}

async function onRecipientsChangedHandler(event) {
  try {
    const item = Office.context.mailbox.item;
    const [toRecipients, ccRecipients, bccRecipients] = await Promise.all([
      getRecipientsAsync(item.to),
      getRecipientsAsync(item.cc),
      getRecipientsAsync(item.bcc)
    ]);

    const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];
    const externalRecipients = allRecipients.filter(r => !isAllowedAddress(r.emailAddress));

    if (externalRecipients.length > 0) {
      item.notificationMessages.replaceAsync("external-warning", {
        type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
        message: "External recipient detected. Please review recipients before sending.",
        icon: "Icon.64",
        persistent: true
      }, () => event.completed());
    } else {
      item.notificationMessages.removeAsync("external-warning", () => event.completed());
    }
  } catch (e) {
    event.completed();
  }
}

Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
