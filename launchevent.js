const ALLOWED_DOMAINS = [
  "maybank.com"
  // add more internal domains here if needed
];

function normalizeEmail(email) {
  return (email || "").trim().toLowerCase();
}

function isAllowedAddress(email) {
  const normalized = normalizeEmail(email);
  return ALLOWED_DOMAINS.some(domain => normalized.endsWith("@" + domain));
}

function getRecipientsAsync(recipientsField) {
  return new Promise((resolve) => {
    recipientsField.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || []);
      } else {
        resolve([]);
      }
    });
  });
}

async function onMessageSendHandler(event) {
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
      const uniqueAddresses = [...new Set(externalRecipients.map(r => normalizeEmail(r.emailAddress)))];
      const addressList = uniqueAddresses.join(", ");

      event.completed({
        allowEvent: false,
        errorMessage: `External recipient detected: ${addressList}`,
        errorMessageMarkdown:
          `This email contains one or more recipients outside **@maybank.com**.\n\n` +
          `External recipient(s): ${addressList}\n\n` +
          `Select **Send Anyway** to continue or **Don't Send** to review the recipients.`
      });
      return;
    }

    event.completed({ allowEvent: true });
  } catch (err) {
    event.completed({
      allowEvent: false,
      errorMessage: "Unable to validate recipients before sending."
    });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
