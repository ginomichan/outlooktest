function normalizeEmail(email) {
  return (email || "").trim().toLowerCase();
}

function isInternalAddress(email) {
  return normalizeEmail(email).endsWith("@maybank.com");
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

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    const [toRecipients, ccRecipients, bccRecipients] = await Promise.all([
      getRecipientsAsync(item.to),
      getRecipientsAsync(item.cc),
      getRecipientsAsync(item.bcc)
    ]);

    const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];
    const externalRecipients = allRecipients.filter(
      (r) => !isInternalAddress(r.emailAddress)
    );

    if (externalRecipients.length > 0) {
      const list = [...new Set(externalRecipients.map(r => normalizeEmail(r.emailAddress)))].join(", ");

      event.completed({
        allowEvent: false,
        errorMessage: `External recipient detected: ${list}`,
        errorMessageMarkdown:
          `This email contains one or more recipients outside **@maybank.com**.\n\n` +
          `External recipient(s): ${list}\n\n` +
          `Select **Send Anyway** to continue or **Don't Send** to review recipients.`
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
