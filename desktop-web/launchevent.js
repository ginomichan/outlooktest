function onMessageSendHandler(event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Test popup from Maybank add-in"
  });
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
