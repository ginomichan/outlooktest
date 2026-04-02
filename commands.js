function action(event) {
  if (event && typeof event.completed === "function") {
    event.completed();
  }
}

Office.actions.associate("action", action);
