/* global Office, console, window */

var dialog: Office.Dialog;

export function openDialog(path: string, params?: Record<string, string>, options?: Office.DialogOptions) {
  const query = Object.keys(params ?? {})
    .map((key) => `${key}=${encodeURIComponent(params[key])}`)
    .join("&");

  const startAddress = window.location.origin + (query ? `/${path}?${query}` : `/${path}`);
  Office.context.ui.displayDialogAsync(startAddress, options, dialogCallback);
}

function dialogCallback(asyncResult: Office.AsyncResult<Office.Dialog>) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    // In addition to general system errors, there are 3 specific errors for
    // displayDialogAsync that you can handle individually.
    switch (asyncResult.error.code) {
      case 12004:
        showNotification("Domain is not trusted");
        break;
      case 12005:
        showNotification("HTTPS is required");
        break;
      case 12007:
        showNotification("A dialog is already opened.");
        break;
      default:
        showNotification(asyncResult.error.message);
        break;
    }
  } else {
    dialog = asyncResult.value;

    // Messages are sent by developers programatically from the dialog using office.context.ui.messageParent(...)
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, messageHandler);

    // Events are sent by the platform in response to user actions or errors.
    // For example, the dialog is closed via the 'x' button.
    dialog.addEventHandler(Office.EventType.DialogEventReceived, eventHandler);
  }
}

function messageHandler(arg: { message: string; origin: string | undefined }) {
  dialog.close();
  showNotification(arg.message);
}

function eventHandler(arg: { error: number }) {
  // In addition to general system errors, there are 2 specific errors
  // and one event that you can handle individually.
  switch (arg.error) {
    case 12002:
      showNotification("Cannot load URL, no such page or bad URL syntax.");
      break;
    case 12003:
      showNotification("HTTPS is required.");
      break;
    case 12006:
      // The dialog was closed, typically because the user the pressed X button.
      showNotification("Dialog closed by user");
      break;
    default:
      showNotification("Undefined error in dialog window");
      break;
  }
}

function showNotification(message: string) {
  console.log(message);
}
