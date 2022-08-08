import * as React from "react";
import { DefaultButton, PrimaryButton } from "@fluentui/react";

/* global Office, window */

/**
 * The message box app.
 *
 * @returns Quick and dirty markup with a headline and message.
 */
export default function App() {
  const params = new URLSearchParams(window.location.search);
  const headline = params.get("headline");
  const message = params.get("message");

  return (
    <>
      <h2>{headline}</h2>
      <p>{message}</p>
      <PrimaryButton text="OK" style={{ marginRight: 10 }} onClick={onOk} />
      <DefaultButton text="Cancel" onClick={onCancel} />
    </>
  );
}

function onOk() {
  Office.context.ui.messageParent("Ok");
}

function onCancel() {
  Office.context.ui.messageParent("Cancel");
}
