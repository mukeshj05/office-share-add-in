import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";
// import { MsalProvider } from "@azure/msal-react";
// import { msalInstance } from "../utils/authConfig";

/* global document, Office, module, require */

const title = "Legistify Add-in";

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

/* Render application after Office initializes */
Office.onReady(() => {
  root?.render(
    // <MsalProvider instance={msalInstance}>
    <FluentProvider theme={webLightTheme}>
      <App title={title} />
    </FluentProvider>
    // </MsalProvider>
  );
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}