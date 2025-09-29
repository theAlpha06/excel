import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";
import { FluentProvider, webLightTheme } from "@fluentui/react-components";

const rootElement = document.getElementById("container");
const root = rootElement ? createRoot(rootElement) : undefined;

export const PRODUCTION_CLIENT = "https://dmpl-connector-cxgdewgsgce5g9ff.canadacentral-01.azurewebsites.net/";
export const PRODUCTION_SERVER = "https://salesforce-connecter-c5hxfvbhgxfgdbbr.canadacentral-01.azurewebsites.net/";
export const LOCAL_SERVER = "http://localhost:5000";
export const LOCAL_CLIENT = "https://localhost:3000";

Office.onReady(() => {
  root?.render(
    <FluentProvider theme={webLightTheme}>
      <App />
    </FluentProvider>
  );
});

if (module.hot) {
  module.hot.accept("./components/App", () => {
    const NextApp = require("./components/App").default;
    root?.render(NextApp);
  });
}