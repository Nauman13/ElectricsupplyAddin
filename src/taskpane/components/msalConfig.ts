import { Configuration, BrowserUtils } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: "a8f0615d-19b4-4c1d-b578-3f754d19e56c",
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://electricsupply-addin.vercel.app/",
  },
  cache: {
    cacheLocation: "localStorage",
  },
  system: {
    iframeHashTimeout: 10000, // Increase for desktop
    navigateFrameWait: 500,
    windowHashTimeout: 10000,
  },
};

// Initialization handler removed: setInteractionInProgress does not exist on BrowserUtils.
