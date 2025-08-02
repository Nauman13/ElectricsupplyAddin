import { Configuration } from "@azure/msal-browser";

export const msalConfig: Configuration = {
  auth: {
    clientId: "a8f0615d-19b4-4c1d-b578-3f754d19e56c", 
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000/",
  },
  cache: {
    cacheLocation: "localStorage",
  },
};
