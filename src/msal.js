// src/msal.js
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "./authConfig";

export const msalInstance = new PublicClientApplication(msalConfig);

export async function getAccessToken() {
  let account = msalInstance.getAllAccounts()[0];
  if (!account) {
    await msalInstance.loginPopup(loginRequest);
    account = msalInstance.getAllAccounts()[0];
  }
  try {
    const { accessToken } = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    return accessToken;
  } catch {
    const { accessToken } = await msalInstance.acquireTokenPopup({ ...loginRequest, account });
    return accessToken;
  }
}
