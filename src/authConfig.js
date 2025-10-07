// src/authConfig.js

// MSAL (Entra ID) configuration for the SPA
export const msalConfig = {
  auth: {
    clientId: "5764301a-9d69-4b4e-8bb6-83dbd4215e24",
    authority: "https://login.microsoftonline.com/b958ba39-da13-4553-a009-35b1649f44eb",
    // IMPORTANT: these must exactly match the SPA Redirect URIs in Entra ID (including trailing slash)
    redirectUri: import.meta.env.DEV
      ? "http://localhost:5173/repair-tracker/"
      : "https://jeaton50.github.io/repair-tracker/",
  },
  cache: {
    cacheLocation: "localStorage",   // keeps you signed in across reloads
    storeAuthStateInCookie: false,   // set true only if you must support old IE
  },
};

// Scopes requested during login / token acquisition
export const loginRequest = {
  scopes: [
    "openid",
    "profile",
    "email",
    "offline_access",
    "User.Read",
    "Files.Read", // use Files.ReadWrite if you plan to write back to OneDrive/SharePoint
  ],
};

// App-specific Graph config
export const graphConfig = {
  // Exact name/path of your OneDrive shortcut as it appears under "My files".
  // Examples: "RepairTracker"  or  "General/Repairs/RepairTracker"
  // You can override via .env:  VITE_RT_FOLDER=Repairs/RepairTracker
  folderPath: import.meta.env.VITE_RT_FOLDER ?? "RepairTracker",
};
