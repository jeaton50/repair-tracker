// src/authConfig.js

export const msalConfig = {
  auth: {
    clientId: "5764301a-9d69-4b4e-8bb6-83dbd4215e24",
    authority: "https://login.microsoftonline.com/b958ba39-da13-4553-a009-35b1649f44eb",
    // Must exactly match your Entra "SPA" redirect URIs (with trailing slash)
    redirectUri: import.meta.env.DEV
      ? "http://localhost:5173/repair-tracker/"
      : "https://jeaton50.github.io/repair-tracker/",
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false,
  },
};

// Request WRITE scopes (admin consent required) so uploads work.
export const loginRequest = {
  scopes: [
    "openid",
    "profile",
    "email",
    "offline_access",
    "User.Read",
    // Write permissions for SharePoint/OneDrive via Graph:
    "Files.ReadWrite.All",
    "Sites.ReadWrite.All",
  ],
};

export const graphConfig = {
  // SharePoint site/library coordinates
  spHostname:  import.meta.env.VITE_SP_HOSTNAME ?? "rentexinc.sharepoint.com",
  spSitePath:  import.meta.env.VITE_SP_SITEPATH ?? "/sites/ProductManagers",
  spBasePath:  import.meta.env.VITE_SP_BASE     ?? "General/Repairs/RepairTracker",

  // Filenames
  ticketsFile: import.meta.env.VITE_SP_TICKETS  ?? "ticket_list.xlsx",
  reportsFile: import.meta.env.VITE_SP_REPORTS  ?? "repair_report.xlsx",
  mappingFile: import.meta.env.VITE_SP_MAPPING  ?? "category_mapping.json",
};
