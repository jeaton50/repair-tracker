import { useState, useEffect } from "react";
import { PublicClientApplication } from "@azure/msal-browser";
import { msalConfig, loginRequest } from "../authConfig";
import toast from "react-hot-toast";

const msalInstance = new PublicClientApplication(msalConfig);

/**
 * Ensure we have a valid access token, prompting for login if needed
 */
export async function ensureAccessToken() {
  let account = msalInstance.getAllAccounts()[0];
  if (!account) {
    await msalInstance.loginPopup(loginRequest);
    account = msalInstance.getAllAccounts()[0];
  }
  try {
    const r = await msalInstance.acquireTokenSilent({ ...loginRequest, account });
    return r.accessToken;
  } catch {
    const r = await msalInstance.acquireTokenPopup({ ...loginRequest, account });
    return r.accessToken;
  }
}

/**
 * Custom hook for authentication state management
 */
export const useAuth = () => {
  const [isAuthenticated, setIsAuthenticated] = useState(false);
  const [accessToken, setAccessToken] = useState(null);
  const [userName, setUserName] = useState("");
  const [msalInitialized, setMsalInitialized] = useState(false);

  // Initialize MSAL
  useEffect(() => {
    msalInstance
      .initialize()
      .then(() => setMsalInitialized(true))
      .catch((e) => {
        console.error("MSAL initialization failed:", e);
        toast.error("Failed to initialize authentication");
      });
  }, []);

  // Check for existing session
  useEffect(() => {
    if (!msalInitialized) return;

    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
      msalInstance
        .acquireTokenSilent({ ...loginRequest, account: accounts[0] })
        .then((resp) => {
          setAccessToken(resp.accessToken);
          setIsAuthenticated(true);
          setUserName(accounts[0].name);
        })
        .catch((e) => console.error("Silent token acquisition failed:", e));
    }
  }, [msalInitialized]);

  const handleLogin = async () => {
    try {
      await msalInstance.loginPopup(loginRequest);
      const account = msalInstance.getAllAccounts()[0];
      const { accessToken: token } = await msalInstance.acquireTokenSilent({
        ...loginRequest,
        account
      });
      setAccessToken(token);
      setIsAuthenticated(true);
      setUserName(account.name);
      toast.success(`Welcome, ${account.name}!`);
    } catch (err) {
      console.error("Login failed:", err);
      toast.error("Failed to sign in to Microsoft. Please try again.");
    }
  };

  const handleLogout = () => {
    msalInstance.logoutPopup();
    setIsAuthenticated(false);
    setAccessToken(null);
    setUserName("");
    toast.success("Signed out successfully");
  };

  return {
    isAuthenticated,
    accessToken,
    userName,
    msalInitialized,
    handleLogin,
    handleLogout,
  };
};
