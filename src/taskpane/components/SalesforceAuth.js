import React, { useState, useEffect } from "react";
import { 
  DefaultButton, 
  MessageBar, 
  MessageBarType, 
  Spinner, 
  SpinnerSize,
  Persona,
  PersonaSize,
  PersonaPresence,
  Stack,
  Text,
  FontWeights,
} from "@fluentui/react";

const CONSUMER_KEY =
  "3MVG9GCMQoQ6rpzQm797ZYHBww6Iye8XtvMqgddHtVjTr83ocGxvaeyGNJuUUf_dDrt688z_OV9wi8c2RC90w";
const REDIRECT_URI = "https://excel-p9o4.onrender.com/auth-callback.html";

const buttonStackTokens = {
  childrenGap: 12
};

const profileStackTokens = {
  childrenGap: 16
};

const SalesforceAuth = ({ setIsConnected, isConnected, orgType }) => {
  const [dialog, setDialog] = useState(null);
  const [isLoading, setIsLoading] = useState(false);
  const [error, setError] = useState(null);
  const [user, setUser] = useState();

  const LOGIN_URL =
    orgType === "production" ? "https://login.salesforce.com" : "https://test.salesforce.com";

  useEffect(() => {
    checkConnectionStatus();
  }, []);

  const checkConnectionStatus = async () => {
    try {
      setIsLoading(true);
      await Office.context.document.settings.refreshAsync();
      const accessToken = Office.context.document.settings.get("salesforce_access_token");
      const instanceUrl = Office.context.document.settings.get("salesforce_instance_url");

      if (accessToken && instanceUrl) {
        try {
          const response = await fetch("https://salesforce-connecter-c5hxfvbhgxfgdbbr.canadacentral-01.azurewebsites.net/salesforce/check-status", {
            method: "GET",
            headers: {
              "sf-instance-url": instanceUrl,
              "sf-access-token": accessToken,
              "Content-Type": "application/json",
            },
          });

          const data = await response.json();
          if (response.ok && data.status === "connected") {
            setIsConnected(true);
            if (data.versions) {
              setUser(data.versions);
            }
          } else {
            handleLogout();
          }
        } catch (apiError) {
          console.error("Error verifying token:", apiError);
          setError("Could not verify Salesforce connection. Check your internet connection.");
        }
      }
    } catch (error) {
      console.error("Error checking connection status:", error);
      setError("Failed to check connection status");
    } finally {
      setIsLoading(false);
    }
  };

  const startOAuth = async () => {
    try {
      setIsLoading(true);
      setError(null);

      const authUrl =
        `${LOGIN_URL}/services/oauth2/authorize?` +
        `response_type=token` +
        `&client_id=${CONSUMER_KEY}` +
        `&redirect_uri=${encodeURIComponent(REDIRECT_URI)}`;

      Office.context.ui.displayDialogAsync(
        authUrl,
        { height: 80, width: 50, promptBeforeOpen: true },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open dialog:", result.error.message);
            setError(`Failed to open login dialog: ${result.error.message}`);
            setIsLoading(false);
          } else {
            const dialogInstance = result.value;
            setDialog(dialogInstance);

            dialogInstance.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
            dialogInstance.addEventHandler(
              Office.EventType.DialogEventReceived,
              processDialogEvent
            );
          }
        }
      );
    } catch (error) {
      console.error("Error starting OAuth flow:", error);
      setError(`Authorization failed: ${error.message}`);
      setIsLoading(false);
    }
  };

  const processMessage = (arg) => {
    try {
      const messageFromDialog = JSON.parse(arg.message);

      if (messageFromDialog.access_token && messageFromDialog.instance_url) {
        if (dialog) {
          try {
            dialog.close();
          } catch (closeError) {
            console.error("Error closing dialog:", closeError);
          }
        }
        storeTokens({
          access_token: messageFromDialog.access_token,
          refresh_token: messageFromDialog.refresh_token || null,
          instance_url: messageFromDialog.instance_url,
        })
          .then(() => {
            setIsConnected(true);
            setDialog(null);
            setIsLoading(false);
          })
          .catch((err) => {
            console.error("Error storing tokens:", err);
            setError(`Failed to store tokens: ${err.message}`);
            setIsLoading(false);
          });
      } else if (messageFromDialog.error) {
        console.error("OAuth error:", messageFromDialog.error);
        setError(`Authentication error: ${messageFromDialog.error}`);

        if (dialog) {
          try {
            dialog.close();
          } catch (closeError) {
            console.error("Error closing dialog:", closeError);
          }
        }
        setDialog(null);
        setIsLoading(false);
      }
    } catch (error) {
      console.error("Error processing message:", error, "Raw message:", arg.message);
      setError(`Error processing authentication response: ${error.message}`);

      if (dialog) {
        try {
          dialog.close();
        } catch (closeError) {
          console.error("Error closing dialog on error:", closeError);
        }
      }
      setDialog(null);
      setIsLoading(false);
    }
  };

  const processDialogEvent = (arg) => {
    setDialog(null);
    setIsLoading(false);
  };

  const handleLogout = () => {
    try {
      Office.context.document.settings.remove("salesforce_access_token");
      Office.context.document.settings.remove("salesforce_refresh_token");
      Office.context.document.settings.remove("salesforce_instance_url");

      Office.context.document.settings.saveAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          setIsConnected(false);
          setUser(null);
        } else {
          setError(`Failed to logout: ${result.error.message}`);
        }
      });
    } catch (err) {
      setError(`Logout failed: ${err.message}`);
    }
  };

  const storeTokens = async (tokenData) => {
    try {
      Office.context.document.settings.set("salesforce_access_token", tokenData.access_token);
      if (tokenData.refresh_token) {
        Office.context.document.settings.set("salesforce_refresh_token", tokenData.refresh_token);
      }
      Office.context.document.settings.set("salesforce_instance_url", tokenData.instance_url);

      await new Promise((resolve, reject) => {
        Office.context.document.settings.saveAsync((result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            resolve();
          } else {
            reject(new Error("Failed to save settings"));
          }
        });
      });
    } catch (error) {
      console.error("Error storing tokens:", error);
      throw error;
    }
  };

  const renderConnectedProfile = () => (
    <div className="salesforce-connected-card">
      <Stack tokens={profileStackTokens}>
        <div className="profile-header">
          <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
            <Text variant="mediumPlus" styles={{ root: { fontWeight: FontWeights.semibold } }}>
              Salesforce Connection
            </Text>
            <div className="status-badge">
              Connected
            </div>
          </Stack>
        </div>
        
        <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }}>
          <Persona
            imageUrl={user?.picture}
            text={user?.name || 'Salesforce User'}
            secondaryText={user?.email}
            tertiaryText={orgType === 'production' ? 'Production Org' : 'Sandbox Org'}
            size={PersonaSize.size72}
            presence={PersonaPresence.online}
            styles={{
              root: { margin: '0 auto' },
              primaryText: { 
                fontWeight: FontWeights.semibold,
                fontSize: '16px',
                color: '#323130'
              },
              secondaryText: { 
                color: '#605e5c',
                fontSize: '14px'
              },
              tertiaryText: {
                color: '#0078d4',
                fontSize: '12px',
                fontWeight: FontWeights.semibold
              }
            }}
          />
        </Stack>

        <Stack horizontal tokens={buttonStackTokens} horizontalAlign="center">
          <DefaultButton
            text="Disconnect"
            onClick={handleLogout}
            iconProps={{ iconName: "PlugDisconnected" }}
            styles={{
              root: {
                minWidth: '120px',
                borderColor: '#d83b01',
                color: '#d83b01'
              },
              rootHovered: {
                borderColor: '#a4262c',
                color: '#a4262c',
                backgroundColor: '#fdf6f6'
              }
            }}
          />
        </Stack>
      </Stack>
    </div>
  );

  const renderConnectButton = () => (
    <div className="salesforce-connect-card">
      <Stack tokens={{ childrenGap: 20 }} horizontalAlign="center">
        <Stack horizontalAlign="center" tokens={{ childrenGap: 12 }}>
          <Text variant="large" styles={{ root: { fontWeight: FontWeights.semibold, textAlign: 'center' } }}>
            Connect to Salesforce
          </Text>
          <Text variant="medium" styles={{ root: { color: '#605e5c', textAlign: 'center' } }}>
            Authenticate with your {orgType === 'production' ? 'Production' : 'Sandbox'} org to get started
          </Text>
        </Stack>
        
        <DefaultButton
          text="Connect to Salesforce"
          onClick={startOAuth}
          iconProps={{ iconName: "PlugConnected" }}
          primary
          styles={{
            root: {
              minWidth: '180px',
              height: '40px',
              fontSize: '14px',
              fontWeight: FontWeights.semibold
            }
          }}
        />
      </Stack>
    </div>
  );

  return (
    <div className="salesforce-auth-container">
      {error && (
        <MessageBar
          messageBarType={MessageBarType.error}
          onDismiss={() => setError(null)}
          dismissButtonAriaLabel="Close"
          className="error-message"
          styles={{ root: { marginBottom: '20px', maxWidth: '450px', margin: '0 auto 20px auto' } }}
        >
          {error}
        </MessageBar>
      )}

      {isLoading ? (
        <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { minHeight: '200px' } }}>
          <Spinner size={SpinnerSize.large} label="Connecting to Salesforce..." />
        </Stack>
      ) : (
        <Stack horizontalAlign="center" verticalAlign="center">
          {isConnected ? renderConnectedProfile() : renderConnectButton()}
        </Stack>
      )}
    </div>
  );
};

export default SalesforceAuth;