import * as React from "react";
import {
  FluentProvider,
  webLightTheme,
  Button,
  Text,
  Title3,
  Card,
  CardHeader,
} from "@fluentui/react-components";
import { useState, useEffect } from "react";
import { app, authentication } from "@microsoft/teams-js";
import jwtDecode from "jwt-decode";
import { useTeamsContext } from "../hooks/useTeamsContext";

/**
 * Implementation of the Archive content page
 */
export const ArchiveTab = () => {
  const { inTeams, theme, context } = useTeamsContext();
  const [entityId, setEntityId] = useState<string | undefined>();
  const [name, setName] = useState<string>();
  const [error, setError] = useState<string>();
  const [chats, setChats] = useState<any[]>([]);

  useEffect(() => {
    const run = async () => {
      try {
        const ssoToken = await authentication.getAuthToken({
          resources: [process.env.TAB_APP_URI as string],
          silent: false,
        });
        const response = await fetch("/api/graph/chats", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ token: ssoToken }),
        });
        if (!response.ok) throw new Error("Failed to fetch chats");
        const chatsData = await response.json();
        setChats(chatsData.value || chatsData);
        app.notifySuccess();
      } catch (err: any) {
        setError(err.message || "Error fetching chats");
        app.notifyFailure({
          reason: app.FailedReason.AuthFailed,
          message: err.message || "Error fetching chats",
        });
      }
    };
    run();
  }, [inTeams, context]);

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <FluentProvider theme={webLightTheme}>
      <Card>
        <CardHeader header={<Title3>This is your tab</Title3>} />
        <div>
          <Text>Hello {name}</Text>
          {error && (
            <div>
              <Text>An SSO error occurred {error}</Text>
            </div>
          )}
          <div>
            <Button onClick={() => alert("It worked!")}>A sample button</Button>
          </div>
        </div>
        <div>
          <Text size={200}>(C) Copyright Christopher Øyen</Text>
        </div>
      </Card>
    </FluentProvider>
  );
};
