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

  useEffect(() => {
    if (inTeams && context) {
      authentication
        .getAuthToken({
          resources: [process.env.TAB_APP_URI as string],
          silent: false,
        } as authentication.AuthTokenRequestParameters)
        .then((token) => {
          const decoded: { [key: string]: any } = jwtDecode(token) as {
            [key: string]: any;
          };
          setName(decoded!.name);
          app.notifySuccess();
        })
        .catch((message) => {
          setError(message);
          app.notifyFailure({
            reason: app.FailedReason.AuthFailed,
            message,
          });
        });
    } else if (!inTeams) {
      setEntityId("Not in Microsoft Teams");
    }
  }, [inTeams, context]);

  useEffect(() => {
    if (context) {
      setEntityId(context.page.id);
    }
  }, [context]);

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
