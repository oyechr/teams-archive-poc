import { ConfidentialClientApplication } from "@azure/msal-node";
import axios from "axios";

const msalConfig = {
  auth: {
    clientId: process.env.APPLICATION_ID!,
    authority: `https://login.microsoftonline.com/${process.env.AZURE_AD_TENANT_ID}`,
    clientSecret: process.env.AZURE_AD_CLIENT_SECRET!,
  },
};

const msalClient = new ConfidentialClientApplication(msalConfig);

export async function getGraphTokenOnBehalfOf(userToken: string) {
  const oboRequest = {
    oboAssertion: userToken,
    scopes: ["https://graph.microsoft.com/.default"],
  };
  const response = await msalClient.acquireTokenOnBehalfOf(oboRequest);
  return response?.accessToken;
}

export async function fetchUserChatsOBO(userToken: string) {
  const graphToken = await getGraphTokenOnBehalfOf(userToken);
  const result = await axios.get("https://graph.microsoft.com/v1.0/me/chats", {
    headers: { Authorization: `Bearer ${graphToken}` },
  });
  return result.data;
}

export async function fetchChatMessagesOBO(userToken: string, chatId: string) {
  const graphToken = await getGraphTokenOnBehalfOf(userToken);
  const result = await axios.get(
    `https://graph.microsoft.com/v1.0/chats/${chatId}/messages`,
    {
      headers: { Authorization: `Bearer ${graphToken}` },
    }
  );
  return result.data;
}

export async function fetchChatMembersOBO(userToken: string, chatId: string) {
  const graphToken = await getGraphTokenOnBehalfOf(userToken);
  const result = await axios.get(
    `https://graph.microsoft.com/v1.0/chats/${chatId}/members`,
    {
      headers: { Authorization: `Bearer ${graphToken}` },
    }
  );
  return result.data;
}