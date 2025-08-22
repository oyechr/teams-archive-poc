import express from "express";
import * as path from "path";
import morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import debug from "debug";
import compression from "compression";
import {
  CloudAdapter,
  ConfigurationServiceClientCredentialFactory,
  createBotFrameworkAuthenticationFromConfiguration,
} from "botbuilder";
import { ArchiveMessagingExtensionBot } from "./ArchiveMessagingExtension";

// Initialize debug logging module
const log = debug("msteams");
const flaggedThreads: any[] = [];
log("Initializing Microsoft Teams Express hosted App...");

require("dotenv").config();
const credentialsFactory = new ConfigurationServiceClientCredentialFactory({
  MicrosoftAppId: process.env.MicrosoftAppId,
  MicrosoftAppPassword: process.env.MicrosoftAppPassword,
  MicrosoftAppTenantId: process.env.AZURE_AD_TENANT_ID,
});
const config = {
  get: <T = unknown>(path?: string[]): T | undefined => {
    if (!path || path.length === 0) return undefined;
    return process.env[path[0]] as unknown as T;
  },
  set: (path: string[], value: string) => {
    process.env[path[0]] = value;
  },
};
const botFrameworkAuthentication =
  createBotFrameworkAuthenticationFromConfiguration(config, credentialsFactory);

const adapter = new CloudAdapter(botFrameworkAuthentication);

adapter.onTurnError = async (context, error) => {
  console.error(`\n [onTurnError] unhandled error: ${error}`);
  await context.sendTraceActivity(
    "OnTurnError Trace",
    `${error}`,
    "https://www.botframework.com/schemas/error",
    "TurnError"
  );
};

const bot = new ArchiveMessagingExtensionBot();

// The import of components has to be done AFTER the dotenv config
import * as allComponents from "./TeamsAppsComponents";
import {
  fetchUserChatsOBO,
  fetchChatMessagesOBO,
  fetchChatMembersOBO,
  fetchUserPresenceOBO,
} from "./graphObo";

// Create the Express webserver
const app = express();
const port = process.env.port || process.env.PORT || 3007;

// Inject the raw request body onto the request object
app.use(
  express.json({
    verify: (req, res, buf: Buffer, encoding: string): void => {
      (req as any).rawBody = buf.toString();
    },
  })
);
app.use(express.urlencoded({ extended: true }));

// Express configuration
app.set("views", path.join(__dirname, "/"));

// Add simple logging
app.use(morgan("tiny"));

// Add compression - uncomment to remove compression
app.use(compression());

// Add /scripts and /assets as static folders
app.use("/scripts", express.static(path.join(__dirname, "web/scripts")));
app.use("/assets", express.static(path.join(__dirname, "web/assets")));

// routing for bots, connectors and incoming web hooks - based on the decorators
// For more information see: https://www.npmjs.com/package/express-msteams-host
app.use(MsTeamsApiRouter(allComponents));

// routing for pages for tabs and connector configuration
// For more information see: https://www.npmjs.com/package/express-msteams-host
app.use(
  MsTeamsPageRouter({
    root: path.join(__dirname, "web/"),
    components: allComponents,
  })
);

// Endpoint for Teams bot activities
app.post("/api/messages", async (req, res) => {
  console.log("Received POST /api/messages");
  await adapter.process(req, res, async (context) => {
    await bot.run(context);
  });
});

// Endpoint to mark a thread for archive (called by bot)
app.post("/api/markForArchive", async (req, res) => {
  const { messagePayload, replies } = req.body; 
  flaggedThreads.push({
    chatId: messagePayload.id,
    metadata: messagePayload.metadata || {},
    chatHistory: [messagePayload, ...(replies || [])], // Store root + replies
  });
  res.json({ status: "success" });
});

// Endpoint for ArchiveTab to fetch archived threads
app.get("/api/archivedThreads", (req, res) => {
  res.json(flaggedThreads);
});

// OBO proxy route
app.post("/api/graph/chats", async (req, res) => {
  try {
    const userToken = req.body.token;
    const chats = await fetchUserChatsOBO(userToken);
    res.json(chats);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});
app.post("/api/graph/chats/:chatId/members", async (req, res) => {
  try {
    const userToken = req.body.token;
    const chatId = req.params.chatId;
    const members = await fetchChatMembersOBO(userToken, chatId);
    res.json(members);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});
app.post("/api/graph/chats/:chatId/messages", async (req, res) => {
  try {
    const userToken = req.body.token;
    const chatId = req.params.chatId;
    const messages = await fetchChatMessagesOBO(userToken, chatId);
    res.json(messages);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});
app.post("/api/graph/users/:userId/presence", async (req, res) => {
  try {
    const userToken = req.body.token;
    const userId = req.params.userId;
    const presence = await fetchUserPresenceOBO(userToken, userId);
    res.json(presence);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});
// app.post("/api/markForArchive", async (req, res) => {
//   const { messagePayload } = req.body;
//   // Store the thread/message in database or in-memory store?
//   // For poc, just log and return success
//   console.log("Marked for archive:", messagePayload);
//   // TODO: Save to persistent storage
//   res.json({ status: "success" });
// });

app.use(
  "/",
  express.static(path.join(__dirname, "web/"), {
    index: "index.html",
  })
);

// Set the port
app.set("port", port);

// Start the webserver
app.listen(port, () => {
  log(`Server running on ${port}`);
});
