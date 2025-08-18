import express from "express";
import * as path from "path";
import morgan from "morgan";
import { MsTeamsApiRouter, MsTeamsPageRouter } from "express-msteams-host";
import debug from "debug";
import compression from "compression";
import axios from "axios";

// Initialize debug logging module
const log = debug("msteams");

log("Initializing Microsoft Teams Express hosted App...");

// Initialize dotenv, to use .env file settings if existing
require("dotenv").config();

// The import of components has to be done AFTER the dotenv config
import * as allComponents from "./TeamsAppsComponents";
import {
  fetchUserChatsOBO,
  fetchChatMessagesOBO,
  fetchChatMembersOBO,
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
    // You need to implement fetchChatMembersOBO
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
    // You need to implement fetchChatMessagesOBO
    const messages = await fetchChatMessagesOBO(userToken, chatId);
    res.json(messages);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get("/api/graph/image", async (req, res) => {
  const url = req.query.url as string;
  const token = req.query.token as string;
  if (!url || !token) {
    return res.status(400).json({ error: "Missing url or token" });
  }
  try {
    const response = await axios.get(url, {
      headers: { Authorization: `Bearer ${token}` },
      responseType: "arraybuffer",
    });
    res.set("Content-Type", response.headers["content-type"]);
    res.send(response.data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Set default web page
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
