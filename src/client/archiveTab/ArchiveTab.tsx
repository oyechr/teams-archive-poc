import * as React from "react";
import DOMPurify from "dompurify";
import {
  FluentProvider,
  teamsLightTheme,
  Text,
  Title3,
  Card,
  CardHeader,
  DataGrid,
  Switch,
  DataGridBody,
  DataGridHeader,
  DataGridHeaderCell,
  createTableColumn,
  DataGridRow,
  DataGridCell,
  TableCellLayout,
  Avatar,
  teamsDarkTheme,
  teamsHighContrastTheme,
} from "@fluentui/react-components";
import { PresenceBadgeStatus } from "@fluentui/react-components";
import { ArchiveSidebar, Metadata } from "./ArchiveSideBar";
import { Threads } from "./Threads";

type UserCell = {
  label: string;
  status: PresenceBadgeStatus;
};
import { useState, useEffect } from "react";
import { app, authentication } from "@microsoft/teams-js";
import { useTeamsContext } from "../hooks/useTeamsContext";

/**
 * Implementation of the Archive content page
 */
function processMessageHtml(html: string) {
  let processed = html.replace(
    /<emoji[^>]*alt="([^"]+)"[^>]*>.*?<\/emoji>/g,
    "$1"
  );
  processed = processed.replace(/\s*<img[^>]*>\s*/g, " [image] ");
  processed = processed.replace(/\s+/g, " ").trim();
  return processed;
}
function mapPresenceToBadgeStatus(presence: string): PresenceBadgeStatus {
  switch (presence) {
    case "Available":
      return "available";
    case "Busy":
      return "busy";
    case "Away":
      return "away";
    case "DoNotDisturb":
      return "do-not-disturb";
    case "Offline":
      return "offline";
    default:
      return "unknown";
  }
}
type ChatRow = {
  rowId: string;
  author: UserCell;
  lastMessage: string;
  archive: string;
};

const themeMap = {
  default: teamsLightTheme,
  light: teamsLightTheme,
  dark: teamsDarkTheme,
  contrast: teamsHighContrastTheme,
};
export const ArchiveTab = () => {
  const { inTeams, theme, context } = useTeamsContext();
  const [error, setError] = useState<string>();
  const [chats, setChats] = useState<any[]>([]);
  const [archivedChats, setArchivedChats] = useState<string[]>([]);
  const [chatMessages, setChatMessages] = useState<Record<string, any[]>>({});
  const [chatDetails, setChatDetails] = useState<
    Record<
      string,
      { name: string; lastMessage: string; presence: PresenceBadgeStatus }
    >
  >({});
  const [ssoToken, setSsoToken] = useState<string>("");
  const [sidebarOpen, setSidebarOpen] = useState(false);
  const [selectedChatId, setSelectedChatId] = useState<string | null>(null);
  const [metadata, setMetadata] = useState<Record<string, Metadata>>({});

  useEffect(() => {
    const fetchChats = async () => {
      try {
        const token = await authentication.getAuthToken({
          resources: [process.env.TAB_APP_URI as string],
          silent: false,
        });
        setSsoToken(token);
        const response = await fetch("/api/graph/chats", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ token }),
        });
        if (!response.ok) throw new Error("Failed to fetch chats");
        const chatsData = await response.json();
        setChats(chatsData.value || chatsData);
        //console.log("Fetched chats:", chatsData.value || chatsData);
        app.notifySuccess();
      } catch (err: any) {
        setError(err.message || "Error fetching chats");
        app.notifyFailure({
          reason: app.FailedReason.AuthFailed,
          message: err.message || "Error fetching chats",
        });
      }
    };
    fetchChats();
  }, [inTeams, context]);

  const userId = context?.user?.id;
  useEffect(() => {
    if (chats.length === 0 || !context || !userId) return;
    const fetchDetails = async () => {
      const details: Record<
        string,
        { name: string; lastMessage: string; presence: PresenceBadgeStatus }
      > = {};

      await Promise.all(
        chats.map(async (chat) => {
          // Fetch members
          const membersRes = await fetch(
            `/api/graph/chats/${chat.id}/members`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ token: ssoToken }),
            }
          );
          let name = "";
          let presence: PresenceBadgeStatus = "unknown";
          if (membersRes.ok) {
            const membersData = await membersRes.json();

            if (
              chat.chatType === "oneOnOne" &&
              Array.isArray(membersData.value)
            ) {
              const other = membersData.value.find(
                (m: any) => m.userId !== context.user?.id
              );
              name = other?.displayName || other?.email || "Unknown";
              if (other?.userId) {
                const presenceRes = await fetch(
                  `/api/graph/users/${other.userId}/presence`,
                  {
                    method: "POST",
                    headers: { "Content-Type": "application/json" },
                    body: JSON.stringify({ token: ssoToken }),
                  }
                );
                if (presenceRes.ok) {
                  const presenceData = await presenceRes.json();
                  presence = mapPresenceToBadgeStatus(
                    presenceData.availability
                  );
                }
              }
            } else if (Array.isArray(membersData.value)) {
              name = membersData.value
                .map((m: any) => m.displayName)
                .join(", ");
            }
          }
          const messagesRes = await fetch(
            `/api/graph/chats/${chat.id}/messages`,
            {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({ token: ssoToken }),
            }
          );
          let lastMessage = "";
          if (messagesRes.ok) {
            const messagesData = await messagesRes.json();
            setChatMessages((prev) => ({
              ...prev,
              [chat.id]: messagesData.value || [],
            }));
            if (
              Array.isArray(messagesData.value) &&
              messagesData.value.length > 0
            ) {
              const validMessages = messagesData.value.filter(
                (msg: any) =>
                  msg.messageType === "message" &&
                  msg.body?.content &&
                  msg.body.content !== "<systemEventMessage/>"
              );
              if (validMessages.length > 0) {
                const latestMsg = validMessages.reduce(
                  (latest: any, msg: any) => {
                    if (!latest) return msg;
                    return new Date(msg.createdDateTime) >
                      new Date(latest.createdDateTime)
                      ? msg
                      : latest;
                  },
                  null
                );
                let rawContent = latestMsg.body.content;
                lastMessage = rawContent.replace(/<img[^>]*>/g, "[image]");
              } else {
                lastMessage = "No user messages";
              }
              //console.log("Chat ID:", chat.id, "Last Message:", lastMessage);
            }
          }

          details[chat.id] = { name, lastMessage, presence };
        })
      );

      setChatDetails(details);
    };
    fetchDetails();
  }, [chats, context, userId, ssoToken]);

  const handleArchiveToggle = (chatId: string) => {
    setArchivedChats((prev) =>
      prev.includes(chatId)
        ? prev.filter((id) => id !== chatId)
        : [...prev, chatId]
    );
    setSelectedChatId(chatId);
    setSidebarOpen(true);
  };

  const columns = [
    createTableColumn<ChatRow>({
      columnId: "author",
      compare: (a, b) => a.author.label.localeCompare(b.author.label),
      renderHeaderCell: () => "Chat",
      renderCell: (item) => (
        <TableCellLayout
          media={
            <Avatar
              aria-label={item.author.label}
              name={item.author.label}
              badge={{ status: item.author.status }}
            />
          }
        >
          {item.author.label}
        </TableCellLayout>
      ),
    }),
    createTableColumn<ChatRow>({
      columnId: "lastMessage",
      compare: (a, b) => a.lastMessage.localeCompare(b.lastMessage),
      renderHeaderCell: () => "Latest Message",
      renderCell: (item) => (
        <span
          style={{
            display: "block",
            maxWidth: 300,
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "pre-line",
            wordBreak: "break-word",
          }}
          dangerouslySetInnerHTML={{
            __html: DOMPurify.sanitize(processMessageHtml(item.lastMessage)),
          }}
        />
      ),
    }),
    createTableColumn<ChatRow>({
      columnId: "archive",
      compare: (a, b) => a.archive.localeCompare(b.archive),
      renderHeaderCell: () => "Archive",
      renderCell: (item) => (
        <Switch
          checked={archivedChats.includes(item.archive)}
          onChange={() => handleArchiveToggle(item.archive)}
        />
      ),
    }),
  ];

  const rows = chats.map((chat) => {
    let label = chatDetails[chat.id]?.name || "Unknown";
    let status: PresenceBadgeStatus =
      chatDetails[chat.id]?.presence || "unknown";
    return {
      rowId: chat.id,
      author: { label, status },
      lastMessage: chatDetails[chat.id]?.lastMessage || "No messages",
      archive: chat.id,
    };
  });

  /**
   * The render() method to create the UI of the tab
   */
  return (
    <FluentProvider theme={themeMap[theme] || teamsLightTheme}>
      <Card>
        <CardHeader header={<Title3>Chats</Title3>} />
        <div>
          {error && (
            <div>
              <Text>An SSO error occurred {error}</Text>
            </div>
          )}
          <div style={{ marginBottom: 12 }}>
            <Text>
              View and select which chats you want to archive.
            </Text>
          </div>
        </div>
        <div style={{ minWidth: 600 }}>
          <DataGrid
            items={rows}
            columns={columns}
            getRowId={(row) => row.rowId}
            style={{ minWidth: "600px" }}
            aria-label="Chats Table"
          >
            <DataGridHeader>
              <DataGridRow>
                {({ renderHeaderCell }) => (
                  <DataGridHeaderCell>{renderHeaderCell()}</DataGridHeaderCell>
                )}
              </DataGridRow>
            </DataGridHeader>
            <DataGridBody>
              {({ item, rowId }) => (
                <DataGridRow key={rowId}>
                  {({ renderCell }) => (
                    <DataGridCell>{renderCell(item)}</DataGridCell>
                  )}
                </DataGridRow>
              )}
            </DataGridBody>
          </DataGrid>
          {selectedChatId && (
            <ArchiveSidebar
              open={sidebarOpen}
              onClose={() => {
                setSidebarOpen(false);
                if (selectedChatId && archivedChats.includes(selectedChatId)) {
                  setArchivedChats((prev) =>
                    prev.filter((id) => id !== selectedChatId)
                  );
                }
                setSelectedChatId(null);
              }}
              chatId={selectedChatId}
              chatHistory={chatMessages[selectedChatId] || []}
              metadata={
                metadata[selectedChatId] || {
                  caseNumber: "",
                  meetingContext: "",
                  participants: "",
                  meetingDate: "",
                  senderRecipient: "",
                }
              }
              onMetadataChange={(meta) =>
                setMetadata((prev) => ({ ...prev, [selectedChatId]: meta }))
              }
              theme={theme}
            />
          )}
        </div>
      </Card>
      <Card style={{ flex: 1, minWidth: 400 }}>
        <Threads threads={archivedChats} />
      </Card>
    </FluentProvider>
  );
};
