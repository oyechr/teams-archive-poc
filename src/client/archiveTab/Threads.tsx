import * as React from "react";
import {
  Card,
  CardHeader,
  Title3,
  Text,
  Link,
  Switch,
} from "@fluentui/react-components";
import { ArchiveSidebar, Metadata } from "./ArchiveSideBar";

type ThreadsProps = {
  threads: any[];
  theme?: string;
};

export const Threads: React.FC<ThreadsProps> = ({ threads, theme = "default" }) => {
  const [archivedThreads, setArchivedThreads] = React.useState<string[]>([]);
  const [sidebarOpen, setSidebarOpen] = React.useState(false);
  const [selectedThreadIdx, setSelectedThreadIdx] = React.useState<
    number | null
  >(null);
  const [metadata, setMetadata] = React.useState<Record<number, Metadata>>({});
  if (!threads || threads.length === 0) {
    return (
      <Card>
        <CardHeader header={<Title3>Message Threads</Title3>} />
        <Text>No threads flagged yet.</Text>
      </Card>
    );
  }

  const handleArchiveToggle = (idx: number) => {
    setArchivedThreads((prev) =>
      prev.includes(idx.toString())
        ? prev.filter((id) => id !== idx.toString())
        : [...prev, idx.toString()]
    );
    setSelectedThreadIdx(idx);
    setSidebarOpen(true);
  };

  return (
    <div>
      <CardHeader
        header={<Title3>Message Threads Flagged for Archive</Title3>}
      />
      {threads.map((thread, idx) => {
        const content = thread?.body?.content || "No content";
        const sender = thread?.from?.user?.displayName || "Unknown sender";
        const link = thread?.linkToMessage;
        const date = thread?.createdDateTime
          ? new Date(thread.createdDateTime).toLocaleString()
          : "Unknown date";
        return (
          <Card
            key={idx}
            style={{
              display: "flex",
              alignItems: "flex-start",
              justifyContent: "space-between",
              marginBottom: 16,
              padding: 16,
            }}
          >
            {/* Left column: thread info */}
            <div style={{ flex: 1, marginRight: 24 }}>
              <Text>
                <strong>{sender}</strong>{" "}
                <span style={{ color: "#888" }}>{date}</span>
                <br />
                <span dangerouslySetInnerHTML={{ __html: content }} />
                <br />
                {link && (
                  <Link href={link} target="_blank">
                    View in Teams
                  </Link>
                )}
              </Text>
            </div>
            {/* Right column: archive toggle and sidebar */}
            <div
              style={{
                display: "flex",
                flexDirection: "column",
                alignItems: "flex-end",
              }}
            >
              <div
                style={{
                  display: "flex",
                  alignItems: "center",
                  marginBottom: 8,
                }}
              >
                <Text size={300} style={{ marginRight: 8 }}>
                  Archive
                </Text>
                <Switch
                  checked={archivedThreads.includes(idx.toString())}
                  onChange={() => handleArchiveToggle(idx)}
                  aria-label="Archive this thread"
                />
              </div>
              {selectedThreadIdx === idx && sidebarOpen && (
                <ArchiveSidebar
                  open={sidebarOpen}
                  onClose={() => {
                    setSidebarOpen(false);
                    setSelectedThreadIdx(null);
                    setArchivedThreads(prev => prev.filter(id => id !== idx.toString()));
                  }}
                  chatId={thread.id || idx.toString()}
                  chatHistory={thread|| []}
                  metadata={
                    metadata[idx] || {
                      caseNumber: "",
                      meetingContext: "",
                      participants: "",
                      meetingDate: "",
                      senderRecipient: "",
                    }
                  }
                  onMetadataChange={(meta) =>
                    setMetadata((prev) => ({ ...prev, [idx]: meta }))
                  }
                  theme={theme}
                />
              )}
            </div>
          </Card>
        );
      })}
    </div>
  );
};
