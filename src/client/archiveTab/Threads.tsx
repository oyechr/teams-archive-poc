import * as React from "react";
import { Card, CardHeader, Title3, Text, Link } from "@fluentui/react-components";

type ThreadsProps = {
  threads: any[];
};

export const Threads: React.FC<ThreadsProps> = ({ threads }) => {
  if (!threads || threads.length === 0) {
    return (
      <Card>
        <CardHeader header={<Title3>Message Threads</Title3>} />
        <Text>No threads flagged yet.</Text>
      </Card>
    );
  }

  return (
    <Card>
      <CardHeader header={<Title3>Message Threads Flagged for Archive</Title3>} />
      <div style={{ marginTop: 12 }}>
        {threads.map((thread, idx) => {
          const content = thread?.body?.content || "No content";
          const sender = thread?.from?.user?.displayName || "Unknown sender";
          const link = thread?.linkToMessage;
          const date = thread?.createdDateTime
            ? new Date(thread.createdDateTime).toLocaleString()
            : "Unknown date";
          return (
            <Card key={idx} style={{ marginBottom: 16 }}>
              <Text>
                <strong>{sender}</strong> <span style={{ color: "#888" }}>{date}</span>
                <br />
                <span dangerouslySetInnerHTML={{ __html: content }} />
                <br />
                {link && (
                  <Link href={link} target="_blank">
                    View in Teams
                  </Link>
                )}
              </Text>
            </Card>
          );
        })}
      </div>
    </Card>
  );
};