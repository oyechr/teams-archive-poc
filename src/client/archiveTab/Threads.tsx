import * as React from "react";
import { Card, CardHeader, Title3, Text } from "@fluentui/react-components";

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
        {threads.map((thread, idx) => (
          <Card key={idx} style={{ marginBottom: 16 }}>
            <Text>
              <pre style={{ whiteSpace: "pre-wrap", wordBreak: "break-word" }}>
                {typeof thread === "string"
                  ? thread
                  : JSON.stringify(thread, null, 2)}
              </pre>
            </Text>
          </Card>
        ))}
      </div>
    </Card>
  );
};