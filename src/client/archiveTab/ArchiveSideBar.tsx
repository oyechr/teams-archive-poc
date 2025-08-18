import * as React from "react";
import {
  Drawer,
  TabList,
  Tab,
  Input,
  Textarea,
  Button,
  Label,
} from "@fluentui/react-components";
import { Dismiss24Regular } from "@fluentui/react-icons";

export type Metadata = {
  caseNumber: string;
  meetingContext: string;
  participants: string;
  meetingDate: string;
  senderRecipient: string;
};

type ArchiveSidebarProps = {
  open: boolean;
  onClose: () => void;
  chatId: string;
  chatHistory: any[];
  metadata: Metadata;
  onMetadataChange: (meta: Metadata) => void;
};

export const ArchiveSidebar: React.FC<ArchiveSidebarProps> = ({
  open,
  onClose,
  chatId,
  chatHistory,
  metadata,
  onMetadataChange,
}) => {
  const [tab, setTab] = React.useState("metadata");

  const handleChange = (field: keyof Metadata, value: string) => {
    onMetadataChange({ ...metadata, [field]: value });
  };

  return (
    <Drawer
      open={open}
      onOpenChange={(_, { open }) => !open && onClose()}
      position="end"
      style={{ width: "600px" }}
    >
      <div
        style={{ display: "flex", justifyContent: "flex-end", padding: "8px" }}
      >
        <Button
          appearance="transparent"
          icon={<Dismiss24Regular />}
          onClick={onClose}
        />
      </div>
      <TabList
        selectedValue={tab}
        onTabSelect={(_, data) => setTab(data.value as string)}
      >
        <Tab value="metadata">Custom Metadata</Tab>
        <Tab value="preview">Content Preview</Tab>
      </TabList>
      {tab === "metadata" && (
        <div
          style={{
            padding: "16px",
            display: "flex",
            flexDirection: "column",
            gap: "16px",
          }}
        >
          <Label>Case Number</Label>
          <Input
            value={metadata.caseNumber}
            onChange={(_, v) => handleChange("caseNumber", v.value)}
            placeholder="Case Number"
          />
          <Label>Meeting Context</Label>
          <Textarea
            value={metadata.meetingContext}
            onChange={(_, v) => handleChange("meetingContext", v.value)}
            placeholder="Meeting Context"
          />
          <Label>Participants</Label>
          <Input
            value={metadata.participants}
            onChange={(_, v) => handleChange("participants", v.value)}
            placeholder="Participants"
          />
          <Label>Meeting Date</Label>
          <Input
            type="date"
            value={metadata.meetingDate}
            onChange={(_, v) => handleChange("meetingDate", v.value)}
            placeholder="Meeting Date"
          />
          <Label>Sender/Recipient</Label>
          <Input
            value={metadata.senderRecipient}
            onChange={(_, v) => handleChange("senderRecipient", v.value)}
            placeholder="Sender/Recipient"
          />
          <Button appearance="primary" style={{ alignSelf: "flex-end" }} onClick={onClose}>
            Save & Archive
          </Button>
        </div>
      )}
      {tab === "preview" && (
        <pre
          style={{
            whiteSpace: "pre-wrap",
            padding: "16px",
            maxHeight: "60vh",
            overflowY: "scroll",
            background: "#f3f3f3",
            borderRadius: "4px",
            fontSize: "14px",
            width: "100%",
            boxSizing: "border-box",
            scrollbarWidth: "thin", 
          }}
        >
          {JSON.stringify(
            {
              chatId,
              metadata,
              chatHistory,
            },
            null,
            2
          )}
        </pre>
      )}
    </Drawer>
  );
};
