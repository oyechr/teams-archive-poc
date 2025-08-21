import { TeamsActivityHandler, TurnContext } from "botbuilder";

export class ArchiveMessagingExtensionBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: any) {
    try {
      const messagePayload = action.messagePayload;
      await fetch("https://eagerly-expert-jaybird.ngrok-free.app/api/markForArchive", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ messagePayload }),
      });

      const selectedMessage = messagePayload?.body?.content || "Message not found.";
      return {
        composeExtension: {
          type: "result" as const,
          attachmentLayout: "list" as const,
          attachments: [
            {
              contentType: "application/vnd.microsoft.card.hero",
              content: {
                title: "Selected Message",
                text: selectedMessage
              }
            }
          ]
        }
      };
    } catch (err) {
      console.error("Error in handleTeamsMessagingExtensionSubmitAction:", err);
      return {
        composeExtension: {
          type: "message" as const,
          text: "An error occurred while archiving the message."
        }
      };
    }
  }
}