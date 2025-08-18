import { TeamsActivityHandler, TurnContext, CardFactory } from "botbuilder";

export class ArchiveMessagingExtensionBot extends TeamsActivityHandler {
  async handleTeamsMessagingExtensionSubmitAction(context: TurnContext, action: any) {
    const messagePayload = action.messagePayload;
    await fetch("https://eagerly-expert-jaybird.ngrok-free.app/api/markForArchive", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ messagePayload }),
    });

    const adaptiveCard = {
      type: "AdaptiveCard",
      version: "1.4",
      body: [
        {
          type: "TextBlock",
          text: "Thread marked for archive",
          weight: "Bolder",
          size: "Medium"
        },
        {
          type: "TextBlock",
          text: JSON.stringify(messagePayload, null, 2),
          wrap: true,
          fontType: "Monospace"
        }
      ]
    };

    return {
      composeExtension: {
        type: "result" as const,
        attachmentLayout: "list" as const,
        attachments: [CardFactory.adaptiveCard(adaptiveCard)]
      }
    };
  }
}