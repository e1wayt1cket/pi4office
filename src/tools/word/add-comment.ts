import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  text: Type.String({
    description: "The comment text to add.",
  }),
  selection_only: Type.Optional(Type.Boolean({
    default: true,
    description: "When true (default), comment is added to the current selection. " +
      "When false, comment is added at the cursor position.",
  })),
});

type Params = Static<typeof schema>;

export function createAddCommentTool(): AgentTool<typeof schema> {
  return {
    name: "add_comment",
    label: t("tools.addComment"),
    description:
      "Add a comment to the selected text in the document. " +
      "The selection determines which text the comment is attached to.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await addComment(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error adding comment: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function addComment(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const selection = context.document.getSelection();
    selection.insertComment(params.text);

    await context.sync();

    const preview = params.text.length > 80
      ? params.text.slice(0, 80) + "..."
      : params.text;
    return `Comment added: "${preview}"`;
  });
}
