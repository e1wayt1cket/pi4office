import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  target: Type.Optional(
    Type.String({
      enum: ["selection", "paragraphs"],
      description:
        "What to delete: 'selection' deletes the currently selected text, " +
        "'paragraphs' deletes a range of paragraphs by index. Defaults to 'selection'.",
    }),
  ),
  paragraph_start: Type.Optional(
    Type.Number({
      description: "First paragraph index to delete (0-based, inclusive). Required when target is 'paragraphs'.",
    }),
  ),
  paragraph_count: Type.Optional(
    Type.Number({
      description: "Number of paragraphs to delete. Defaults to 1.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createDeleteTextTool(): AgentTool<typeof schema> {
  return {
    name: "delete_text",
    label: t("tools.deleteText"),
    description:
      "Delete text from the document. Can delete the current selection or a range " +
      "of paragraphs by index. Use get_document_outline or read_document first to " +
      "identify the paragraph indices to delete.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await deleteText(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error deleting text: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function deleteText(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const target = params.target ?? "selection";

    if (target === "selection") {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text || "";
      if (text.length === 0) {
        return "Selection is empty — nothing to delete.";
      }

      const preview = text.length > 80 ? text.slice(0, 80) + "..." : text;
      selection.delete();
      await context.sync();
      return `Deleted selection: "${preview}"`;
    }

    // Delete a range of paragraphs
    const body = context.document.body;
    body.paragraphs.load("items");
    await context.sync();

    const paragraphs = body.paragraphs.items;
    const totalParagraphs = paragraphs.length;
    if (totalParagraphs === 0) {
      return "Document has no paragraphs to delete.";
    }

    const startIdx = params.paragraph_start;
    if (startIdx === undefined) {
      return "paragraph_start is required when target is 'paragraphs'.";
    }

    if (startIdx < 0 || startIdx >= totalParagraphs) {
      return `Paragraph index ${startIdx} out of range (document has ${totalParagraphs} paragraphs, 0-based).`;
    }

    const count = params.paragraph_count ?? 1;
    const endIdx = Math.min(startIdx + count - 1, totalParagraphs - 1);
    const deletedCount = endIdx - startIdx + 1;

    const firstPara = paragraphs[startIdx];
    const lastPara = paragraphs[endIdx];
    if (!firstPara || !lastPara) {
      return "Could not access target paragraphs.";
    }

    const firstRange = firstPara.getRange("Start");
    const lastRange = lastPara.getRange("End");

    const deleteRange = firstRange.expandTo(lastRange);
    deleteRange.delete();

    await context.sync();

    const rangeLabel = deletedCount === 1
      ? `paragraph ${startIdx}`
      : `paragraphs ${startIdx}-${endIdx}`;
    return `Deleted ${deletedCount} paragraph(s) (${rangeLabel}).`;
  });
}
