import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  text: Type.String({
    description: "The text to insert into the document.",
  }),
  position: Type.Optional(
    Type.String({
      enum: ["Replace", "Start", "End", "Before", "After"],
      description: "Where to insert relative to the current selection. " +
        "Replace (default): replace the selection. " +
        "Start: insert at start of selection. " +
        "End: insert at end of selection. " +
        "Before: insert before selection. " +
        "After: insert after selection.",
    }),
  ),
  paragraph_index: Type.Optional(
    Type.Number({
      description:
        "0-based paragraph index. When set, text is inserted relative to this paragraph " +
        "instead of the current selection. Use with 'Before' or 'After' position. " +
        "Use get_document_outline or read_document to find the right index.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createInsertTextTool(): AgentTool<typeof schema> {
  return {
    name: "insert_text",
    label: t("tools.insertText"),
    description:
      "Insert text into the document at the current cursor position or selection. " +
      "By default replaces the selection. Use the position parameter to control " +
      "where text is inserted relative to the selection.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await insertText(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error inserting text: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

type ParaInsertLoc = "Before" | "After";

function toParaInsertLoc(pos: string): ParaInsertLoc {
  return pos === "Before" ? "Before" : "After";
}

async function insertText(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const position = params.position ?? "Replace";

    if (params.paragraph_index !== undefined) {
      const body = context.document.body;
      body.paragraphs.load("items");
      await context.sync();

      const idx = params.paragraph_index;
      const paragraphs = body.paragraphs.items;
      if (idx < 0 || idx >= paragraphs.length) {
        throw new Error(
          `Paragraph index ${idx} out of range (document has ${paragraphs.length} paragraphs, 0-based).`,
        );
      }

      const targetParagraph = paragraphs[idx];
      if (!targetParagraph) {
        throw new Error(`Paragraph at index ${idx} is not accessible.`);
      }

      const insertLoc: ParaInsertLoc = toParaInsertLoc(position);
      const paraRange = targetParagraph.getRange();
      paraRange.insertText(params.text, insertLoc);
      await context.sync();

      const preview = params.text.length > 100
        ? params.text.slice(0, 100) + "..."
        : params.text;
      return `Text inserted (${insertLoc} paragraph ${idx}): "${preview}"`;
    }

    const selection = context.document.getSelection();
    selection.insertText(params.text, position as Word.InsertLocation);

    await context.sync();

    const preview = params.text.length > 100
      ? params.text.slice(0, 100) + "..."
      : params.text;
    return `Text inserted (${position}): "${preview}"`;
  });
}
