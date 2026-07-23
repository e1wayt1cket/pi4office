import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  text: Type.String({
    description: "The text to write into the document.",
  }),
  location: Type.Optional(
    Type.String({
      enum: ["end", "start", "selection"],
      description:
        "Where to write: 'end' (default) appends at document end, " +
        "'start' inserts at document beginning, " +
        "'selection' inserts at current cursor/selection.",
    }),
  ),
  paragraph_index: Type.Optional(
    Type.Number({
      description:
        "0-based paragraph index. When set, text is inserted AFTER this paragraph. " +
        "Takes precedence over 'location'. Use with get_document_outline or read_document " +
        "to find the right paragraph index.",
    }),
  ),
  position: Type.Optional(
    Type.String({
      enum: ["Before", "After"],
      description:
        "When paragraph_index is set, whether to insert before or after it. Defaults to 'After'.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createWriteDocumentTool(): AgentTool<typeof schema> {
  return {
    name: "write_document",
    label: t("tools.writeDocument"),
    description:
      "Write text into the document at a specific location. By default appends to the end. " +
      "Use paragraph_index to insert after a specific paragraph (e.g. after a heading). " +
      "Use 'start' to insert at the document beginning, or 'selection' to insert at cursor.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await writeDocument(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error writing to document: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

type BodyInsertLoc = "Start" | "End" | "Replace";

async function writeDocument(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const body = context.document.body;
    const location = params.location ?? "end";

    // Paragraph-indexed insertion — most precise, takes priority
    if (params.paragraph_index !== undefined) {
      body.paragraphs.load("items");
      await context.sync();

      const paragraphs = body.paragraphs.items;
      const idx = params.paragraph_index;
      if (idx < 0 || idx >= paragraphs.length) {
        throw new Error(
          `Paragraph index ${idx} out of range (document has ${paragraphs.length} paragraphs, 0-based).`,
        );
      }

      const targetParagraph = paragraphs[idx];
      if (!targetParagraph) {
        throw new Error(`Paragraph at index ${idx} is not accessible.`);
      }

      const insertPos: "Before" | "After" = (params.position ?? "After") as "Before" | "After";
      const paraRange = targetParagraph.getRange();
      paraRange.insertText(params.text, insertPos);
      await context.sync();

      const rel = insertPos === "Before" ? "before" : "after";
      const preview = params.text.length > 100 ? params.text.slice(0, 100) + "..." : params.text;
      return `Text written ${rel} paragraph ${idx}: "${preview}"`;
    }

    // Location-based insertion
    let wordLocation: BodyInsertLoc;
    let label: string;

    switch (location) {
      case "start":
        wordLocation = "Start";
        label = "start of document";
        break;
      case "selection":
        wordLocation = "Replace";
        label = "selection/cursor";
        break;
      case "end":
      default:
        wordLocation = "End";
        label = "end of document";
        break;
    }

    if (location === "selection") {
      const selection = context.document.getSelection();
      selection.insertText(params.text, wordLocation);
    } else {
      body.insertText(params.text, wordLocation);
    }

    await context.sync();

    const preview = params.text.length > 100 ? params.text.slice(0, 100) + "..." : params.text;
    return `Text written at ${label}: "${preview}"`;
  });
}
