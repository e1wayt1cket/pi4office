import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  selection_only: Type.Optional(
    Type.Boolean({
      description: "If true, read only the currently selected text. If false or omitted, read the full document body.",
    }),
  ),
  max_paragraphs: Type.Optional(
    Type.Number({
      description: "Maximum number of paragraphs to return. Defaults to 50.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createReadDocumentTool(): AgentTool<typeof schema> {
  return {
    name: "read_document",
    label: t("tools.readDocument"),
    description:
      "Read text from the document. By default reads the full document body " +
      "(up to max_paragraphs). Set selection_only to true to read only the " +
      "currently selected text.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = await readDocument(params);
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error reading document: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function readDocument(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const maxParagraphs = params.max_paragraphs ?? 50;

    if (params.selection_only) {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      const text = selection.text || "";
      if (!text.trim()) return "(Selection is empty.)";
      return `## Selected Text\n\n${text}`;
    }

    const body = context.document.body;
    body.paragraphs.load("items/text,items/style");
    await context.sync();

    const allParagraphs = body.paragraphs.items;
    const paragraphs = allParagraphs.slice(0, maxParagraphs);
    const lines: string[] = [];

    lines.push("## Document Content");
    lines.push("");

    for (let i = 0; i < paragraphs.length; i++) {
      const para = paragraphs[i];
      if (!para) continue;
      const text = para.text?.trim() ?? "";
      const style = para.style || "";

      if (text.length === 0) {
        lines.push("");
        continue;
      }

      // Prefix headings
      const headingMatch = /heading\s*(\d+)/i.exec(style);
      if (headingMatch?.[1]) {
        const level = parseInt(headingMatch[1], 10);
        const prefix = "#".repeat(Math.min(level, 6));
        lines.push(`${prefix} ${text}`);
      } else {
        lines.push(text);
      }
    }

    if (allParagraphs.length > maxParagraphs) {
      lines.push("");
      lines.push(`... (${allParagraphs.length - maxParagraphs} more paragraphs. Use max_paragraphs to increase.)`);
    }

    return lines.join("\n");
  });
}
