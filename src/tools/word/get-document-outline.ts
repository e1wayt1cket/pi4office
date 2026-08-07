import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({});

type Params = Static<typeof schema>;

export function createGetDocumentOutlineTool(): AgentTool<typeof schema> {
  return {
    name: "get_document_outline",
    label: t("tools.documentOutline"),
    description:
      "Get a structural outline of the document: section count, paragraph count, " +
      "heading structure, and overall document length. Use this at the start of a " +
      "conversation or when you need to understand the document structure.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = await buildOutline();
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error getting document outline: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function buildOutline(): Promise<string> {
  return wordRun(async (context) => {
    const body = context.document.body;
    const sections = context.document.sections;
    const properties = context.document.properties;

    body.paragraphs.load("items/style,items/text");
    sections.load("items");
    properties.load("title");

    await context.sync();

    const lines: string[] = [];

    const docTitle = properties.title || "(Untitled)";
    lines.push(`## Document: ${docTitle}`);
    lines.push("");
    lines.push(`Sections: ${sections.items.length}`);
    lines.push(`Paragraphs: ${body.paragraphs.items.length}`);

    // Extract headings (paragraphs with heading styles)
    const headings: Array<{ level: number; text: string }> = [];
    for (const para of body.paragraphs.items) {
      const style = para.style;
      const match = /heading\s*(\d+)/i.exec(style);
      if (match?.[1]) {
        const level = parseInt(match[1], 10);
        const preview = (para.text || "").trim().slice(0, 80);
        if (preview) {
          headings.push({ level, text: preview });
        }
      }
    }

    if (headings.length > 0) {
      lines.push("");
      lines.push("### Outline");
      for (const h of headings) {
        const indent = "  ".repeat(Math.max(0, h.level - 1));
        lines.push(`${indent}- ${h.text}`);
      }
    }

    // Document stats
    const fullText = body.text || "";
    const charCount = fullText.length;
    const wordCount = fullText.split(/\s+/).filter(Boolean).length;
    lines.push("");
    lines.push(`Characters: ${charCount}`);
    lines.push(`Words (approx): ${wordCount}`);

    return lines.join("\n");
  });
}
