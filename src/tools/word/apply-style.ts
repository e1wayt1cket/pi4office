import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  style_name: Type.String({
    description:
      "The Word style name to apply. Common values: 'Heading 1', 'Heading 2', " +
      "'Heading 3', 'Normal', 'Title', 'Subtitle', 'Quote', 'List Paragraph'. " +
      "The style must exist in the document.",
  }),
  paragraph_index: Type.Optional(
    Type.Number({
      description:
        "0-based paragraph index to apply the style to. If omitted, applies to all paragraphs " +
        "in the current selection.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createApplyStyleTool(): AgentTool<typeof schema> {
  return {
    name: "apply_style",
    label: t("tools.applyStyle"),
    description:
      "Apply a Word named style to one or more paragraphs. " +
      "Targets the current selection by default, or a specific paragraph by index. " +
      "Use get_document_outline to see existing headings and paragraph structure.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await applyStyle(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error applying style: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function applyStyle(params: Params): Promise<string> {
  return wordRun(async (context) => {
    let targetParagraphs: Word.Paragraph[];

    if (params.paragraph_index !== undefined) {
      const body = context.document.body;
      body.paragraphs.load("items");
      await context.sync();

      const idx = params.paragraph_index;
      const paragraphs = body.paragraphs.items;
      if (idx < 0 || idx >= paragraphs.length) {
        return `Paragraph index ${idx} out of range (document has ${paragraphs.length} paragraphs).`;
      }

      const para = paragraphs[idx];
      if (!para) {
        return `Paragraph at index ${idx} is not accessible.`;
      }

      targetParagraphs = [para];
    } else {
      const selection = context.document.getSelection();
      selection.paragraphs.load("items");
      await context.sync();

      targetParagraphs = selection.paragraphs.items;
      if (targetParagraphs.length === 0) {
        return "No paragraphs in current selection.";
      }
    }

    for (const para of targetParagraphs) {
      if (para) {
        para.style = params.style_name;
      }
    }

    await context.sync();

    const scope = params.paragraph_index !== undefined
      ? `paragraph ${params.paragraph_index}`
      : `${targetParagraphs.length} selected paragraph(s)`;
    return `Applied style "${params.style_name}" to ${scope}.`;
  });
}
