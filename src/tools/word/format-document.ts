import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  bold: Type.Optional(Type.Boolean({
    description: "Apply or remove bold formatting.",
  })),
  italic: Type.Optional(Type.Boolean({
    description: "Apply or remove italic formatting.",
  })),
  underline: Type.Optional(Type.Boolean({
    description: "Apply or remove underline formatting.",
  })),
  font_name: Type.Optional(Type.String({
    description: "Font name to apply (e.g. 'Calibri', 'Arial').",
  })),
  font_size: Type.Optional(Type.Number({
    description: "Font size in points.",
  })),
  font_color: Type.Optional(Type.String({
    description: "Font color as hex (e.g. '#FF0000') or CSS color name.",
  })),
  highlight_color: Type.Optional(Type.String({
    description: "Text highlight color (e.g. 'Yellow', 'Cyan', 'Lime').",
  })),
  alignment: Type.Optional(Type.String({
    enum: ["Left", "Centered", "Right", "Justified"],
    description: "Paragraph alignment.",
  })),
});

type Params = Static<typeof schema>;

// Deferred: referencing Word.Alignment at module scope throws when the bundle
// loads outside a real Word host (browser-host fallback), so resolve lazily.
function resolveAlignment(name: string | undefined): Word.Alignment {
  switch (name) {
    case "Centered": return Word.Alignment.centered;
    case "Right": return Word.Alignment.right;
    case "Justified": return Word.Alignment.justified;
    default: return Word.Alignment.left;
  }
}

function buildChangeList(params: Params): string[] {
  const changes: string[] = [];
  if (params.bold !== undefined) changes.push(`bold: ${params.bold}`);
  if (params.italic !== undefined) changes.push(`italic: ${params.italic}`);
  if (params.underline !== undefined) changes.push(`underline: ${params.underline}`);
  if (params.font_name !== undefined) changes.push(`font: ${params.font_name}`);
  if (params.font_size !== undefined) changes.push(`size: ${params.font_size}pt`);
  if (params.font_color !== undefined) changes.push(`color: ${params.font_color}`);
  if (params.highlight_color !== undefined) changes.push(`highlight: ${params.highlight_color}`);
  if (params.alignment !== undefined) changes.push(`alignment: ${params.alignment}`);
  return changes;
}

export function createFormatDocumentTool(): AgentTool<typeof schema> {
  return {
    name: "format_document",
    label: t("tools.formatDocument"),
    description:
      "Apply formatting to the currently selected text in the document. " +
      "Supports bold, italic, underline, font name/size/color, highlight, " +
      "and paragraph alignment. Only specified properties are changed.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await formatSelection(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error formatting document: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function formatSelection(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const selection = context.document.getSelection();
    const font = selection.font;

    if (params.bold !== undefined) font.bold = params.bold;
    if (params.italic !== undefined) font.italic = params.italic;
    if (params.underline !== undefined) {
      font.underline = params.underline
        ? Word.UnderlineType.single
        : Word.UnderlineType.none;
    }
    if (params.font_name !== undefined) font.name = params.font_name;
    if (params.font_size !== undefined) font.size = params.font_size;
    if (params.font_color !== undefined) font.color = params.font_color;
    if (params.highlight_color !== undefined) font.highlightColor = params.highlight_color;
    if (params.alignment !== undefined) {
      const firstParagraph = selection.paragraphs.getFirst();
      firstParagraph.alignment = resolveAlignment(params.alignment);
    }

    await context.sync();

    const changes = buildChangeList(params);
    return changes.length > 0
      ? `Formatting applied to selection: ${changes.join(", ")}.`
      : "No formatting changes specified. Check the parameters.";
  });
}
