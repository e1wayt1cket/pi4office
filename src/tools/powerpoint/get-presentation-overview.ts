import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { pptRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({});

type Params = Static<typeof schema>;

export function createGetPresentationOverviewTool(): AgentTool<typeof schema> {
  return {
    name: "get_presentation_overview",
    label: t("tools.presentationOverview"),
    description:
      "Get an overview of the presentation: slide count, slide titles, " +
      "and layout information. Use this at the start of a conversation " +
      "or when you need to understand the presentation structure.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      _params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = await buildOverview();
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error getting presentation overview: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function buildOverview(): Promise<string> {
  return pptRun(async (context) => {
    const presentation = context.presentation;
    presentation.load("slides");
    presentation.slides.load("items/id,items/layout");
    await context.sync();

    const slides = presentation.slides.items;
    const lines: string[] = [];

    lines.push(`## Presentation`);
    lines.push(`Slides: ${slides.length}`);
    lines.push("");

    // Load text from each slide's shapes
    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      if (!slide) continue;

      slide.load("shapes");
      slide.shapes.load("items/textFrame/textRange/text");
      await context.sync();

      const titles: string[] = [];
      for (const shape of slide.shapes.items) {
        const text = shape.textFrame?.textRange?.text?.trim();
        if (text && text.length > 0) {
          titles.push(text.slice(0, 60));
          break; // Use first non-empty shape as title
        }
      }

      const titlePreview = titles.length > 0
        ? titles[0] ?? "(empty)"
        : "(no text)";
      const layoutId = slide.layout?.id ?? "unknown";

      lines.push(`### Slide ${i + 1}`);
      lines.push(`Title: ${titlePreview}`);
      lines.push(`Layout: ${layoutId}`);
      lines.push("");
    }

    return lines.join("\n");
  });
}
