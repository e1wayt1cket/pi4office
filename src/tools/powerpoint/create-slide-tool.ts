import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { pptRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  title: Type.Optional(Type.String({
    description: "Title text for the new slide.",
  })),
  content: Type.Optional(Type.String({
    description: "Content/body text for the new slide.",
  })),
  layout: Type.Optional(Type.String({
    enum: ["blank", "titleOnly", "title", "sectionHeader", "twoColumnText", "objectAndText"],
    default: "title",
    description:
      "Layout style: blank (empty), titleOnly (title placeholder only), " +
      "title (title + subtitle), sectionHeader, twoColumnText, objectAndText.",
  })),
});

type Params = Static<typeof schema>;

function getSlideLayout(layout: string): PowerPoint.SlideLayoutType {
  switch (layout) {
    case "blank": return PowerPoint.SlideLayoutType.blank;
    case "titleOnly": return PowerPoint.SlideLayoutType.titleOnly;
    case "title": return PowerPoint.SlideLayoutType.title;
    case "sectionHeader": return PowerPoint.SlideLayoutType.sectionHeader;
    case "twoColumnText": return PowerPoint.SlideLayoutType.twoColumnText;
    case "objectAndText": return PowerPoint.SlideLayoutType.objectAndText;
    default: return PowerPoint.SlideLayoutType.title;
  }
}

export function createSlideTool(): AgentTool<typeof schema> {
  return {
    name: "create_slide",
    label: t("tools.createSlide"),
    description:
      "Add a new slide to the presentation. Optionally set title and content text. " +
      "Layout options include blank, titleOnly, title (title+subtitle), sectionHeader, " +
      "twoColumnText, and objectAndText.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await createSlide(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error creating slide: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function createSlide(params: Params): Promise<string> {
  return pptRun(async (context) => {
    const layout = getSlideLayout(params.layout ?? "title");

    context.presentation.slides.add({ layoutId: layout });
    await context.sync();

    // Reload to find the new slide at the end
    context.presentation.slides.load("items");
    await context.sync();

    const lastSlide = context.presentation.slides.items[
      context.presentation.slides.items.length - 1
    ];

    if (!lastSlide) {
      return `Slide created with layout "${params.layout ?? "title"}". Could not set text.`;
    }

    let description = `Slide created (${context.presentation.slides.items.length} total) with layout "${params.layout ?? "title"}".`;

    if (params.title || params.content) {
      lastSlide.load("shapes/items/textFrame/textRange");
      await context.sync();

      const shapes = lastSlide.shapes.items;

      if (params.title && shapes[0]?.textFrame?.textRange) {
        shapes[0].textFrame.textRange.text = params.title;
      }

      if (params.content && shapes[1]?.textFrame?.textRange) {
        shapes[1].textFrame.textRange.text = params.content;
      }

      await context.sync();

      if (params.title && params.content) {
        description += " Title and content set.";
      } else if (params.title) {
        description += " Title set.";
      } else if (params.content) {
        description += " Content set.";
      }
    }

    return description;
  });
}
