import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { pptRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  slide_index: Type.Number({
    description: "The 1-based index of the slide to modify.",
  }),
  text: Type.String({
    description: "Content text to set on the slide.",
  }),
  shape_index: Type.Optional(
    Type.Number({
      description: "Zero-based index of the shape to modify. " +
        "0 = title placeholder (default), 1 = content placeholder, etc.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createSetSlideTextTool(): AgentTool<typeof schema> {
  return {
    name: "set_slide_text",
    label: t("tools.setSlideText"),
    description:
      "Set or replace text on a specific slide shape. Specify the slide index (1-based) " +
      "and the text to set. Optionally specify which shape to target by index " +
      "(0 = title, 1 = content, etc.).",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = await setSlideText(params);
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error setting slide text: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function setSlideText(params: Params): Promise<string> {
  return pptRun(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    const zeroIndex = params.slide_index - 1;
    if (zeroIndex < 0 || zeroIndex >= slides.items.length) {
      return `Slide index ${params.slide_index} is out of range. Presentation has ${slides.items.length} slide(s).`;
    }

    const slide = slides.items[zeroIndex];
    if (!slide) return `Slide ${params.slide_index} not found.`;

    slide.load("shapes");
    slide.shapes.load("items/textFrame/textRange");
    await context.sync();

    const shapeIndex = params.shape_index ?? 0;
    const shapes = slide.shapes.items;

    if (shapeIndex >= shapes.length) {
      return `Shape index ${shapeIndex} is out of range. Slide has ${shapes.length} shape(s).`;
    }

    const shape = shapes[shapeIndex];
    if (!shape?.textFrame?.textRange) {
      return `Shape ${shapeIndex} does not support text.`;
    }

    shape.textFrame.textRange.text = params.text;
    await context.sync();

    return `Text set on slide ${params.slide_index}, shape ${shapeIndex}.`;
  });
}
