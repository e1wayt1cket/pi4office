import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { pptRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  slide_index: Type.Optional(
    Type.Number({
      description: "1-based slide index to read. If omitted, reads the currently selected slide(s).",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createReadSlideTool(): AgentTool<typeof schema> {
  return {
    name: "read_slide",
    label: t("tools.readSlide"),
    description:
      "Read the text content of a slide. Specify a slide index (1-based) or " +
      "omit to read the currently selected slide(s).",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = await readSlide(params);
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error reading slide: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function readSlide(params: Params): Promise<string> {
  return pptRun(async (context) => {
    if (params.slide_index !== undefined) {
      return readSlideByIndex(context, params.slide_index);
    }

    // Read selected slides
    const selectedSlides = context.presentation.getSelectedSlides();
    selectedSlides.load("items/shapes");
    await context.sync();

    if (selectedSlides.items.length === 0) {
      return "No slides selected. Specify a slide_index to read a specific slide.";
    }

    const lines: string[] = [];
    for (const slide of selectedSlides.items) {
      slide.shapes.load("items/textFrame/textRange/text");
      await context.sync();
      lines.push(formatSlideShapes(slide.shapes.items, slide.id ?? ""));
    }

    return lines.join("\n\n");
  });
}

async function readSlideByIndex(context: PowerPoint.RequestContext, index: number): Promise<string> {
  const slides = context.presentation.slides;
  slides.load("items/id");
  await context.sync();

  const zeroIndex = index - 1;
  if (zeroIndex < 0 || zeroIndex >= slides.items.length) {
    return `Slide index ${index} is out of range. Presentation has ${slides.items.length} slide(s).`;
  }

  const slide = slides.items[zeroIndex];
  if (!slide) return `Slide ${index} not found.`;

  slide.load("shapes");
  slide.shapes.load("items/textFrame/textRange/text");
  await context.sync();

  return formatSlideShapes(slide.shapes.items, slide.id ?? "");
}

function formatSlideShapes(shapes: PowerPoint.Shape[], slideId: string): string {
  const lines: string[] = [];
  lines.push(`## Slide (id: ${slideId})`);
  lines.push("");

  for (let i = 0; i < shapes.length; i++) {
    const shape = shapes[i];
    if (!shape) continue;
    const text = shape.textFrame?.textRange?.text?.trim();
    if (text && text.length > 0) {
      const shapeType = shape.type || "Text";
      lines.push(`**[${shapeType}]** ${text}`);
      lines.push("");
    }
  }

  if (lines.length === 2) {
    lines.push("(Slide has no text content)");
  }

  return lines.join("\n");
}
