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

const LAYOUT_TYPE_MAP: Record<string, PowerPoint.SlideLayoutType> = {
  blank: PowerPoint.SlideLayoutType.blank,
  titleOnly: PowerPoint.SlideLayoutType.titleOnly,
  title: PowerPoint.SlideLayoutType.title,
  sectionHeader: PowerPoint.SlideLayoutType.sectionHeader,
  twoColumnText: PowerPoint.SlideLayoutType.twoColumnText,
  objectAndText: PowerPoint.SlideLayoutType.objectAndText,
};

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

async function resolveLayoutId(
  context: PowerPoint.RequestContext,
  requestedType: PowerPoint.SlideLayoutType,
): Promise<string> {
  const master = context.presentation.slideMasters.getItemAt(0);
  master.layouts.load("items");
  await context.sync();

  const layouts = master.layouts.items;
  if (layouts.length === 0) {
    throw new Error("No slide layouts available from slide master.");
  }

  for (const layout of layouts) {
    layout.load("type,id");
  }
  await context.sync();

  for (const layout of layouts) {
    if (layout.type === requestedType) {
      return layout.id;
    }
  }

  // Fallback: return the first layout's id
  const first = layouts[0];
  if (first) return first.id;
  throw new Error("No slide layouts available from slide master.");
}

function isPlaceholder(shape: PowerPoint.Shape): boolean {
  return shape.type === PowerPoint.ShapeType.placeholder;
}

function matchPlaceholderType(
  shape: PowerPoint.Shape,
  targetType: PowerPoint.PlaceholderType,
): boolean {
  return isPlaceholder(shape) && shape.placeholderFormat.type === targetType;
}

async function createSlide(params: Params): Promise<string> {
  return pptRun(async (context) => {
    const requestedType = LAYOUT_TYPE_MAP[params.layout ?? "title"]
      ?? PowerPoint.SlideLayoutType.title;
    const layoutId = await resolveLayoutId(context, requestedType);

    context.presentation.slides.add({ layoutId });
    await context.sync();

    context.presentation.slides.load("items");
    await context.sync();

    const items = context.presentation.slides.items;
    const lastSlide = items[items.length - 1];

    if (!lastSlide) {
      return `Slide created with layout "${params.layout ?? "title"}". Could not set text.`;
    }

    let description = `Slide created (${items.length} total) with layout "${params.layout ?? "title"}".`;

    if (params.title || params.content) {
      lastSlide.shapes.load("items/type,items/placeholderFormat/type");
      await context.sync();

      const shapes = lastSlide.shapes.items;

      let titleShape: PowerPoint.Shape | undefined;

      if (params.title) {
        for (const s of shapes) {
          if (
            matchPlaceholderType(s, PowerPoint.PlaceholderType.title) ||
            matchPlaceholderType(s, PowerPoint.PlaceholderType.centerTitle) ||
            matchPlaceholderType(s, PowerPoint.PlaceholderType.verticalTitle)
          ) {
            titleShape = s;
            break;
          }
        }
        if (titleShape?.textFrame) {
          titleShape.textFrame.textRange.text = params.title;
        }
      }

      if (params.content) {
        // Prefer Subtitle, Body, or Content placeholder (not used for title)
        let contentShape: PowerPoint.Shape | undefined;
        for (const s of shapes) {
          if (s === titleShape) continue;
          if (
            matchPlaceholderType(s, PowerPoint.PlaceholderType.subtitle) ||
            matchPlaceholderType(s, PowerPoint.PlaceholderType.body) ||
            matchPlaceholderType(s, PowerPoint.PlaceholderType.content)
          ) {
            contentShape = s;
            break;
          }
        }
        // Fallback: second Title placeholder
        if (!contentShape) {
          for (const s of shapes) {
            if (s === titleShape) continue;
            if (matchPlaceholderType(s, PowerPoint.PlaceholderType.title)) {
              contentShape = s;
              break;
            }
          }
        }
        if (contentShape?.textFrame) {
          contentShape.textFrame.textRange.text = params.content;
        }
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
