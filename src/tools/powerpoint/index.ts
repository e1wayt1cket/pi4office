import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";
import { createGetPresentationOverviewTool } from "./get-presentation-overview.js";
import { createReadSlideTool } from "./read-slide.js";
import { createSetSlideTextTool } from "./set-slide-text.js";
import { createSearchPresentationTool } from "./search-presentation.js";
import { createSlideTool } from "./create-slide-tool.js";

export type PptToolName =
  | "get_presentation_overview"
  | "read_slide"
  | "set_slide_text"
  | "search_presentation"
  | "create_slide";

export const PPT_TOOL_NAMES: readonly PptToolName[] = [
  "get_presentation_overview",
  "read_slide",
  "set_slide_text",
  "search_presentation",
  "create_slide",
];

export type AnyPptTool = AgentTool<TSchema, DynamicValue>;

export function createPptTools(): AnyPptTool[] {
  return [
    createGetPresentationOverviewTool(),
    createReadSlideTool(),
    createSetSlideTextTool(),
    createSearchPresentationTool(),
    createSlideTool(),
  ];
}
