import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { pptRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  query: Type.String({
    description: "The text to search for across all slides.",
  }),
  match_case: Type.Optional(
    Type.Boolean({
      description: "If true, perform a case-sensitive search. Defaults to false.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createSearchPresentationTool(): AgentTool<typeof schema> {
  return {
    name: "search_presentation",
    label: t("tools.searchPresentation"),
    description:
      "Search for text across all slides in the presentation. Returns matching " +
      "slides with the surrounding context.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = await searchPresentation(params);
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error searching presentation: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function searchPresentation(params: Params): Promise<string> {
  return pptRun(async (context) => {
    const presentation = context.presentation;
    presentation.slides.load("items");
    await context.sync();

    const query = params.match_case ? params.query : params.query.toLowerCase();
    const lines: string[] = [];
    lines.push(`## Search: "${params.query}"`);
    lines.push("");

    let matchCount = 0;
    const slides = presentation.slides.items;

    for (let i = 0; i < slides.length; i++) {
      const slide = slides[i];
      if (!slide) continue;

      slide.load("shapes");
      slide.shapes.load("items/textFrame/textRange/text");
      await context.sync();

      const matches: string[] = [];
      for (const shape of slide.shapes.items) {
        const text = shape.textFrame?.textRange?.text || "";
        const checkText = params.match_case ? text : text.toLowerCase();
        if (checkText.includes(query)) {
          const preview = text.trim().slice(0, 150);
          matches.push(`> ${preview}${text.length > 150 ? "..." : ""}`);
        }
      }

      if (matches.length > 0) {
        matchCount++;
        lines.push(`### Slide ${i + 1}`);
        for (const match of matches) {
          lines.push(match);
        }
        lines.push("");
      }
    }

    if (matchCount === 0) {
      lines.push("No matches found.");
    } else {
      lines.push(`Found matches in ${matchCount} slide(s).`);
    }

    return lines.join("\n");
  });
}
