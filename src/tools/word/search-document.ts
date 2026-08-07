import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  query: Type.String({
    description: "The text to search for in the document.",
  }),
  match_case: Type.Optional(
    Type.Boolean({
      description: "If true, perform a case-sensitive search. Defaults to false.",
    }),
  ),
  max_results: Type.Optional(
    Type.Number({
      description: "Maximum number of results to return. Defaults to 20.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createSearchDocumentTool(): AgentTool<typeof schema> {
  return {
    name: "search_document",
    label: t("tools.searchDocument"),
    description:
      "Search for text within the document. Returns matching paragraphs with " +
      "surrounding context.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const text = await searchDocument(params);
        return {
          content: [{ type: "text", text }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error searching document: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function searchDocument(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const maxResults = params.max_results ?? 20;
    const matchCase = params.match_case ?? false;

    const searchResults = context.document.body.search(params.query, {
      matchCase,
      ignorePunct: true,
      ignoreSpace: true,
    });
    searchResults.load("items/length");
    await context.sync();

    const lines: string[] = [];
    lines.push(`## Search: "${params.query}"`);
    lines.push("");

    const count = searchResults.items.length;
    if (count === 0) {
      lines.push("No matches found.");
      return lines.join("\n");
    }

    const limited = searchResults.items.slice(0, maxResults);
    for (let i = 0; i < limited.length; i++) {
      const range = limited[i];
      if (!range) continue;

      // Get the paragraph containing this range
      const paragraph = range.paragraphs.getFirst();
      range.load("text");
      paragraph.load("text");
      await context.sync();

      const paraText = (paragraph.text || "").trim();
      const matchText = range.text || "";

      // Show paragraph with match marked
      const preview = paraText.length > 200
        ? paraText.slice(0, 200) + "..."
        : paraText;

      lines.push(`### Match ${i + 1}`);
      lines.push(`**"${matchText}"**`);
      lines.push(`> ${preview}`);
      lines.push("");
    }

    if (count > maxResults) {
      lines.push(`... (${count - maxResults} more results. Use max_results to increase.)`);
    } else {
      lines.push(`Found ${count} match(es).`);
    }

    return lines.join("\n");
  });
}
