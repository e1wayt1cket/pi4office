import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  search: Type.String({
    description: "The text to search for.",
  }),
  replace: Type.String({
    description: "The text to replace each match with.",
  }),
  match_case: Type.Optional(
    Type.Boolean({
      description: "Case-sensitive search. Defaults to true.",
    }),
  ),
  replace_all: Type.Optional(
    Type.Boolean({
      description: "Replace all matches. Defaults to true. Set to false to replace only the first match.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createReplaceTextTool(): AgentTool<typeof schema> {
  return {
    name: "replace_text",
    label: t("tools.replaceText"),
    description:
      "Find and replace text in the document. Replaces all matches by default. " +
      "The search is not case-sensitive by default.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await replaceText(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error replacing text: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function replaceText(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const matchCase = params.match_case ?? false;
    const replaceAll = params.replace_all ?? true;

    const searchResults = context.document.body.search(params.search, {
      matchCase,
      ignorePunct: true,
      ignoreSpace: true,
    });
    searchResults.load("items/length");
    await context.sync();

    const totalCount = searchResults.items.length;
    if (totalCount === 0) {
      return `No matches found for "${params.search}".`;
    }

    const count = replaceAll ? totalCount : 1;

    for (let i = 0; i < count; i++) {
      const range = searchResults.items[i];
      if (!range) continue;

      range.insertText(params.replace, "Replace");
    }

    await context.sync();

    const summary = replaceAll
      ? `Replaced ${count} occurrence(s) of "${params.search}" with "${params.replace}".`
      : `Replaced first occurrence of "${params.search}" with "${params.replace}" (${totalCount} total matches).`;
    return summary;
  });
}
