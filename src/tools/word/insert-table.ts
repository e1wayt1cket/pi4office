import { t } from "../../language/index.js";
import { Type, type Static } from "@sinclair/typebox";
import type { AgentTool, AgentToolResult } from "@earendil-works/pi-agent-core";
import { wordRun } from "./helpers.js";
import { getErrorMessage } from "../../utils/errors.js";

const schema = Type.Object({
  rows: Type.Number({
    description: "Number of rows (including optional header row). Must be at least 1.",
  }),
  columns: Type.Number({
    description: "Number of columns. Must be at least 1.",
  }),
  headers: Type.Optional(
    Type.Array(Type.String(), {
      description:
        "Optional header texts for each column. Length should match 'columns'. " +
        "The first row will contain these values and will be bold by default.",
    }),
  ),
  data: Type.Optional(
    Type.Array(Type.Array(Type.String()), {
      description:
        "Optional 2D array of cell values. First dimension is rows, second is columns. " +
        "Rows beyond the data length will be empty.",
    }),
  ),
});

type Params = Static<typeof schema>;

export function createInsertTableTool(): AgentTool<typeof schema> {
  return {
    name: "insert_table",
    label: t("tools.insertTable"),
    description:
      "Insert a table into the document at the current cursor position. " +
      "Optionally provide header texts and data rows to populate the table.",
    parameters: schema,
    execute: async (
      _toolCallId: string,
      params: Params,
    ): Promise<AgentToolResult<undefined>> => {
      try {
        const result = await insertTable(params);
        return {
          content: [{ type: "text", text: result }],
          details: undefined,
        };
      } catch (e) {
        return {
          content: [{ type: "text", text: `Error inserting table: ${getErrorMessage(e)}` }],
          details: undefined,
        };
      }
    },
  };
}

async function insertTable(params: Params): Promise<string> {
  return wordRun(async (context) => {
    const rows = Math.max(1, Math.floor(params.rows));
    const cols = Math.max(1, Math.floor(params.columns));

    // Build values array: optional headers + optional data
    const values: string[][] = [];

    if (params.headers && params.headers.length > 0) {
      const headerRow: string[] = [];
      for (let c = 0; c < cols; c++) {
        headerRow.push(params.headers[c] ?? "");
      }
      values.push(headerRow);
    }

    if (params.data && params.data.length > 0) {
      for (let r = 0; r < params.data.length && values.length < rows; r++) {
        const dataRow = params.data[r];
        if (!dataRow) continue;
        const row: string[] = [];
        for (let c = 0; c < cols; c++) {
          row.push(dataRow[c] ?? "");
        }
        values.push(row);
      }
    }

    // Fill remaining rows with empty strings
    while (values.length < rows) {
      values.push(new Array<string>(cols).fill(""));
    }

    const selection = context.document.getSelection();
    const table = selection.insertTable(rows, cols, "After", values);

    // Bold the header row if headers were provided
    if (params.headers && params.headers.length > 0 && rows >= 1) {
      table.rows.load("items");
      await context.sync();
      const firstRow = table.rows.items[0];
      if (firstRow) {
        firstRow.font.bold = true;
      }
    }

    await context.sync();

    const headerLabel = (params.headers && params.headers.length > 0)
      ? ` with headers: ${params.headers.slice(0, 3).join(", ")}`
      : "";
    return `Inserted ${rows}x${cols} table at cursor${headerLabel}.`;
  });
}
