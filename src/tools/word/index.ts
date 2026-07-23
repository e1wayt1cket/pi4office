import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";
import { createGetDocumentOutlineTool } from "./get-document-outline.js";
import { createReadDocumentTool } from "./read-document.js";
import { createInsertTextTool } from "./insert-text.js";
import { createWriteDocumentTool } from "./write-document.js";
import { createSearchDocumentTool } from "./search-document.js";
import { createReplaceTextTool } from "./replace-text.js";
import { createDeleteTextTool } from "./delete-text.js";
import { createFormatDocumentTool } from "./format-document.js";
import { createApplyStyleTool } from "./apply-style.js";
import { createAddCommentTool } from "./add-comment.js";
import { createInsertTableTool } from "./insert-table.js";

export type WordToolName =
  | "get_document_outline"
  | "read_document"
  | "insert_text"
  | "write_document"
  | "search_document"
  | "replace_text"
  | "delete_text"
  | "format_document"
  | "apply_style"
  | "add_comment"
  | "insert_table";

export const WORD_TOOL_NAMES: readonly WordToolName[] = [
  "get_document_outline",
  "read_document",
  "insert_text",
  "write_document",
  "search_document",
  "replace_text",
  "delete_text",
  "format_document",
  "apply_style",
  "add_comment",
  "insert_table",
];

export type AnyWordTool = AgentTool<TSchema, DynamicValue>;

export function createWordTools(): AnyWordTool[] {
  return [
    createGetDocumentOutlineTool(),
    createReadDocumentTool(),
    createInsertTextTool(),
    createWriteDocumentTool(),
    createSearchDocumentTool(),
    createReplaceTextTool(),
    createDeleteTextTool(),
    createFormatDocumentTool(),
    createApplyStyleTool(),
    createAddCommentTool(),
    createInsertTableTool(),
  ];
}
