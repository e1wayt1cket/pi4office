import type { AgentTool } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";
import { createGetDocumentOutlineTool } from "./get-document-outline.js";
import { createReadDocumentTool } from "./read-document.js";
import { createInsertTextTool } from "./insert-text.js";
import { createSearchDocumentTool } from "./search-document.js";
import { createFormatDocumentTool } from "./format-document.js";
import { createAddCommentTool } from "./add-comment.js";

export type WordToolName =
  | "get_document_outline"
  | "read_document"
  | "insert_text"
  | "search_document"
  | "format_document"
  | "add_comment";

export const WORD_TOOL_NAMES: readonly WordToolName[] = [
  "get_document_outline",
  "read_document",
  "insert_text",
  "search_document",
  "format_document",
  "add_comment",
];

export type AnyWordTool = AgentTool<TSchema, DynamicValue>;

export function createWordTools(): AnyWordTool[] {
  return [
    createGetDocumentOutlineTool(),
    createReadDocumentTool(),
    createInsertTextTool(),
    createSearchDocumentTool(),
    createFormatDocumentTool(),
    createAddCommentTool(),
  ];
}
