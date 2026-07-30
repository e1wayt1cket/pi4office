/**
 * Tool registry — creates all built-in tools for the agent.
 *
 * Canonical source of truth for core tools lives in `src/tools/registry.ts`.
 * Experimental/non-core tools are appended here.
 */

import type { OfficeAppType, SpreadsheetHostKind } from "../host/index.js";
import { createCoreTools, type AnyCoreTool } from "./registry.js";
import { selectOfficeCoupledToolForHost } from "./host-selection.js";
import type { SkillReadCache } from "../skills/read-cache.js";
import { createTmuxTool } from "./tmux.js";
import { createPythonRunTool } from "./python-run.js";
import { createLibreOfficeConvertTool } from "./libreoffice-convert.js";
import { createPythonTransformRangeTool } from "./python-transform-range.js";
import { createFilesTool } from "./files.js";
import { createExecuteOfficeJsTool } from "./execute-office-js.js";
import { createExecuteWpsJsTool } from "./execute-wps-js.js";
import {
  createExtensionsManagerTool,
  type ExtensionsManagerToolRuntime,
} from "./extensions-manager.js";
import { createWordTools } from "./word/index.js";

export interface CreateAllToolsOptions {
  hostKind?: SpreadsheetHostKind;
  appType?: OfficeAppType;
  getExtensionManager?: () => ExtensionsManagerToolRuntime | null;
  getSessionId?: () => string | null;
  skillReadCache?: SkillReadCache;
}

export function createAllTools(options: CreateAllToolsOptions = {}): AnyCoreTool[] {
  const getExtensionManager = options.getExtensionManager ?? (() => null);
  const hostKind = options.hostKind ?? "office";
  const appType = options.appType ?? "excel";

  const skills = {
    ...(options.getSessionId !== undefined ? { getSessionId: options.getSessionId } : {}),
    ...(options.skillReadCache !== undefined ? { readCache: options.skillReadCache } : {}),
  };

  const tools: AnyCoreTool[] = [
    ...createCoreTools({
      hostKind,
      appType,
      skills,
    }),
    // Shared across all app types
    createTmuxTool(),
    createPythonRunTool(),
    createLibreOfficeConvertTool(),
    createFilesTool(),
    createExtensionsManagerTool({ getManager: getExtensionManager }),
  ];

  // Excel-specific non-core tools
  if (appType === "excel" || appType === "unknown") {
    tools.push(
      selectOfficeCoupledToolForHost(createPythonTransformRangeTool(), hostKind),
      selectOfficeCoupledToolForHost(createExecuteOfficeJsTool(), hostKind),
    );
  }

  // execute_office_js for Word (different API surface)
  if (appType === "word") {
    tools.push(
      selectOfficeCoupledToolForHost(createExecuteOfficeJsTool(), hostKind),
    );
  }

  // Word-specific tools
  if (appType === "word") {
    tools.push(...createWordTools());
  }

  if (hostKind === "wps") {
    tools.push(createExecuteWpsJsTool());
  }

  return tools;
}
