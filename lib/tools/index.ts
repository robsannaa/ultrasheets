/**
 * Unified Tool Registry Entry (domain-organized)
 *
 * Single, permanent export of ALL_TOOLS composed from domain modules.
 */

import { DATA_TOOLS } from "./domains/data";
import { FORMAT_TOOLS } from "./domains/format";
import { ANALYSIS_TOOLS } from "./domains/analysis";
import { STRUCTURE_TOOLS } from "./domains/structure";
import { NAVIGATION_TOOLS } from "./domains/navigation";

export const ALL_TOOLS = [
  ...DATA_TOOLS,
  ...FORMAT_TOOLS,
  ...ANALYSIS_TOOLS,
  ...STRUCTURE_TOOLS,
  ...NAVIGATION_TOOLS,
];

export type { ToolDefinition } from "../tool-executor";
