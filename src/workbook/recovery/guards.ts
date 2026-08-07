import { isDictionaryValue } from "./helpers.js";

/** Runtime guards for persisted recovery payloads. */

import type { RecoveryConditionalFormatRule } from "../recovery-states.js";
import {
  isRecoveryConditionalCellValueOperator,
  isRecoveryConditionalColorScaleState,
  isRecoveryConditionalDataBarState,
  isRecoveryConditionalIconSetState,
  isRecoveryConditionalPresetCriterion,
  isRecoveryConditionalTextOperator,
  isRecoveryConditionalTopBottomCriterionType,
} from "./conditional-format-normalization.js";

export function isRecoveryConditionalFormatRule(value: DynamicValue): value is RecoveryConditionalFormatRule {
  if (!isDictionaryValue(value)) return false;

  if (value.stopIfTrue !== undefined && typeof value.stopIfTrue !== "boolean") return false;
  if (value.formula !== undefined && typeof value.formula !== "string") return false;
  if (value.formula1 !== undefined && typeof value.formula1 !== "string") return false;
  if (value.formula2 !== undefined && typeof value.formula2 !== "string") return false;
  if (value.text !== undefined && typeof value.text !== "string") return false;
  if (value.rank !== undefined && (typeof value.rank !== "number" || !Number.isFinite(value.rank))) return false;
  if (value.fillColor !== undefined && typeof value.fillColor !== "string") return false;
  if (value.fontColor !== undefined && typeof value.fontColor !== "string") return false;
  if (value.bold !== undefined && typeof value.bold !== "boolean") return false;
  if (value.italic !== undefined && typeof value.italic !== "boolean") return false;
  if (value.underline !== undefined && typeof value.underline !== "boolean") return false;
  if (value.appliesToAddress !== undefined && typeof value.appliesToAddress !== "string") return false;

  const type = value.type;
  if (type === "custom") {
    return typeof value.formula === "string";
  }

  if (type === "cell_value") {
    return isRecoveryConditionalCellValueOperator(value.operator) && typeof value.formula1 === "string";
  }

  if (type === "text_comparison") {
    return isRecoveryConditionalTextOperator(value.textOperator) && typeof value.text === "string";
  }

  if (type === "top_bottom") {
    return isRecoveryConditionalTopBottomCriterionType(value.topBottomType) && typeof value.rank === "number";
  }

  if (type === "preset_criteria") {
    return isRecoveryConditionalPresetCriterion(value.presetCriterion);
  }

  if (type === "data_bar") {
    return isRecoveryConditionalDataBarState(value.dataBar);
  }

  if (type === "color_scale") {
    return isRecoveryConditionalColorScaleState(value.colorScale);
  }

  if (type === "icon_set") {
    return isRecoveryConditionalIconSetState(value.iconSet);
  }

  return false;
}
