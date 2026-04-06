import { TextRun } from "docx";

import { JsonObject } from "../../core/payload";
import { Block, RunDefaults, ThemeConfig } from "./types";
import {
  arrayValue,
  asObject,
  booleanValue,
  halfPoint,
  numberValue,
  resolveBold,
  resolveTextColor,
  stringValue,
} from "./utils";

export function normalizeInlineItem(
  item: unknown,
  theme: ThemeConfig,
  defaults?: RunDefaults,
): { runs: TextRun[] } {
  if (typeof item === "string") {
    return { runs: buildRuns({ text: item }, theme, defaults) };
  }
  const obj = asObject(item);
  if (arrayValue(obj.runs).length > 0) {
    return { runs: buildRuns(obj, theme, defaults) };
  }
  return { runs: buildRuns({ text: stringValue(obj.text) }, theme, defaults) };
}

export function buildRuns(
  block: JsonObject,
  theme: ThemeConfig,
  defaults?: RunDefaults,
): TextRun[] {
  const runs = arrayValue(block.runs);
  const defaultColor = resolveTextColor(theme, defaults?.emphasis);
  const defaultSize =
    defaults?.fontSize !== undefined
      ? halfPoint(defaults.fontSize * (defaults.fontScale ?? 1))
      : undefined;

  if (runs.length === 0) {
    return [
      new TextRun({
        text: stringValue(block.text),
        bold: resolveBold(false, defaults?.emphasis),
        color: defaultColor,
        size: defaultSize,
      }),
    ];
  }

  return runs.map((rawRun) => {
    const run = asObject(rawRun);
    return new TextRun({
      text: stringValue(run.text),
      bold: resolveBold(booleanValue(run.bold) === true, defaults?.emphasis),
      italics: booleanValue(run.italic) === true,
      underline: booleanValue(run.underline) === true ? {} : undefined,
      color: stringValue(run.color) || defaultColor,
      font: booleanValue(run.code) === true ? "Consolas" : undefined,
      size: defaultSize,
    });
  });
}

export function paragraphPlainText(block: Block): string {
  const runs = arrayValue(block.runs);
  if (runs.length > 0) {
    return runs.map((run) => stringValue(asObject(run).text)).join("");
  }
  return stringValue(block.text);
}

export function mergeStyleDefaults(
  style: JsonObject,
  defaults: { align?: string; emphasis?: string; fontScale?: number },
): JsonObject {
  return {
    ...style,
    align: stringValue(style.align) || defaults.align,
    emphasis: stringValue(style.emphasis) || defaults.emphasis,
    font_scale: numberValue(style.font_scale) ?? defaults.fontScale,
  };
}

export function mergeLayoutDefaults(
  layout: JsonObject,
  defaults: { spacingBefore?: number; spacingAfter?: number },
): JsonObject {
  return {
    ...layout,
    spacing_before: numberValue(layout.spacing_before) ?? defaults.spacingBefore,
    spacing_after: numberValue(layout.spacing_after) ?? defaults.spacingAfter,
  };
}
