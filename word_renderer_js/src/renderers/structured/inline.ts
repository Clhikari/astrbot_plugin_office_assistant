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

export function buildFontAttributes(fontName: string) {
  return {
    ascii: fontName,
    hAnsi: fontName,
    eastAsia: fontName,
    cs: fontName,
  };
}

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

function normalizeLineBreaks(text: string): string {
  return text.replace(/\\n/g, "\n").replace(/\r\n/g, "\n");
}

function buildTextRuns(
  text: string,
  options: Record<string, unknown>,
): TextRun[] {
  const normalizedText = normalizeLineBreaks(text);
  const segments = normalizedText.split("\n");
  if (segments.length === 1) {
    return [new TextRun({ ...options, text: normalizedText })];
  }
  return segments.map((segment, index) =>
    new TextRun({
      ...options,
      text: segment,
      break: index === 0 ? undefined : 1,
    }),
  );
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
    const fontName = defaults?.fontName;
    return buildTextRuns(stringValue(block.text), {
      bold: resolveBold(false, defaults?.emphasis),
      color: defaultColor,
      size: defaultSize,
      font: fontName ? buildFontAttributes(fontName) : undefined,
    });
  }

  return runs.flatMap((rawRun) => {
    const run = asObject(rawRun);
    const codeFontName = defaults?.codeFontName || "Consolas";
    const bodyFontName = defaults?.fontName;
    const fontName =
      booleanValue(run.code) === true ? codeFontName : bodyFontName;
    return buildTextRuns(stringValue(run.text), {
      bold: resolveBold(booleanValue(run.bold) === true, defaults?.emphasis),
      italics: booleanValue(run.italic) === true,
      underline: booleanValue(run.underline) === true ? {} : undefined,
      color: stringValue(run.color) || defaultColor,
      font: fontName ? buildFontAttributes(fontName) : undefined,
      size: defaultSize,
    });
  });
}

export function paragraphPlainText(block: Block): string {
  const runs = arrayValue(block.runs);
  if (runs.length > 0) {
    return runs.map((run) => normalizeLineBreaks(stringValue(asObject(run).text))).join("");
  }
  return normalizeLineBreaks(stringValue(block.text));
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
