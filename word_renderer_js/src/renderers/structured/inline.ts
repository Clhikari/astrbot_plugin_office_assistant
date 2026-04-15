import { ExternalHyperlink, ParagraphChild } from "docx";

import { JsonObject } from "../../core/payload";
import { Block, RunDefaults, ThemeConfig } from "./types";
import {
  buildTextRuns,
  DEFAULT_HYPERLINK_COLOR,
  normalizeHyperlinkTarget,
  normalizeLineBreaks,
} from "./run-helpers";
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
): { children: ParagraphChild[] } {
  if (typeof item === "string") {
    return { children: buildRuns({ text: item }, theme, defaults) };
  }
  const obj = asObject(item);
  if (arrayValue(obj.runs).length > 0) {
    return { children: buildRuns(obj, theme, defaults) };
  }
  return { children: buildRuns({ text: stringValue(obj.text) }, theme, defaults) };
}

function resolveRunTextOptions(
  run: JsonObject,
  theme: ThemeConfig,
  defaults?: RunDefaults,
  hyperlinkTarget?: string,
): Record<string, unknown> {
  const defaultColor = stringValue(defaults?.color) || resolveTextColor(theme, defaults?.emphasis);
  const fontName =
    stringValue(run.font_name) ||
    (booleanValue(run.code) === true ? defaults?.codeFontName || "Consolas" : defaults?.fontName);
  const runFontScale = numberValue(run.font_scale) ?? defaults?.fontScale ?? 1;
  const bold = booleanValue(run.bold) ?? defaults?.bold ?? resolveBold(false, defaults?.emphasis);
  const italic = booleanValue(run.italic) ?? defaults?.italic ?? false;
  const underline = booleanValue(run.underline) ?? defaults?.underline ?? false;
  const strikethrough =
    booleanValue(run.strikethrough) ?? defaults?.strikethrough ?? false;

  return {
    bold,
    italics: italic,
    underline: hyperlinkTarget || underline ? {} : undefined,
    strike: strikethrough,
    color:
      stringValue(run.color) ||
      (hyperlinkTarget ? DEFAULT_HYPERLINK_COLOR : defaultColor),
    font: fontName ? buildFontAttributes(fontName) : undefined,
    size:
      defaults?.fontSize !== undefined
        ? halfPoint(defaults.fontSize * runFontScale)
        : undefined,
  };
}

export function buildRuns(
  block: JsonObject,
  theme: ThemeConfig,
  defaults?: RunDefaults,
): ParagraphChild[] {
  const runs = arrayValue(block.runs);

  if (runs.length === 0) {
    return buildTextRuns(
      stringValue(block.text),
      resolveRunTextOptions({}, theme, defaults),
    );
  }

  const children: ParagraphChild[] = [];
  for (const rawRun of runs) {
    const run = asObject(rawRun);
    const hyperlinkTarget = normalizeHyperlinkTarget(run.url);
    const textRuns = buildTextRuns(
      stringValue(run.text),
      resolveRunTextOptions(run, theme, defaults, hyperlinkTarget),
    );
    if (!hyperlinkTarget) {
      children.push(...textRuns);
      continue;
    }
    children.push(
      new ExternalHyperlink({ link: hyperlinkTarget, children: textRuns }),
    );
  }
  return children;
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
