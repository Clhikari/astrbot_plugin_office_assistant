import { ExternalHyperlink, ParagraphChild, TextRun } from "docx";

import { JsonObject } from "../../core/payload";
import { RenderCliError } from "../../core/errors";
import { readSharedContract } from "../../core/shared-contracts";
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

type HyperlinkUrlContract = {
  allowed_schemes: string[];
  schemes_requiring_authority: string[];
  schemes_requiring_path: string[];
  error_message: string;
};

const HYPERLINK_URL_CONTRACT =
  readSharedContract<HyperlinkUrlContract>("hyperlink_url.json");
const DEFAULT_HYPERLINK_COLOR = "0563C1";
const SUPPORTED_HYPERLINK_PROTOCOLS = new Set(
  HYPERLINK_URL_CONTRACT.allowed_schemes.map((scheme) => `${scheme}:`),
);
const HYPERLINK_PROTOCOLS_REQUIRING_HOST = new Set(
  HYPERLINK_URL_CONTRACT.schemes_requiring_authority.map(
    (scheme) => `${scheme}:`,
  ),
);
const HYPERLINK_PROTOCOLS_REQUIRING_PATH = new Set(
  HYPERLINK_URL_CONTRACT.schemes_requiring_path.map((scheme) => `${scheme}:`),
);
const HYPERLINK_URL_INVALID_CODE = "HYPERLINK_URL_INVALID";
const HYPERLINK_URL_INVALID_MESSAGE = `Hyperlink ${HYPERLINK_URL_CONTRACT.error_message}`;

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

function throwInvalidHyperlinkUrl(target: string): never {
  throw new RenderCliError(
    HYPERLINK_URL_INVALID_CODE,
    `${HYPERLINK_URL_INVALID_MESSAGE}: ${target}`,
  );
}

function formatHyperlinkTargetForError(value: unknown): string {
  if (typeof value === "string") {
    return value;
  }
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "bigint"
  ) {
    return String(value);
  }
  if (value === null) {
    return "null";
  }
  try {
    const serialized = JSON.stringify(value);
    if (typeof serialized === "string") {
      return serialized;
    }
  } catch {
    // Fall through to String(value) when JSON serialization fails.
  }
  return String(value);
}

function normalizeHyperlinkTarget(value: unknown): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    throwInvalidHyperlinkUrl(formatHyperlinkTargetForError(value));
  }

  const target = value.trim();
  if (!target) {
    return undefined;
  }

  let parsed: URL;
  try {
    parsed = new URL(target);
  } catch {
    throwInvalidHyperlinkUrl(target);
  }

  if (!SUPPORTED_HYPERLINK_PROTOCOLS.has(parsed.protocol)) {
    throwInvalidHyperlinkUrl(target);
  }

  if (HYPERLINK_PROTOCOLS_REQUIRING_HOST.has(parsed.protocol) && !parsed.host) {
    throwInvalidHyperlinkUrl(target);
  }

  if (HYPERLINK_PROTOCOLS_REQUIRING_PATH.has(parsed.protocol) && !parsed.pathname) {
    throwInvalidHyperlinkUrl(target);
  }

  return parsed.toString();
}

export function buildRuns(
  block: JsonObject,
  theme: ThemeConfig,
  defaults?: RunDefaults,
): ParagraphChild[] {
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

  const children: ParagraphChild[] = [];
  for (const rawRun of runs) {
    const run = asObject(rawRun);
    const hyperlinkTarget = normalizeHyperlinkTarget(run.url);
    const codeFontName = defaults?.codeFontName || "Consolas";
    const bodyFontName = defaults?.fontName;
    const fontName =
      booleanValue(run.code) === true ? codeFontName : bodyFontName;
    const textRuns = buildTextRuns(stringValue(run.text), {
      bold: resolveBold(booleanValue(run.bold) === true, defaults?.emphasis),
      italics: booleanValue(run.italic) === true,
      underline:
        hyperlinkTarget || booleanValue(run.underline) === true ? {} : undefined,
      color:
        stringValue(run.color) ||
        (hyperlinkTarget ? DEFAULT_HYPERLINK_COLOR : defaultColor),
      font: fontName ? buildFontAttributes(fontName) : undefined,
      size: defaultSize,
    });
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
