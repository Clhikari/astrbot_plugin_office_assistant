import { JsonObject } from "../../core/payload";
import { THEMES } from "./constants";
import { ThemeConfig } from "./types";
import {
  asObject,
  cmToTwip,
  normalizeHexColor,
  numberValue,
  stringValue,
} from "./utils";

export function resolveTheme(metadata: JsonObject): ThemeConfig {
  const themeName = stringValue(metadata.theme_name);
  const baseTheme = THEMES[themeName] ?? THEMES.business_report;
  const documentStyle = asObject(metadata.document_style);
  const fontName = stringValue(documentStyle.font_name) || baseTheme.fontName;
  const accentColor =
    normalizeHexColor(stringValue(metadata.accent_color)) || baseTheme.accent;
  const usesCustomAccent = accentColor !== baseTheme.accent;
  const density = stringValue(metadata.density) || "comfortable";
  const densityOverrides =
    density === "compact"
      ? {
          margins: { topCm: 2.2, rightCm: 2.3, bottomCm: 2.1, leftCm: 2.4 },
          titleSpacingAfter: 14,
          headingSpaceBefore: 10,
          headingSpaceAfter: 5,
          bodyIndent: 18,
          bodySpaceAfter: 6,
          bodyLineSpacing: 1.2,
          listSpaceAfter: 4,
          tableFontSize: 9.5,
        }
      : {};
  return {
    ...baseTheme,
    ...densityOverrides,
    accent: accentColor,
    accentSoft: usesCustomAccent
      ? blendHex(accentColor, "FFFFFF", 0.84)
      : baseTheme.accentSoft,
    fontName,
    headingFontName:
      stringValue(documentStyle.heading_font_name) || fontName,
    tableFontName: stringValue(documentStyle.table_font_name) || fontName,
    codeFontName:
      stringValue(documentStyle.code_font_name) || baseTheme.codeFontName,
    summaryFill: usesCustomAccent
      ? blendHex(accentColor, "FFFFFF", 0.92)
      : baseTheme.summaryFill,
  };
}

export function buildPageMargins(
  margins: JsonObject | undefined,
  theme: ThemeConfig,
): JsonObject {
  return {
    top: cmToTwip(numberValue(margins?.top_cm) ?? theme.margins.topCm),
    right: cmToTwip(numberValue(margins?.right_cm) ?? theme.margins.rightCm),
    bottom: cmToTwip(numberValue(margins?.bottom_cm) ?? theme.margins.bottomCm),
    left: cmToTwip(numberValue(margins?.left_cm) ?? theme.margins.leftCm),
  };
}

function blendHex(source: string, target: string, ratio: number): string {
  const normalizedRatio = Math.min(Math.max(ratio, 0), 1);
  const sourceChannels = splitHexChannels(source);
  const targetChannels = splitHexChannels(target);
  const blended = sourceChannels.map((value, index) =>
    Math.round(
      value * (1 - normalizedRatio) + targetChannels[index] * normalizedRatio,
    ),
  );
  return blended.map((value) => value.toString(16).padStart(2, "0")).join("").toUpperCase();
}

function splitHexChannels(color: string): number[] {
  return [0, 2, 4].map((index) => Number.parseInt(color.slice(index, index + 2), 16));
}
