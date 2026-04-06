import { JsonObject } from "../../core/payload";
import { THEMES } from "./constants";
import { ThemeConfig } from "./types";
import { cmToTwip, normalizeHexColor, numberValue, stringValue } from "./utils";

export function resolveTheme(metadata: JsonObject): ThemeConfig {
  const themeName = stringValue(metadata.theme_name);
  const baseTheme = THEMES[themeName] ?? THEMES.business_report;
  const accentColor =
    normalizeHexColor(stringValue(metadata.accent_color)) || baseTheme.accent;
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
