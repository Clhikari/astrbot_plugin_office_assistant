export interface PptTheme {
  backgroundColor: string;
  titleColor: string;
  bodyColor: string;
  accentColor: string;
  fontFace: string;
  titleFontFace: string;
}

const DEFAULT_THEME: PptTheme = {
  backgroundColor: "FFFFFF",
  titleColor: "1F2937",
  bodyColor: "374151",
  accentColor: "2563EB",
  fontFace: "Microsoft YaHei",
  titleFontFace: "Microsoft YaHei",
};

const THEMES: Record<string, PptTheme> = {
  business_report: DEFAULT_THEME,
  modern_minimal: {
    backgroundColor: "FAFAFA",
    titleColor: "111827",
    bodyColor: "4B5563",
    accentColor: "6366F1",
    fontFace: "Microsoft YaHei",
    titleFontFace: "Microsoft YaHei",
  },
  dark: {
    backgroundColor: "1F2937",
    titleColor: "F9FAFB",
    bodyColor: "D1D5DB",
    accentColor: "60A5FA",
    fontFace: "Microsoft YaHei",
    titleFontFace: "Microsoft YaHei",
  },
};

export function resolveTheme(themeName: string | undefined): PptTheme {
  if (themeName && themeName in THEMES) {
    return THEMES[themeName];
  }
  return DEFAULT_THEME;
}
