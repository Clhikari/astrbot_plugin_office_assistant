import { HeaderFooterConfig } from "./types";
import { booleanValue, stringValue } from "./utils";

export function normalizeInheritedHeaderFooter(
  config: HeaderFooterConfig,
): HeaderFooterConfig {
  const next = { ...config };
  if (usesFirstPageVariants(next)) {
    delete next.different_first_page;
    delete next.first_page_header_text;
    delete next.first_page_footer_text;
    delete next.first_page_show_page_number;
  }
  return next;
}

export function mergeHeaderFooter(
  baseConfig: HeaderFooterConfig,
  overrideConfig: HeaderFooterConfig,
): HeaderFooterConfig {
  return { ...baseConfig, ...overrideConfig };
}

export function hasHeaderFooterOverride(config: HeaderFooterConfig): boolean {
  return Object.keys(config).length > 0;
}

export function usesFirstPageVariants(config: HeaderFooterConfig): boolean {
  return (
    booleanValue(config.different_first_page) === true ||
    stringValue(config.first_page_header_text).trim().length > 0 ||
    stringValue(config.first_page_footer_text).trim().length > 0 ||
    booleanValue(config.first_page_show_page_number) !== undefined
  );
}

export function usesEvenPageVariants(config: HeaderFooterConfig): boolean {
  return (
    booleanValue(config.different_odd_even) === true ||
    stringValue(config.even_page_header_text).trim().length > 0 ||
    stringValue(config.even_page_footer_text).trim().length > 0 ||
    booleanValue(config.even_page_show_page_number) !== undefined
  );
}

export function usesSplitLayout(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
): boolean {
  return Boolean(
    stringValue(config[`${kind}_left`]).trim() ||
      stringValue(config[`${kind}_right`]).trim(),
  );
}

export function resolveHeaderFooterLeft(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
  variant: "default" | "first" | "even",
): string {
  const splitLeft = stringValue(config[`${kind}_left`]);
  if (splitLeft.trim()) {
    return splitLeft;
  }
  if (variant === "first") {
    return stringValue(config[`first_page_${kind}_text`]);
  }
  if (variant === "even") {
    return (
      stringValue(config[`even_page_${kind}_text`]) ||
      stringValue(config[`${kind}_text`])
    );
  }
  return stringValue(config[`${kind}_text`]);
}

export function resolveHeaderFooterRight(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
): string {
  return stringValue(config[`${kind}_right`]);
}

export function resolveShowPageNumberSetting(
  config: HeaderFooterConfig,
  variant: "default" | "first" | "even",
): boolean | undefined {
  if (variant === "first") {
    return (
      booleanValue(config.first_page_show_page_number) ??
      booleanValue(config.show_page_number)
    );
  }
  if (variant === "even") {
    return (
      booleanValue(config.even_page_show_page_number) ??
      booleanValue(config.show_page_number)
    );
  }
  return booleanValue(config.show_page_number);
}

export function containsPagePlaceholder(text: string): boolean {
  return text.includes("{PAGE}");
}

export function suppressPagePlaceholder(text: string): string {
  return containsPagePlaceholder(text) ? "" : text;
}
