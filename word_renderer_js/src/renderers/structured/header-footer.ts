import {
  Footer,
  Header,
  Paragraph,
} from "docx";

import { HeaderFooterConfig, ThemeConfig } from "./types";
import { buildHeaderFooterParagraphs } from "./header-footer-content";
import {
  hasHeaderFooterOverride as hasHeaderFooterOverrideRule,
  mergeHeaderFooter as mergeHeaderFooterRule,
  normalizeInheritedHeaderFooter as normalizeInheritedHeaderFooterRule,
  usesEvenPageVariants as usesEvenPageVariantsRule,
  usesFirstPageVariants as usesFirstPageVariantsRule,
} from "./header-footer-rules";

export function buildHeaders(
  config: HeaderFooterConfig,
  theme: ThemeConfig,
): { default?: Header; first?: Header; even?: Header } {
  const headers: { default?: Header; first?: Header; even?: Header } = {};
  const normalHeader = buildHeaderFooterParagraphs(
    config,
    "header",
    "default",
    theme,
  );
  if (normalHeader) {
    headers.default = new Header({ children: normalHeader });
  } else if (hasHeaderFooterOverride(config)) {
    headers.default = new Header({ children: [new Paragraph("")] });
  }
  if (usesFirstPageVariants(config)) {
    headers.first = new Header({
      children:
        buildHeaderFooterParagraphs(config, "header", "first", theme) ?? [
          new Paragraph(""),
        ],
    });
  }
  if (usesEvenPageVariants(config)) {
    headers.even = new Header({
      children:
        buildHeaderFooterParagraphs(config, "header", "even", theme) ?? [
          new Paragraph(""),
        ],
    });
  }
  return headers;
}

export function buildFooters(
  config: HeaderFooterConfig,
  theme: ThemeConfig,
): { default?: Footer; first?: Footer; even?: Footer } {
  const footers: { default?: Footer; first?: Footer; even?: Footer } = {};
  const normalFooter = buildHeaderFooterParagraphs(
    config,
    "footer",
    "default",
    theme,
  );
  if (normalFooter) {
    footers.default = new Footer({ children: normalFooter });
  } else if (hasHeaderFooterOverride(config)) {
    footers.default = new Footer({ children: [new Paragraph("")] });
  }
  if (usesFirstPageVariants(config)) {
    footers.first = new Footer({
      children:
        buildHeaderFooterParagraphs(config, "footer", "first", theme) ?? [
          new Paragraph(""),
        ],
    });
  }
  if (usesEvenPageVariants(config)) {
    footers.even = new Footer({
      children:
        buildHeaderFooterParagraphs(config, "footer", "even", theme) ?? [
          new Paragraph(""),
        ],
    });
  }
  return footers;
}

export function normalizeInheritedHeaderFooter(
  config: HeaderFooterConfig,
): HeaderFooterConfig {
  return normalizeInheritedHeaderFooterRule(config);
}

export function mergeHeaderFooter(
  baseConfig: HeaderFooterConfig,
  overrideConfig: HeaderFooterConfig,
): HeaderFooterConfig {
  return mergeHeaderFooterRule(baseConfig, overrideConfig);
}

export function hasHeaderFooterOverride(config: HeaderFooterConfig): boolean {
  return hasHeaderFooterOverrideRule(config);
}

export function usesFirstPageVariants(config: HeaderFooterConfig): boolean {
  return usesFirstPageVariantsRule(config);
}

export function usesEvenPageVariants(config: HeaderFooterConfig): boolean {
  return usesEvenPageVariantsRule(config);
}
