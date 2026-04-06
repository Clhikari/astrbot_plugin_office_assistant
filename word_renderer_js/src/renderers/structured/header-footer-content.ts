import {
  AlignmentType,
  BorderStyle,
  Paragraph,
  SimpleField,
  TabStopPosition,
  TabStopType,
  TextRun,
} from "docx";

import { JsonObject } from "../../core/payload";
import { DEFAULT_DIVIDER_COLOR } from "./constants";
import { HeaderFooterConfig, ThemeConfig } from "./types";
import { booleanValue, mapAlignment, normalizeHexColor, stringValue } from "./utils";
import { buildFontAttributes } from "./inline";
import {
  containsPagePlaceholder,
  resolveHeaderFooterLeft,
  resolveHeaderFooterRight,
  resolveShowPageNumberSetting,
  suppressPagePlaceholder,
  usesSplitLayout,
} from "./header-footer-rules";

export function buildHeaderFooterParagraphs(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
  variant: "default" | "first" | "even",
  theme: ThemeConfig,
): Paragraph[] | null {
  const splitLayout = usesSplitLayout(config, kind);
  let left = resolveHeaderFooterLeft(config, kind, variant);
  let right = resolveHeaderFooterRight(config, kind);
  const showPageNumberSetting = resolveShowPageNumberSetting(config, variant);
  const showPageNumber = showPageNumberSetting === true;

  if (showPageNumberSetting === false) {
    left = suppressPagePlaceholder(left);
    right = suppressPagePlaceholder(right);
  }

  if (
    showPageNumber &&
    !containsPagePlaceholder(left) &&
    !containsPagePlaceholder(right) &&
    kind === "footer"
  ) {
    if (splitLayout || left.trim()) {
      right = right.trim() ? `${right} {PAGE}` : "{PAGE}";
    } else {
      right = "{PAGE}";
    }
  }

  if (!(left.trim() || right.trim() || showPageNumber)) {
    return null;
  }

  if (right.trim()) {
    return [
      new Paragraph({
        border: buildHeaderFooterDivider(config, kind),
        tabStops: [
          {
            type: TabStopType.RIGHT,
            position: TabStopPosition.MAX,
          },
        ],
        children: buildSplitHeaderFooterRuns(left, right, theme),
      }),
    ];
  }

  if (kind === "footer" && showPageNumber) {
    if (left.trim()) {
      return [
        new Paragraph({
          border: buildHeaderFooterDivider(config, kind),
          children: buildInlineRuns(left, theme),
        }),
        new Paragraph({
          alignment: resolvePageNumberAlignment(config),
          children: [new SimpleField("PAGE")],
        }),
      ];
    }
    return [
      new Paragraph({
        border: buildHeaderFooterDivider(config, kind),
        alignment: resolvePageNumberAlignment(config),
        children: [new SimpleField("PAGE")],
      }),
    ];
  }

  return [
    new Paragraph({
      border: buildHeaderFooterDivider(config, kind),
      children: buildInlineRuns(left, theme),
    }),
  ];
}

function buildSplitHeaderFooterRuns(
  left: string,
  right: string,
  theme: ThemeConfig,
): Array<TextRun | SimpleField> {
  const children: Array<TextRun | SimpleField> = [];
  if (left.trim()) {
    children.push(...buildInlineRuns(left, theme));
  }
  if (left.trim() || right.trim()) {
    children.push(
      new TextRun({
        text: "\t",
        font: buildFontAttributes(theme.fontName),
      }),
    );
  }
  if (right.trim()) {
    children.push(...buildInlineRuns(right, theme));
  }
  return children.length > 0
    ? children
    : [new TextRun({ text: "", font: buildFontAttributes(theme.fontName) })];
}

function buildInlineRuns(
  text: string,
  theme: ThemeConfig,
): Array<TextRun | SimpleField> {
  if (!text.includes("{PAGE}")) {
    return [
      new TextRun({
        text,
        font: buildFontAttributes(theme.fontName),
      }),
    ];
  }
  const children: Array<TextRun | SimpleField> = [];
  const parts = text.split("{PAGE}");
  parts.forEach((part, index) => {
    if (part) {
      children.push(
        new TextRun({
          text: part,
          font: buildFontAttributes(theme.fontName),
        }),
      );
    }
    if (index < parts.length - 1) {
      children.push(new SimpleField("PAGE"));
    }
  });
  return children.length > 0 ? children : [new SimpleField("PAGE")];
}

function resolvePageNumberAlignment(config: HeaderFooterConfig) {
  return mapAlignment(stringValue(config.page_number_align)) ?? AlignmentType.RIGHT;
}

function buildHeaderFooterDivider(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
): JsonObject | undefined {
  const enabled =
    kind === "header"
      ? booleanValue(config.header_border_bottom) === true
      : booleanValue(config.footer_border_top) === true;
  if (!enabled) {
    return undefined;
  }
  const color =
    normalizeHexColor(
      stringValue(
        config[
          kind === "header" ? "header_border_color" : "footer_border_color"
        ],
      ),
    ) || DEFAULT_DIVIDER_COLOR;

  return kind === "header"
    ? {
        bottom: {
          color,
          style: BorderStyle.SINGLE,
          size: 4,
        },
      }
    : {
        top: {
          color,
          style: BorderStyle.SINGLE,
          size: 4,
        },
      };
}
