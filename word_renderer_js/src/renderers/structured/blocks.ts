import {
  AlignmentType,
  BorderStyle,
  LineRuleType,
  PageBreak,
  Paragraph,
  SimpleField,
  TextRun,
} from "docx";

import { JsonObject } from "../../core/payload";
import { DEFAULT_DIVIDER_COLOR } from "./constants";
import { Block, FileChild, ThemeConfig } from "./types";
import {
  arrayValue,
  asObject,
  booleanValue,
  halfPoint,
  mapAlignment,
  mapHeadingLevel,
  numberValue,
  point,
  resolveBold,
  resolveTextColor,
  stringValue,
} from "./utils";
import {
  buildFontAttributes,
  buildRuns,
  mergeLayoutDefaults,
  mergeStyleDefaults,
  normalizeInlineItem,
  paragraphPlainText,
} from "./inline";

export function renderDocumentTitle(
  metadata: JsonObject,
  theme: ThemeConfig,
): Paragraph | null {
  const title = stringValue(metadata.title);
  if (!title.trim()) {
    return null;
  }
  const documentStyle = asObject(metadata.document_style);
  return new Paragraph({
    alignment:
      mapAlignment(stringValue(documentStyle.title_align) || theme.titleAlign) ??
      AlignmentType.LEFT,
    spacing: {
      after: point(theme.titleSpacingAfter),
    },
    children: [
      new TextRun({
        text: title,
        bold: true,
        color: stringValue(documentStyle.heading_color) || "000000",
        size: halfPoint(theme.titleSize),
        font: buildFontAttributes(theme.headingFontName),
      }),
    ],
  });
}

export function renderHeading(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Paragraph {
  const documentStyle = asObject(metadata.document_style);
  const level = numberValue(block.level) ?? 1;
  const style = asObject(block.style);
  const layout = asObject(block.layout);
  const color =
    stringValue(block.color) ||
    stringValue(documentStyle[`heading_level_${level}_color`]) ||
    stringValue(documentStyle.heading_color) ||
    theme.accent;
  const fontScale = numberValue(style.font_scale) ?? 1;
  const baseSize =
    level <= 1 ? theme.headingSize : Math.max(theme.bodySize + 1, 11.5);
  const border =
    booleanValue(block.bottom_border) === true
      ? {
          bottom: {
            color:
              stringValue(block.bottom_border_color) ||
              stringValue(documentStyle.heading_bottom_border_color) ||
              DEFAULT_DIVIDER_COLOR,
            style: BorderStyle.SINGLE,
            size: Math.max(
              4,
              Math.round(
                (numberValue(block.bottom_border_size_pt) ||
                  numberValue(documentStyle.heading_bottom_border_size_pt) ||
                  0.5) * 8,
              ),
            ),
          },
        }
      : undefined;

  return new Paragraph({
    heading: mapHeadingLevel(level),
    alignment: mapAlignment(stringValue(style.align)) ?? AlignmentType.LEFT,
    border,
    spacing: {
      before: point(numberValue(layout.spacing_before) ?? theme.headingSpaceBefore),
      after: point(numberValue(layout.spacing_after) ?? theme.headingSpaceAfter),
    },
    children: [
      new TextRun({
        text: stringValue(block.text),
        bold: true,
        color,
        size: halfPoint(baseSize * fontScale),
        font: buildFontAttributes(theme.headingFontName),
      }),
    ],
  });
}

export function renderParagraph(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  const documentStyle = asObject(metadata.document_style);
  const style = asObject(block.style);
  const layout = asObject(block.layout);
  const variant = stringValue(block.variant) || "body";
  const bodyFontSize = numberValue(documentStyle.body_font_size) || theme.bodySize;
  const bodyLineSpacing =
    numberValue(documentStyle.body_line_spacing) || theme.bodyLineSpacing;

  if (variant === "summary_box" || variant === "key_takeaway") {
    return renderSummaryCardLikeGroup(
      {
        title:
          stringValue(block.title) ||
          (variant === "summary_box" ? "Summary" : "Key Takeaway"),
        items: [paragraphPlainText(block)],
        variant: "summary",
        style,
        layout,
      },
      metadata,
      theme,
    );
  }

  return [
    new Paragraph({
      children: buildRuns(block, theme, {
        fontSize: bodyFontSize,
        emphasis: stringValue(style.emphasis),
        fontScale: numberValue(style.font_scale),
        fontName: theme.fontName,
        codeFontName: theme.codeFontName,
      }),
      spacing: {
        before: point(numberValue(layout.spacing_before) ?? 0),
        after: point(
          numberValue(layout.spacing_after) ??
            numberValue(documentStyle.paragraph_space_after) ??
            theme.bodySpaceAfter,
        ),
        line: point(bodyFontSize * bodyLineSpacing),
        lineRule: LineRuleType.AUTO,
      },
      indent: {
        firstLine: point(theme.bodyIndent),
      },
      alignment: mapAlignment(stringValue(style.align)) ?? AlignmentType.LEFT,
    }),
  ];
}

export function renderList(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Paragraph[] {
  const documentStyle = asObject(metadata.document_style);
  const style = asObject(block.style);
  const ordered = booleanValue(block.ordered) === true;
  const layout = asObject(block.layout);
  const bodyFontSize = numberValue(documentStyle.body_font_size) || theme.bodySize;

  return arrayValue(block.items).map((item, index) => {
    const normalized = normalizeInlineItem(item, theme, {
      fontSize: bodyFontSize,
      emphasis: stringValue(style.emphasis),
      fontScale: numberValue(style.font_scale),
      fontName: theme.fontName,
      codeFontName: theme.codeFontName,
    });
    const marker = ordered ? `${index + 1}. ` : "• ";
    return new Paragraph({
      children: [
        new TextRun({
          text: marker,
          bold: resolveBold(false, stringValue(style.emphasis)),
          color: resolveTextColor(theme, stringValue(style.emphasis)),
          size: halfPoint(bodyFontSize * (numberValue(style.font_scale) ?? 1)),
          font: buildFontAttributes(theme.fontName),
        }),
        ...normalized.runs,
      ],
      spacing: {
        before: point(numberValue(layout.spacing_before) ?? 0),
        after: point(
          numberValue(layout.spacing_after) ??
            numberValue(documentStyle.list_space_after) ??
            theme.listSpaceAfter,
        ),
      },
      indent: {
        left: point(Math.max(theme.bodyIndent - 6, 12)),
        firstLine: 0,
      },
      alignment: mapAlignment(stringValue(style.align)) ?? AlignmentType.LEFT,
    });
  });
}

export function renderSummaryCard(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  return renderSummaryCardLikeGroup(
    {
      title: stringValue(block.title),
      items: arrayValue(block.items),
      variant: stringValue(block.variant) || "summary",
      style: asObject(block.style),
      layout: asObject(block.layout),
    },
    metadata,
    theme,
  );
}

export function renderToc(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  const children: FileChild[] = [];
  const title = stringValue(block.title) || "Contents";
  if (booleanValue(block.start_on_new_page) === true) {
    children.push(
      new Paragraph({
        children: [new PageBreak()],
      }),
    );
  }
  children.push(
    new Paragraph({
      spacing: {
        after: point(
          numberValue(asObject(block.layout).spacing_after) ??
            theme.headingSpaceAfter,
        ),
      },
      children: [
        new TextRun({
          text: title,
          bold: true,
          color:
            stringValue(asObject(metadata.document_style).heading_color) ||
            theme.accent,
          size: halfPoint(theme.headingSize),
          font: buildFontAttributes(theme.headingFontName),
        }),
      ],
    }),
  );
  children.push(
    new Paragraph({
      spacing: {
        after: point(theme.bodySpaceAfter),
      },
      children: [
        new SimpleField(
          `TOC \\o "${resolveTocHeadingStyleRange(block) ?? "1-3"}" \\h \\z \\u`,
        ),
      ],
    }),
  );
  return children;
}

function renderSummaryCardLikeGroup(
  block: {
    title: string;
    items: unknown[];
    variant: string;
    style: JsonObject;
    layout: JsonObject;
  },
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  const documentStyle = asObject(metadata.document_style);
  const summaryDefaults = asObject(documentStyle.summary_card_defaults);
  const normalizedVariant =
    block.variant === "conclusion" ? "conclusion" : "summary";
  const titleStyle = mergeStyleDefaults(block.style, {
    align: stringValue(summaryDefaults.title_align) || undefined,
    emphasis:
      stringValue(summaryDefaults.title_emphasis) ||
      (normalizedVariant === "conclusion" ? "subtle" : "strong"),
    fontScale: numberValue(summaryDefaults.title_font_scale) ?? 1.05,
  });
  const bodyStyle = mergeStyleDefaults(block.style, {
    emphasis: normalizedVariant === "conclusion" ? "subtle" : undefined,
  });
  const titleLayout = mergeLayoutDefaults(block.layout, {
    spacingBefore: numberValue(summaryDefaults.title_space_before) ?? 6,
    spacingAfter: numberValue(summaryDefaults.title_space_after) ?? 2,
  });
  const listLayout = mergeLayoutDefaults(block.layout, {
    spacingBefore: 0,
    spacingAfter:
      numberValue(summaryDefaults.list_space_after) ?? theme.listSpaceAfter,
  });

  return [
    ...renderParagraph(
      {
        type: "paragraph",
        text: block.title,
        style: titleStyle,
        layout: titleLayout,
      } as Block,
      metadata,
      theme,
    ),
    ...renderList(
      {
        type: "list",
        items: block.items,
        ordered: false,
        style: bodyStyle,
        layout: listLayout,
      } as Block,
      metadata,
      theme,
    ),
  ];
}

function resolveTocHeadingStyleRange(block: Block): string | undefined {
  const levels = numberValue(block.levels);
  if (levels === undefined) {
    return undefined;
  }
  const normalizedLevels = Math.max(1, Math.min(6, Math.trunc(levels)));
  return `1-${normalizedLevels}`;
}
