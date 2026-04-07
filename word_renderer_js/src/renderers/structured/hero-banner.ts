import {
  AlignmentType,
  HeightRule,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

import { Block, ThemeConfig } from "./types";
import { buildFontAttributes } from "./inline";
import {
  asObject,
  booleanValue,
  halfPoint,
  mapAlignment,
  numberValue,
  point,
  resolveBoxPadding,
  resolveBoxPaddingEdges,
  stringValue,
} from "./utils";

const BUSINESS_REPORT_CONTENT_WIDTH_DXA = 9360;

export function renderHeroBanner(block: Block, theme: ThemeConfig): Table {
  const isBusinessReport = theme.themeName === "business_report";
  const layout = asObject(block.layout);
  const style = asObject(block.style);
  const backgroundColor = stringValue(block.theme_color) || theme.accent;
  const titleColor = stringValue(block.text_color) || "FFFFFF";
  const subtitleColor = stringValue(block.subtitle_color) || "EAF1F8";
  const padding = isBusinessReport
    ? resolveBoxPaddingEdges(layout, {
        top: 10,
        right: 18,
        bottom: 10,
        left: 18,
      })
    : resolveBoxPadding(layout, 18);
  const titleFontScale = numberValue(style.font_scale) ?? 1;
  const titleSize = isBusinessReport
    ? Math.max(theme.titleSize + 6, 26)
    : Math.max(theme.titleSize + 4, theme.headingSize + 8);
  const subtitleSize = isBusinessReport
    ? Math.max(theme.bodySize, 10.5)
    : Math.max(theme.bodySize + 0.5, 11);
  const widthDxa =
    numberValue(block.width_dxa) ??
    (isBusinessReport && numberValue(block.width_pct) === undefined
      ? BUSINESS_REPORT_CONTENT_WIDTH_DXA
      : undefined);
  const widthPercent =
    numberValue(block.width_pct) ??
    (booleanValue(block.full_width) === false ? 88 : 100);
  const alignment =
    mapAlignment(stringValue(style.align) || "left") ?? AlignmentType.LEFT;
  const tableAlignment =
    mapAlignment(stringValue(block.table_align)) ??
    (widthDxa !== undefined || widthPercent < 100 ? AlignmentType.CENTER : undefined);
  const minHeight =
    numberValue(block.min_height_pt) ?? (isBusinessReport ? 56 : undefined);
  const exactHeight = booleanValue(block.exact_height) === true;
  const titleSpacingAfter = numberValue(block.title_spacing_after_pt) ?? (isBusinessReport ? 2 : 5);

  const children = [
    new Paragraph({
      alignment,
      spacing: { after: point(titleSpacingAfter) },
      children: [
        new TextRun({
          text: stringValue(block.title),
          bold: true,
          color: titleColor,
          size: halfPoint(titleSize * titleFontScale),
          font: buildFontAttributes(theme.headingFontName),
        }),
      ],
    }),
  ];

  const subtitle = stringValue(block.subtitle);
  if (subtitle.trim()) {
    children.push(
      new Paragraph({
        alignment,
        children: [
          new TextRun({
            text: subtitle,
            color: subtitleColor,
            size: halfPoint(subtitleSize),
            font: buildFontAttributes(theme.fontName),
          }),
        ],
      }),
    );
  }

  return new Table({
    width:
      widthDxa !== undefined
        ? { size: widthDxa, type: WidthType.DXA }
        : { size: widthPercent, type: WidthType.PERCENTAGE },
    columnWidths: widthDxa !== undefined ? [widthDxa] : undefined,
    alignment: tableAlignment,
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        cantSplit: true,
        height: minHeight !== undefined
          ? {
              value: point(minHeight) ?? 0,
              rule: exactHeight ? HeightRule.EXACT : HeightRule.ATLEAST,
            }
          : undefined,
        children: [
          new TableCell({
            width:
              widthDxa !== undefined
                ? { size: widthDxa, type: WidthType.DXA }
                : undefined,
            children,
            margins: padding,
            shading: {
              fill: backgroundColor,
              color: "auto",
              type: ShadingType.CLEAR,
            },
          }),
        ],
      }),
    ],
  });
}
