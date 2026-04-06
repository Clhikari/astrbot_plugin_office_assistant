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
  stringValue,
} from "./utils";

export function renderHeroBanner(block: Block, theme: ThemeConfig): Table {
  const layout = asObject(block.layout);
  const style = asObject(block.style);
  const backgroundColor = stringValue(block.theme_color) || theme.accent;
  const titleColor = stringValue(block.text_color) || "FFFFFF";
  const subtitleColor = stringValue(block.subtitle_color) || "EAF1F8";
  const padding = resolveBoxPadding(layout, 18);
  const titleFontScale = numberValue(style.font_scale) ?? 1;
  const titleSize = Math.max(theme.titleSize + 4, theme.headingSize + 8);
  const subtitleSize = Math.max(theme.bodySize + 0.5, 11);
  const widthPercent = booleanValue(block.full_width) === false ? 88 : 100;
  const alignment =
    mapAlignment(stringValue(style.align) || "left") ?? AlignmentType.LEFT;
  const minHeight = numberValue(block.min_height_pt);

  const children = [
    new Paragraph({
      alignment,
      spacing: { after: point(5) },
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
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        cantSplit: true,
        height: minHeight !== undefined
          ? {
              value: point(minHeight) ?? 0,
              rule: HeightRule.ATLEAST,
            }
          : undefined,
        children: [
          new TableCell({
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
