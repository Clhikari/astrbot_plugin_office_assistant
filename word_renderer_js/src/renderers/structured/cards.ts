import {
  BorderStyle,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

import { JsonObject } from "../../core/payload";
import { Block, ThemeConfig } from "./types";
import { buildFontAttributes, buildRuns, normalizeInlineItem } from "./inline";
import {
  arrayValue,
  asObject,
  borderSize,
  halfPoint,
  numberValue,
  resolveBoxPadding,
  stringValue,
} from "./utils";

export function renderAccentBox(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Table {
  const documentStyle = asObject(metadata.document_style);
  const accentColor = stringValue(block.accent_color) || theme.accent;
  const fillColor = stringValue(block.fill_color) || theme.summaryFill;
  const titleColor = stringValue(block.title_color) || accentColor;
  const borderColor = stringValue(block.border_color) || "D9E1E8";
  const bodyFontSize = numberValue(documentStyle.body_font_size) || theme.bodySize;
  const padding = resolveBoxPadding(
    asObject(block.layout),
    numberValue(block.padding_pt) ?? 14,
  );
  const titleFontScale = numberValue(block.title_font_scale) ?? 1.08;
  const bodyFontScale = numberValue(block.body_font_scale) ?? 1;
  const content: Paragraph[] = [];

  if (stringValue(block.title).trim()) {
    content.push(
      new Paragraph({
        spacing: { after: 90 },
        children: [
          new TextRun({
            text: stringValue(block.title),
            bold: true,
            color: titleColor,
            size: halfPoint(Math.max(theme.bodySize + 1.5, 12) * titleFontScale),
            font: buildFontAttributes(theme.headingFontName),
          }),
        ],
      }),
    );
  }

  const items = arrayValue(block.items);
  if (items.length > 0) {
    for (const item of items) {
      const normalized = normalizeInlineItem(item, theme, {
        fontSize: bodyFontSize,
        fontScale: bodyFontScale,
        fontName: theme.fontName,
        codeFontName: theme.codeFontName,
      });
      content.push(
        new Paragraph({
          spacing: { after: 60 },
          children: normalized.runs,
        }),
      );
    }
  } else if (arrayValue(block.runs).length > 0 || stringValue(block.text).trim()) {
    content.push(
      new Paragraph({
        spacing: { after: 20 },
        children: buildRuns(block, theme, {
          fontSize: bodyFontSize,
          fontScale: bodyFontScale,
          fontName: theme.fontName,
          codeFontName: theme.codeFontName,
        }),
      }),
    );
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            children: content.length > 0 ? content : [new Paragraph("")],
            margins: padding,
            shading: {
              fill: fillColor,
              color: "auto",
              type: ShadingType.CLEAR,
            },
            borders: {
              left: {
                color: accentColor,
                style: BorderStyle.SINGLE,
                size: borderSize(numberValue(block.accent_border_width_pt), 2.25),
              },
              top: {
                color: borderColor,
                style: BorderStyle.SINGLE,
                size: borderSize(numberValue(block.border_width_pt), 0.5),
              },
              right: {
                color: borderColor,
                style: BorderStyle.SINGLE,
                size: borderSize(numberValue(block.border_width_pt), 0.5),
              },
              bottom: {
                color: borderColor,
                style: BorderStyle.SINGLE,
                size: borderSize(numberValue(block.border_width_pt), 0.5),
              },
            },
          }),
        ],
      }),
    ],
  });
}

export function renderMetricCards(
  block: Block,
  _metadata: JsonObject,
  theme: ThemeConfig,
): Table {
  const accentColor = stringValue(block.accent_color) || theme.accent;
  const fillColor = stringValue(block.fill_color) || "F8FAFC";
  const labelColor = stringValue(block.label_color) || "666666";
  const borderColor = stringValue(block.border_color) || "E5E7EB";
  const dividerColor = stringValue(block.divider_color) || borderColor;
  const padding = resolveBoxPadding(
    asObject(block.layout),
    numberValue(block.padding_pt) ?? 12,
  );
  const labelFontScale = numberValue(block.label_font_scale) ?? 0.92;
  const valueFontScale = numberValue(block.value_font_scale) ?? 1.75;
  const deltaFontScale = numberValue(block.delta_font_scale) ?? 0.92;
  const noteFontScale = numberValue(block.note_font_scale) ?? 0.88;
  const labelFontSize = theme.bodySize * labelFontScale;
  const valueFontSize = theme.bodySize * valueFontScale;
  const deltaFontSize = theme.bodySize * deltaFontScale;
  const noteFontSize = theme.bodySize * noteFontScale;

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: {
        style: BorderStyle.SINGLE,
        color: borderColor,
        size: borderSize(numberValue(block.border_width_pt), 0.5),
      },
      bottom: {
        style: BorderStyle.SINGLE,
        color: borderColor,
        size: borderSize(numberValue(block.border_width_pt), 0.5),
      },
      left: {
        style: BorderStyle.SINGLE,
        color: borderColor,
        size: borderSize(numberValue(block.border_width_pt), 0.5),
      },
      right: {
        style: BorderStyle.SINGLE,
        color: borderColor,
        size: borderSize(numberValue(block.border_width_pt), 0.5),
      },
      insideHorizontal: {
        style: BorderStyle.SINGLE,
        color: dividerColor,
        size: borderSize(numberValue(block.divider_width_pt), 0.5),
      },
      insideVertical: {
        style: BorderStyle.SINGLE,
        color: dividerColor,
        size: borderSize(numberValue(block.divider_width_pt), 0.5),
      },
    },
    rows: [
      new TableRow({
        cantSplit: true,
        children: arrayValue(block.metrics).map((metric) => {
          const metricObject = asObject(metric);
          const metricValueFontScale =
            numberValue(metricObject.value_font_scale) ?? valueFontScale;
          const metricDeltaFontScale =
            numberValue(metricObject.delta_font_scale) ?? deltaFontScale;
          const paragraphs: Paragraph[] = [
            new Paragraph({
              children: [
                new TextRun({
                  text: stringValue(metricObject.label),
                  bold: true,
                  color: stringValue(metricObject.label_color) || labelColor,
                  size: halfPoint(labelFontSize),
                  font: buildFontAttributes(theme.fontName),
                }),
              ],
            }),
            new Paragraph({
              spacing: { before: 50, after: 45 },
              children: [
                new TextRun({
                  text: stringValue(metricObject.value),
                  bold: true,
                  color: stringValue(metricObject.value_color) || accentColor,
                  size: halfPoint(theme.bodySize * metricValueFontScale),
                  font: buildFontAttributes(theme.headingFontName),
                }),
              ],
            }),
          ];

          if (stringValue(metricObject.delta).trim()) {
            paragraphs.push(
              new Paragraph({
                spacing: { after: 35 },
                children: [
                  new TextRun({
                    text: stringValue(metricObject.delta),
                    color: stringValue(metricObject.delta_color) || "15803D",
                    size: halfPoint(theme.bodySize * metricDeltaFontScale),
                    font: buildFontAttributes(theme.fontName),
                  }),
                ],
              }),
            );
          }
          if (stringValue(metricObject.note).trim()) {
            paragraphs.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: stringValue(metricObject.note),
                    color: stringValue(metricObject.note_color) || "666666",
                    size: halfPoint(noteFontSize),
                    font: buildFontAttributes(theme.fontName),
                  }),
                ],
              }),
            );
          }

          return new TableCell({
            children: paragraphs,
            margins: padding,
            shading: {
              fill: stringValue(metricObject.fill_color) || fillColor,
              color: "auto",
              type: ShadingType.CLEAR,
            },
          });
        }),
      }),
    ],
  });
}
