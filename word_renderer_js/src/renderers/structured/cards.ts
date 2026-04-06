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
import { arrayValue, asObject, stringValue } from "./utils";
import { normalizeInlineItem } from "./inline";

export function renderAccentBox(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
  renderParagraph: (
    block: Block,
    metadata: JsonObject,
    theme: ThemeConfig,
  ) => Array<Paragraph | Table>,
): Table {
  const accentColor = stringValue(block.accent_color) || theme.accent;
  const fillColor = stringValue(block.fill_color) || theme.summaryFill;
  const titleColor = stringValue(block.title_color) || accentColor;
  const content: Paragraph[] = [];

  if (stringValue(block.title).trim()) {
    content.push(
      new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({
            text: stringValue(block.title),
            bold: true,
            color: titleColor,
          }),
        ],
      }),
    );
  }

  const items = arrayValue(block.items);
  if (items.length > 0) {
    for (const item of items) {
      const normalized = normalizeInlineItem(item, theme);
      content.push(
        new Paragraph({
          spacing: { after: 40 },
          children: normalized.runs,
        }),
      );
    }
  } else if (stringValue(block.text).trim()) {
    content.push(
      ...renderParagraph(block, metadata, theme).filter(
        (child): child is Paragraph => child instanceof Paragraph,
      ),
    );
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            children: content.length > 0 ? content : [new Paragraph("")],
            shading: {
              fill: fillColor,
              color: "auto",
              type: ShadingType.CLEAR,
            },
            borders: {
              left: {
                color: accentColor,
                style: BorderStyle.SINGLE,
                size: 18,
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

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      bottom: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      left: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      right: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      insideHorizontal: {
        style: BorderStyle.SINGLE,
        color: "E5E7EB",
        size: 4,
      },
      insideVertical: {
        style: BorderStyle.SINGLE,
        color: "E5E7EB",
        size: 4,
      },
    },
    rows: [
      new TableRow({
        children: arrayValue(block.metrics).map((metric) => {
          const metricObject = asObject(metric);
          const paragraphs: Paragraph[] = [
            new Paragraph({
              children: [
                new TextRun({
                  text: stringValue(metricObject.label),
                  bold: true,
                  color: labelColor,
                }),
              ],
            }),
            new Paragraph({
              spacing: { before: 40, after: 40 },
              children: [
                new TextRun({
                  text: stringValue(metricObject.value),
                  bold: true,
                  color: stringValue(metricObject.value_color) || accentColor,
                  size: 28,
                }),
              ],
            }),
          ];

          if (stringValue(metricObject.delta).trim()) {
            paragraphs.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: stringValue(metricObject.delta),
                    color: stringValue(metricObject.delta_color) || "15803D",
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
                    color: "666666",
                  }),
                ],
              }),
            );
          }

          return new TableCell({
            children: paragraphs,
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
