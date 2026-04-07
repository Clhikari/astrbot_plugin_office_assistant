import {
  AlignmentType,
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
  booleanValue,
  halfPoint,
  numberValue,
  resolveBoxPadding,
  resolveBoxPaddingEdges,
  stringValue,
} from "./utils";

const BUSINESS_REPORT_CONTENT_WIDTH_DXA = 9360;

export function renderAccentBox(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Table {
  const documentStyle = asObject(metadata.document_style);
  const isBusinessReport = stringValue(metadata.theme_name) === "business_report";
  const useStripLayout = isBusinessReport && booleanValue(block.strip_layout) !== false;
  const accentColor = stringValue(block.accent_color) || theme.accent;
  const fillColor = stringValue(block.fill_color) || theme.summaryFill;
  const titleColor =
    stringValue(block.title_color) || (useStripLayout ? "595959" : accentColor);
  const borderColor = stringValue(block.border_color) || "D9E1E8";
  const bodyFontSize = numberValue(documentStyle.body_font_size) || theme.bodySize;
  const uniformPadding = numberValue(block.padding_pt);
  const padding = useStripLayout
    ? resolveBoxPaddingEdges(asObject(block.layout), {
        top: uniformPadding ?? 5,
        right: uniformPadding ?? 9,
        bottom: uniformPadding ?? 5,
        left: uniformPadding ?? 9,
      })
    : resolveBoxPadding(asObject(block.layout), uniformPadding ?? 14);
  const titleFontScale = numberValue(block.title_font_scale) ?? 1.08;
  const bodyFontScale = numberValue(block.body_font_scale) ?? 1;
  const content: Paragraph[] = [];
  const titleBaseSize = useStripLayout ? 9.5 : Math.max(theme.bodySize + 1.5, 12);
  const bodyBaseSize = useStripLayout ? 10 : bodyFontSize;

  if (stringValue(block.title).trim()) {
    content.push(
      new Paragraph({
        spacing: { after: useStripLayout ? 0 : 90 },
        children: [
          new TextRun({
            text: stringValue(block.title),
            bold: true,
            color: titleColor,
            size: halfPoint(titleBaseSize * titleFontScale),
            font: buildFontAttributes(
              useStripLayout ? theme.fontName : theme.headingFontName,
            ),
          }),
        ],
      }),
    );
  }

  const items = arrayValue(block.items);
  if (items.length > 0) {
    for (const item of items) {
      const normalized = normalizeInlineItem(item, theme, {
        fontSize: bodyBaseSize,
        fontScale: bodyFontScale,
        fontName: theme.fontName,
        codeFontName: theme.codeFontName,
      });
      content.push(
        new Paragraph({
          spacing: { after: useStripLayout ? 20 : 60 },
          children: normalized.runs,
        }),
      );
    }
  } else if (arrayValue(block.runs).length > 0 || stringValue(block.text).trim()) {
    content.push(
        new Paragraph({
        spacing: { after: useStripLayout ? 0 : 20 },
        children: buildRuns(block, theme, {
          fontSize: bodyBaseSize,
          fontScale: bodyFontScale,
          fontName: theme.fontName,
          codeFontName: theme.codeFontName,
        }),
      }),
    );
  }

  if (useStripLayout) {
    const totalWidthDxa =
      numberValue(block.width_dxa) ?? BUSINESS_REPORT_CONTENT_WIDTH_DXA;
    const stripWidthPt =
      numberValue(block.accent_border_width_pt) ??
      ((numberValue(block.strip_width_dxa) ?? 160) / 20);

    return new Table({
      width: { size: totalWidthDxa, type: WidthType.DXA },
      columnWidths: [totalWidthDxa],
      alignment: AlignmentType.CENTER,
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: totalWidthDxa, type: WidthType.DXA },
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
                  size: borderSize(stripWidthPt, stripWidthPt),
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
  metadata: JsonObject,
  theme: ThemeConfig,
): Table {
  const isBusinessReport = stringValue(metadata.theme_name) === "business_report";
  const accentColor = stringValue(block.accent_color) || theme.accent;
  const fillColor =
    stringValue(block.fill_color) || (isBusinessReport ? "F2F7FC" : "F8FAFC");
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
  const metrics = arrayValue(block.metrics);
  const tableWidthDxa = isBusinessReport
    ? numberValue(block.width_dxa) ?? BUSINESS_REPORT_CONTENT_WIDTH_DXA
    : undefined;
  const metricWidths =
    isBusinessReport && metrics.length > 0
      ? resolveMetricCardWidths(tableWidthDxa ?? 9360, metrics.length)
      : [];

  return new Table({
    width:
      tableWidthDxa !== undefined
        ? { size: tableWidthDxa, type: WidthType.DXA }
        : { size: 100, type: WidthType.PERCENTAGE },
    alignment: tableWidthDxa !== undefined ? AlignmentType.CENTER : undefined,
    layout: TableLayoutType.FIXED,
    columnWidths: metricWidths.length > 0 ? metricWidths : undefined,
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
        children: metrics.map((metric, metricIndex) => {
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
            width:
              metricWidths.length > 0
                ? { size: metricWidths[metricIndex], type: WidthType.DXA }
                : undefined,
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

function resolveMetricCardWidths(totalWidthDxa: number, metricCount: number): number[] {
  if (metricCount <= 0) {
    return [];
  }
  if (metricCount === 3 && totalWidthDxa >= 9360) {
    return [3120, 3120, 3120];
  }
  const baseWidth = Math.floor(totalWidthDxa / metricCount);
  return Array.from({ length: metricCount }, (_, index) =>
    index === metricCount - 1
      ? totalWidthDxa - baseWidth * (metricCount - 1)
      : baseWidth,
  );
}
