import {
  AlignmentType,
  BorderStyle,
  PageBreak,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

import { JsonObject } from "../../../core/payload";
import { renderHeroBanner } from "../hero-banner";
import { buildFontAttributes } from "../inline";
import { Block, FileChild, ThemeConfig } from "../types";
import {
  arrayValue,
  asObject,
  borderSize,
  booleanValue,
  halfPoint,
  point,
  resolveBoxPadding,
  stringValue,
} from "../utils";

const COVER_TABLE_WIDTH_DXA = 9360;
const SUMMARY_STRIP_WIDTH_DXA = 160;
const FOOTER_STRIP_WIDTH_DXA = 160;
const FOOTER_CONTENT_WIDTH_DXA = 9200;
const THREE_METRIC_CELL_WIDTH_DXA = 3120;

export function renderBusinessReviewCover(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  const data = asObject(block.data);
  const children: FileChild[] = [
    renderHeroBanner(
      {
        type: "hero_banner",
        title: stringValue(data.title),
        subtitle: stringValue(data.subtitle),
        theme_color: theme.accent,
        text_color: "FFFFFF",
        subtitle_color: "D7E4F2",
        min_height_pt: 80,
        exact_height: true,
        full_width: false,
        width_dxa: COVER_TABLE_WIDTH_DXA,
        table_align: "center",
        title_spacing_after_pt: 0,
        style: {
          font_scale: 1.0,
        },
        layout: {
          padding_top_pt: 16,
          padding_right_pt: 20,
          padding_bottom_pt: 0,
          padding_left_pt: 20,
        },
      } as Block,
      theme,
    ),
    renderSummaryStripBox(data, theme),
    renderCompactMetricBand(data, theme),
  ];

  const footerNote = stringValue(data.footer_note).trim();
  if (footerNote) {
    children.push(...buildBusinessReviewFooterNote(footerNote, theme));
  }

  if (booleanValue(data.auto_page_break) === true) {
    children.push(
      new Paragraph({
        spacing: { before: point(4) },
        children: [new PageBreak()],
      }),
    );
  }

  return children;
}

export function buildBusinessReviewFooterNote(
  footerNote: string,
  theme: ThemeConfig,
): FileChild[] {
  if (!footerNote.trim()) {
    return [];
  }

  return [
    new Table({
      width: { size: COVER_TABLE_WIDTH_DXA, type: WidthType.DXA },
      columnWidths: [FOOTER_STRIP_WIDTH_DXA, FOOTER_CONTENT_WIDTH_DXA],
      alignment: AlignmentType.CENTER,
      layout: TableLayoutType.FIXED,
      rows: [
        new TableRow({
          cantSplit: true,
          children: [
            new TableCell({
              width: { size: FOOTER_STRIP_WIDTH_DXA, type: WidthType.DXA },
              children: [new Paragraph("")],
              shading: {
                fill: theme.accent,
                color: "auto",
                type: ShadingType.CLEAR,
              },
              borders: noCellBorders(),
            }),
            new TableCell({
              width: { size: FOOTER_CONTENT_WIDTH_DXA, type: WidthType.DXA },
              margins: resolveBoxPadding(
                {
                  padding_top_pt: 4,
                  padding_right_pt: 9,
                  padding_bottom_pt: 4,
                  padding_left_pt: 9,
                },
                8,
              ),
              shading: {
                fill: "EEF3FB",
                color: "auto",
                type: ShadingType.CLEAR,
              },
              borders: noCellBorders(),
              children: [
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  spacing: { before: 0, after: 0 },
                  children: [
                    new TextRun({
                      text: footerNote,
                      italics: true,
                      color: "595959",
                      size: halfPoint(9),
                      font: buildFontAttributes(theme.fontName),
                    }),
                  ],
                }),
              ],
            }),
          ],
        }),
      ],
    }),
  ];
}

function renderSummaryStripBox(
  data: JsonObject,
  theme: ThemeConfig,
): Table {
  const stripWidthPt = SUMMARY_STRIP_WIDTH_DXA / 20;
  return new Table({
    width: { size: COVER_TABLE_WIDTH_DXA, type: WidthType.DXA },
    alignment: AlignmentType.CENTER,
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            width: { size: COVER_TABLE_WIDTH_DXA, type: WidthType.DXA },
            margins: resolveBoxPadding(
              {
                padding_top_pt: 5,
                padding_right_pt: 9,
                padding_bottom_pt: 5,
                padding_left_pt: 9,
              },
              10,
            ),
            shading: {
              fill: "EEF9F5",
              color: "auto",
              type: ShadingType.CLEAR,
            },
            borders: {
              left: {
                color: "0F6E56",
                style: BorderStyle.SINGLE,
                size: borderSize(stripWidthPt, stripWidthPt),
              },
              top: { style: BorderStyle.NONE, size: borderSize(undefined, 0.25) },
              bottom: { style: BorderStyle.NONE, size: borderSize(undefined, 0.25) },
              right: { style: BorderStyle.NONE, size: borderSize(undefined, 0.25) },
            },
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: stringValue(data.summary_title) || "核心摘要",
                    bold: true,
                    color: "595959",
                    size: halfPoint(9.5),
                    font: buildFontAttributes(theme.fontName),
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: stringValue(data.summary_text),
                    size: halfPoint(10),
                    font: buildFontAttributes(theme.fontName),
                  }),
                ],
              }),
            ],
          }),
        ],
      }),
    ],
  });
}

function renderCompactMetricBand(
  data: JsonObject,
  theme: ThemeConfig,
): Table {
  const metrics = arrayValue(data.metrics).slice(0, 4).map((metric) => asObject(metric));
  const displayMetrics = padMetricSlots(metrics);
  const metricWidths = resolveMetricColumnWidths(displayMetrics.length);
  return new Table({
    width: { size: COVER_TABLE_WIDTH_DXA, type: WidthType.DXA },
    columnWidths: metricWidths,
    alignment: AlignmentType.CENTER,
    layout: TableLayoutType.FIXED,
    borders: {
      top: { color: "BFBFBF", style: BorderStyle.SINGLE, size: 1 },
      bottom: { color: "BFBFBF", style: BorderStyle.SINGLE, size: 1 },
      left: { color: "BFBFBF", style: BorderStyle.SINGLE, size: 1 },
      right: { color: "BFBFBF", style: BorderStyle.SINGLE, size: 1 },
      insideHorizontal: { color: "BFBFBF", style: BorderStyle.SINGLE, size: 1 },
      insideVertical: { color: "BFBFBF", style: BorderStyle.SINGLE, size: 1 },
    },
    rows: [
      new TableRow({
        cantSplit: true,
        children: displayMetrics.map((metricObject, index) => {
          const deltaText = stringValue(metricObject.delta);
          const deltaColor =
            stringValue(metricObject.delta_color) ||
            (deltaText.includes("↓") ? "843C0C" : "375623");
          return new TableCell({
            width: { size: metricWidths[index], type: WidthType.DXA },
            shading: {
              fill: "F2F7FC",
              color: "auto",
              type: ShadingType.CLEAR,
            },
            margins: resolveBoxPadding(
              {
                padding_top_pt: 6,
                padding_right_pt: 7,
                padding_bottom_pt: 6,
                padding_left_pt: 7,
              },
              8,
            ),
            children: [
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: stringValue(metricObject.label),
                    color: "595959",
                    size: halfPoint(9),
                    font: buildFontAttributes(theme.fontName),
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: stringValue(metricObject.value),
                    bold: true,
                    color: theme.accent,
                    size: halfPoint(17),
                    font: buildFontAttributes(theme.headingFontName),
                  }),
                ],
              }),
              new Paragraph({
                spacing: { before: 0, after: 0 },
                children: [
                  new TextRun({
                    text: deltaText,
                    color: deltaColor,
                    size: halfPoint(9),
                    font: buildFontAttributes(theme.fontName),
                  }),
                ],
              }),
            ],
          });
        }),
      }),
    ],
  });
}

function noCellBorders() {
  return {
    top: { style: BorderStyle.NONE, size: borderSize(undefined, 0.25) },
    bottom: { style: BorderStyle.NONE, size: borderSize(undefined, 0.25) },
    left: { style: BorderStyle.NONE, size: borderSize(undefined, 0.25) },
    right: { style: BorderStyle.NONE, size: borderSize(undefined, 0.25) },
  };
}

function resolveMetricColumnWidths(metricCount: number): number[] {
  if (metricCount <= 0) {
    return [];
  }
  if (metricCount === 3) {
    return [
      THREE_METRIC_CELL_WIDTH_DXA,
      THREE_METRIC_CELL_WIDTH_DXA,
      THREE_METRIC_CELL_WIDTH_DXA,
    ];
  }
  const baseWidth = Math.floor(COVER_TABLE_WIDTH_DXA / metricCount);
  return Array.from({ length: metricCount }, (_, index) =>
    index === metricCount - 1
      ? COVER_TABLE_WIDTH_DXA - baseWidth * (metricCount - 1)
      : baseWidth,
  );
}

function padMetricSlots(metrics: JsonObject[]): JsonObject[] {
  if (metrics.length >= 3) {
    return metrics;
  }
  return [
    ...metrics,
    ...Array.from({ length: 3 - metrics.length }, () => ({} as JsonObject)),
  ];
}
