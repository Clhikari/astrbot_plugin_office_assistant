import {
  HeightRule,
  LineRuleType,
  Paragraph,
  ShadingType,
  TableCell,
  TableRow,
  TextRun,
  VerticalMergeType,
  WidthType,
} from "docx";

import { RenderCliError } from "../../core/errors";
import { JsonObject } from "../../core/payload";
import { Block, TableCellValue, ThemeConfig } from "./types";
import { buildFontAttributes } from "./inline";
import {
  arrayValue,
  asObject,
  booleanValue,
  cmToTwip,
  halfPoint,
  mapAlignment,
  numberValue,
  stringValue,
} from "./utils";
import {
  resolveBandedRowFill,
  resolveFirstColumnBold,
  resolveTableBodyAlignment,
  resolveTableFontSize,
  resolveTableParagraphSpacing,
  resolveTableRowHeight,
} from "./table-style";

export function buildTableBodyRows(
  block: Block,
  columnCount: number,
  tableStyleName: string,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
  bodyAlignment: string,
  numericColumns: Set<number>,
  columnWidths: number[],
): TableRow[] {
  const rows = arrayValue(block.rows);
  const pendingRowSpans = new Array(Math.max(columnCount, 0)).fill(0);
  const firstColumnBold = resolveFirstColumnBold(block, tableDefaults);
  const defaultBodyFill = stringValue(tableDefaults.body_fill) || undefined;
  const bodyParagraphSpacing = resolveTableParagraphSpacing(tableStyleName, false);
  const bodyRowHeight = resolveTableRowHeight(tableStyleName, false);

  return rows.map((row, rowIndex) => {
    const rowItems = arrayValue(row);
    const children: TableCell[] = [];
    let rowCursor = 0;

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      if (pendingRowSpans[columnIndex] > 0) {
        if (isPlaceholderCell(rowItems[rowCursor])) {
          rowCursor += 1;
        }
        children.push(
          new TableCell({
            children: [new Paragraph("")],
            verticalMerge: VerticalMergeType.CONTINUE,
            width: resolveTableCellWidth(columnWidths, columnIndex),
          }),
        );
        pendingRowSpans[columnIndex] -= 1;
        continue;
      }

      const rawCell = rowItems[rowCursor];
      if (rawCell === undefined) {
        throw new RenderCliError(
          "TABLE_ROW_SHAPE_INVALID",
          `Table row exceeds logical column count (${columnCount})`,
        );
      }
      rowCursor += 1;

      const cell = normalizeTableCell(rawCell);
      if (cell.rowSpan > 1) {
        pendingRowSpans[columnIndex] = cell.rowSpan - 1;
      }

      const fill =
        cell.fill ??
        resolveBandedRowFill(block, tableDefaults, tableStyleName, rowIndex + 1) ??
        defaultBodyFill;

      children.push(
        new TableCell({
          width: resolveTableCellWidth(columnWidths, columnIndex),
          children: [
            new Paragraph({
              alignment:
                mapAlignment(cell.align) ??
                resolveTableBodyAlignment(
                  tableStyleName,
                  bodyAlignment,
                  numericColumns,
                  columnIndex,
                ),
              spacing: {
                before: bodyParagraphSpacing.before,
                after: bodyParagraphSpacing.after,
                line: bodyParagraphSpacing.line,
                lineRule: LineRuleType.AUTO,
              },
              children: [
                new TextRun({
                  text: cell.text,
                  bold: cell.bold ?? (firstColumnBold && columnIndex === 0),
                  color: cell.textColor,
                  size: halfPoint(
                    resolveTableFontSize(
                      block,
                      tableStyleName,
                      theme,
                      false,
                      cell.fontScale,
                    ),
                  ),
                  font: buildFontAttributes(theme.tableFontName),
                }),
              ],
            }),
          ],
          verticalMerge:
            cell.rowSpan > 1 ? VerticalMergeType.RESTART : undefined,
          shading: fill
            ? {
                fill,
                color: "auto",
                type: ShadingType.CLEAR,
              }
            : undefined,
        }),
      );
    }

    while (rowCursor < rowItems.length && isPlaceholderCell(rowItems[rowCursor])) {
      rowCursor += 1;
    }
    if (rowCursor !== rowItems.length) {
      throw new RenderCliError(
        "TABLE_ROW_SHAPE_INVALID",
        `Table row exceeds logical column count (${columnCount})`,
      );
    }

    return new TableRow({
      cantSplit: true,
      height: bodyRowHeight
        ? { value: bodyRowHeight.value, rule: HeightRule.ATLEAST }
        : undefined,
      children,
    });
  });
}

export function normalizeTableCell(cell: unknown): TableCellValue {
  if (typeof cell === "string") {
    return { text: cell, rowSpan: 1 };
  }
  const obj = asObject(cell);
  return {
    text: stringValue(obj.text),
    rowSpan: numberValue(obj.row_span) ?? 1,
    fill: stringValue(obj.fill) || undefined,
    textColor: stringValue(obj.text_color) || undefined,
    bold: booleanValue(obj.bold),
    align: stringValue(obj.align) || undefined,
    fontScale: numberValue(obj.font_scale),
  };
}

export function resolveTableColumnCount(headers: string[], rows: unknown[]): number {
  if (headers.length > 0) {
    return headers.length;
  }
  return rows.reduce<number>(
    (max, row) => Math.max(max, countLogicalColumns(arrayValue(row))),
    0,
  );
}

export function normalizeColumnWidths(block: Block, columnCount: number): number[] {
  const widths = arrayValue(block.column_widths)
    .map((value) => numberValue(value))
    .filter((value): value is number => value !== undefined && value > 0)
    .slice(0, columnCount)
    .map((value) => cmToTwip(value));
  if (widths.length === columnCount) {
    return widths;
  }
  const presetWidths = resolvePresetColumnWidths(block, columnCount);
  if (presetWidths) {
    return presetWidths;
  }
  return inferColumnWidths(block, columnCount).map((value) => cmToTwip(value));
}

export function resolveTableCellWidth(
  columnWidths: number[],
  startIndex: number,
  columnSpan = 1,
): { size: number; type: (typeof WidthType)[keyof typeof WidthType] } | undefined {
  if (columnWidths.length === 0) {
    return undefined;
  }
  const width = columnWidths
    .slice(startIndex, startIndex + columnSpan)
    .reduce((sum, value) => sum + value, 0);
  return width > 0 ? { size: width, type: WidthType.DXA } : undefined;
}

function countLogicalColumns(rowItems: unknown[]): number {
  let count = 0;
  for (const rawCell of rowItems) {
    if (!isPlaceholderCell(rawCell)) {
      count += 1;
    }
  }
  return count;
}

function isPlaceholderCell(cell: unknown): boolean {
  if (cell === "") {
    return true;
  }
  if (!cell || typeof cell !== "object" || Array.isArray(cell)) {
    return false;
  }
  const obj = asObject(cell);
  return stringValue(obj.text) === "" && (numberValue(obj.row_span) ?? 1) === 1;
}

type ColumnMetric = {
  header: string;
  maxUnits: number;
  sampleCount: number;
  numericLikeCount: number;
};

const MIN_TABLE_TOTAL_WIDTH_CM = [0, 8.2, 9.8, 11.0, 12.0, 13.2, 14.2];
const MAX_TABLE_TOTAL_WIDTH_CM = [0, 9.8, 11.2, 12.6, 13.8, 14.8, 15.4];

function inferColumnWidths(block: Block, columnCount: number): number[] {
  if (columnCount <= 0) {
    return [];
  }

  const metrics = collectColumnMetrics(block, columnCount);
  const explicitNumericColumns = new Set(
    arrayValue(block.numeric_columns)
      .map((value) => numberValue(value))
      .filter((value): value is number => value !== undefined),
  );
  const widths = metrics.map((metric, columnIndex) =>
    resolveHeuristicColumnWidth(
      metric,
      explicitNumericColumns.has(columnIndex),
    ),
  );
  const currentTotal = widths.reduce((sum, value) => sum + value, 0);
  if (currentTotal <= 0) {
    return [];
  }

  const minTotal =
    MIN_TABLE_TOTAL_WIDTH_CM[Math.min(columnCount, MIN_TABLE_TOTAL_WIDTH_CM.length - 1)] ??
    14.2;
  const maxTotal =
    MAX_TABLE_TOTAL_WIDTH_CM[Math.min(columnCount, MAX_TABLE_TOTAL_WIDTH_CM.length - 1)] ??
    15.4;
  const targetTotal = Math.min(maxTotal, Math.max(minTotal, currentTotal));
  const scale = targetTotal / currentTotal;

  return widths.map((value) => Number((value * scale).toFixed(2)));
}

function resolvePresetColumnWidths(
  block: Block,
  columnCount: number,
): number[] | null {
  if (columnCount !== 5) {
    return null;
  }
  const headers = arrayValue(block.headers)
    .slice(0, columnCount)
    .map((value) => stringValue(value).trim());
  if (headers.length !== 5) {
    return null;
  }

  const firstColumnLike = /(区域|分区|地区|大区|市场|团队|部门|板块|region|area)/i.test(
    headers[0],
  );
  const lastColumnLike = /(备注|说明|comment|note|风险|措施|结论|行动|计划|分析|建议)/i.test(
    headers[4],
  );
  const middleColumnMatches = headers
    .slice(1, 4)
    .filter((header) =>
      /(收入|营收|目标|预算|完成率|利润|毛利|客户|增速|同比|环比|回款|订单|gmv|revenue|target|budget|rate|margin|profit|growth)/i.test(
        header,
      ),
    ).length;

  return firstColumnLike && lastColumnLike && middleColumnMatches >= 2
    ? [1720, 1280, 1280, 1280, 3800]
    : null;
}

function collectColumnMetrics(block: Block, columnCount: number): ColumnMetric[] {
  const headers = arrayValue(block.headers)
    .slice(0, columnCount)
    .map((value) => stringValue(value));
  const metrics: ColumnMetric[] = Array.from({ length: columnCount }, (_, index) => {
    const header = headers[index] ?? "";
    return {
      header,
      maxUnits: Math.max(measureTextUnits(header), 2.2),
      sampleCount: 0,
      numericLikeCount: 0,
    };
  });

  const rows = arrayValue(block.rows);
  const pendingRowSpans = new Array(Math.max(columnCount, 0)).fill(0);

  for (const row of rows) {
    const rowItems = arrayValue(row);
    let rowCursor = 0;

    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      if (pendingRowSpans[columnIndex] > 0) {
        if (isPlaceholderCell(rowItems[rowCursor])) {
          rowCursor += 1;
        }
        pendingRowSpans[columnIndex] -= 1;
        continue;
      }

      const rawCell = rowItems[rowCursor];
      if (rawCell === undefined) {
        break;
      }
      rowCursor += 1;

      const cell = normalizeTableCell(rawCell);
      if (cell.rowSpan > 1) {
        pendingRowSpans[columnIndex] = cell.rowSpan - 1;
      }

      const text = cell.text.trim();
      if (!text) {
        continue;
      }

      metrics[columnIndex].maxUnits = Math.max(
        metrics[columnIndex].maxUnits,
        measureTextUnits(text),
      );
      metrics[columnIndex].sampleCount += 1;
      if (isNumericLike(text)) {
        metrics[columnIndex].numericLikeCount += 1;
      }
    }
  }

  return metrics;
}

function resolveHeuristicColumnWidth(
  metric: ColumnMetric,
  explicitNumericColumn: boolean,
): number {
  const header = metric.header;
  const textUnits = Math.max(metric.maxUnits, measureTextUnits(header), 2.2);
  const numericLike =
    explicitNumericColumn ||
    (metric.sampleCount > 0 && metric.numericLikeCount / metric.sampleCount >= 0.7);
  const remarkLike = /(备注|说明|comment|note|摘要|风险|措施|结论|行动|计划|分析|建议)/i.test(
    header,
  );
  const dateLike = /(日期|时间|月份|季度|周期|date|time|month|quarter)/i.test(header);

  if (remarkLike) {
    return clampWidth(3.9 + Math.min(1.6, Math.max(textUnits - 6, 0) * 0.14), 4.2, 5.6);
  }
  if (numericLike) {
    return clampWidth(1.8 + Math.min(0.8, textUnits * 0.06), 1.9, 3.0);
  }
  if (dateLike) {
    return clampWidth(2.2 + Math.min(0.6, textUnits * 0.05), 2.3, 3.1);
  }
  return clampWidth(2.35 + Math.min(1.25, textUnits * 0.08), 2.4, 4.2);
}

function clampWidth(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function measureTextUnits(text: string): number {
  let units = 0;
  for (const char of text.trim()) {
    if (/\s/.test(char)) {
      units += 0.2;
    } else if (/[0-9]/.test(char)) {
      units += 0.78;
    } else if (/[A-Za-z]/.test(char)) {
      units += 0.92;
    } else if (/[\u3400-\u9FFF\uF900-\uFAFF\u3040-\u30FF\uAC00-\uD7AF]/.test(char)) {
      units += 1.7;
    } else {
      units += 0.55;
    }
  }
  return units;
}

function isNumericLike(text: string): boolean {
  const trimmed = text.trim();
  if (!trimmed) {
    return false;
  }
  const residue = trimmed.replace(/[0-9\s.,%¥￥$€£+\-–—()/:]/g, "");
  return residue.length <= Math.ceil(trimmed.length * 0.2);
}
