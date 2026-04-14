import {
  ExternalHyperlink,
  HeightRule,
  LineRuleType,
  Paragraph,
  ParagraphChild,
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
import { resolveBorderSpec, resolveTableCellBorders } from "./borders";
import {
  buildTextRuns,
  DEFAULT_HYPERLINK_COLOR,
  normalizeHyperlinkTarget,
  normalizeLineBreaks,
} from "./run-helpers";
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
  resolveTableBorders,
  resolveTableFontSize,
  resolveTableParagraphSpacing,
  resolveTableRowHeight,
} from "./table-style";

type CellRunDefaults = {
  fontSize: number;
  fontScale?: number;
  fontName?: string;
  codeFontName: string;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  color?: string;
};

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
  const tableBorders = resolveTableBorders(block, tableDefaults, theme, tableStyleName);

  return rows.map((row, rowIndex) => {
    const rowItems = arrayValue(row);
    const children: TableCell[] = [];
    let rowCursor = 0;
    let columnIndex = 0;

    while (columnIndex < columnCount) {
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
        columnIndex += 1;
        continue;
      }

      const rawCell = rowItems[rowCursor];
      if (rawCell === undefined) {
        throw new RenderCliError(
          "TABLE_ROW_SHAPE_INVALID",
          `Table row is missing cells (expected ${columnCount})`,
        );
      }
      rowCursor += 1;

      const cell = normalizeTableCell(rawCell);
      if (columnIndex + cell.colSpan > columnCount) {
        throw new RenderCliError(
          "TABLE_ROW_SHAPE_INVALID",
          `Table row exceeds logical column count (${columnCount})`,
        );
      }
      for (let spanIndex = 0; spanIndex < cell.colSpan; spanIndex += 1) {
        if (pendingRowSpans[columnIndex + spanIndex] > 0) {
          throw new RenderCliError(
            "TABLE_ROW_SHAPE_INVALID",
            `Table row ${rowIndex + 1} overlaps active row spans`,
          );
        }
      }
      if (cell.rowSpan > 1) {
        for (let spanIndex = 0; spanIndex < cell.colSpan; spanIndex += 1) {
          pendingRowSpans[columnIndex + spanIndex] = cell.rowSpan - 1;
        }
      }

      const fill =
        cell.fill ??
        resolveBandedRowFill(block, tableDefaults, tableStyleName, rowIndex + 1) ??
        defaultBodyFill;

      children.push(
        new TableCell({
          width: resolveTableCellWidth(columnWidths, columnIndex, cell.colSpan),
          columnSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
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
              children: buildTableCellTextRuns(
                cell,
                firstColumnBold,
                columnIndex,
                block,
                tableStyleName,
                theme,
              ),
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
          borders: cell.border
            ? resolveTableCellBorders(tableBorders, cell.border)
            : undefined,
        }),
      );
      columnIndex += cell.colSpan;
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
    return { text: cell, rowSpan: 1, colSpan: 1 };
  }
  const obj = asObject(cell);
  const rowSpan = numberValue(obj.row_span) ?? 1;
  const colSpan = numberValue(obj.col_span) ?? 1;
  if (
    !Number.isInteger(rowSpan) ||
    !Number.isInteger(colSpan) ||
    rowSpan < 1 ||
    colSpan < 1
  ) {
    throw new RenderCliError(
      "TABLE_CELL_SPAN_INVALID",
      "Table cell spans must be integers greater than or equal to 1",
    );
  }
  if (rowSpan > 1 && colSpan > 1) {
    throw new RenderCliError(
      "TABLE_CELL_SPAN_INVALID",
      "Table cell cannot combine row_span and col_span",
    );
  }
  return {
    text: stringValue(obj.text),
    runs: arrayValue(obj.runs) as JsonObject[],
    rowSpan,
    colSpan,
    fill: stringValue(obj.fill) || undefined,
    textColor: stringValue(obj.text_color) || undefined,
    bold: booleanValue(obj.bold),
    italic: booleanValue(obj.italic),
    underline: booleanValue(obj.underline),
    strikethrough: booleanValue(obj.strikethrough),
    align: stringValue(obj.align) || undefined,
    fontName: stringValue(obj.font_name) || undefined,
    fontScale: numberValue(obj.font_scale),
    border: resolveBorderSpec(obj.border),
  };
}

function buildTableCellTextRuns(
  cell: TableCellValue,
  firstColumnBold: boolean,
  columnIndex: number,
  block: Block,
  tableStyleName: string,
  theme: ThemeConfig,
): ParagraphChild[] {
  const defaults = resolveCellRunDefaults(
    cell,
    firstColumnBold,
    columnIndex,
    block,
    tableStyleName,
    theme,
  );
  const cellRuns = arrayValue(cell.runs);
  if (cellRuns.length > 0) {
    return cellRuns.flatMap((run) =>
      buildRichCellRunChildren(run, defaults, block, tableStyleName, theme),
    );
  }
  return buildPlainCellRunChildren(cell.text, defaults);
}

function resolveCellRunDefaults(
  cell: TableCellValue,
  firstColumnBold: boolean,
  columnIndex: number,
  block: Block,
  tableStyleName: string,
  theme: ThemeConfig,
): CellRunDefaults {
  return {
    fontSize: resolveTableFontSize(
      block,
      tableStyleName,
      theme,
      false,
      cell.fontScale,
    ),
    fontScale: cell.fontScale,
    fontName: cell.fontName || theme.tableFontName,
    codeFontName: theme.codeFontName || "Consolas",
    bold: cell.bold ?? (firstColumnBold && columnIndex === 0),
    italic: booleanValue(cell.italic) ?? false,
    underline: booleanValue(cell.underline) ?? false,
    strikethrough: booleanValue(cell.strikethrough) ?? false,
    color: cell.textColor || undefined,
  };
}

function buildPlainCellRunChildren(
  text: string,
  defaults: CellRunDefaults,
): TextRun[] {
  return buildTextRuns(text, {
    bold: defaults.bold,
    italics: defaults.italic,
    underline: defaults.underline ? {} : undefined,
    strike: defaults.strikethrough,
    color: defaults.color,
    size: halfPoint(defaults.fontSize),
    font: defaults.fontName ? buildFontAttributes(defaults.fontName) : undefined,
  });
}

function buildRichCellRunChildren(
  rawRun: unknown,
  defaults: CellRunDefaults,
  block: Block,
  tableStyleName: string,
  theme: ThemeConfig,
): ParagraphChild[] {
  const run = asObject(rawRun);
  const hyperlinkTarget = normalizeHyperlinkTarget(run.url);
  const text = stringValue(run.text);
  const effectiveFontScale = numberValue(run.font_scale) ?? defaults.fontScale;
  const resolvedFontSize = halfPoint(
    resolveTableFontSize(
      block,
      tableStyleName,
      theme,
      false,
      effectiveFontScale,
    ),
  );
  const fontName =
    stringValue(run.font_name) ||
    (booleanValue(run.code) === true ? defaults.codeFontName : defaults.fontName);
  const textRuns = buildTextRuns(text, {
    bold: booleanValue(run.bold) ?? defaults.bold,
    italics: booleanValue(run.italic) ?? defaults.italic,
    underline: hyperlinkTarget
      ? {}
      : ((booleanValue(run.underline) ?? defaults.underline) ? {} : undefined),
    strike: booleanValue(run.strikethrough) ?? defaults.strikethrough,
    color:
      stringValue(run.color) ||
      (hyperlinkTarget ? DEFAULT_HYPERLINK_COLOR : defaults.color),
    size: resolvedFontSize,
    font: fontName ? buildFontAttributes(fontName) : undefined,
  });

  if (!hyperlinkTarget) {
    return textRuns;
  }
  return [new ExternalHyperlink({ link: hyperlinkTarget, children: textRuns })];
}

export function resolveTableColumnCount(headers: string[], rows: unknown[]): number {
  if (headers.length > 0) {
    return headers.length;
  }
  return resolveBodyColumnCount(rows);
}

function resolveBodyColumnCount(rows: unknown[]): number {
  let activeSpans: number[] = [];
  let maxColumns = 0;

  for (const row of rows) {
    const rowItems = arrayValue(row);
    const nextSpans = activeSpans.map((s) => Math.max(s - 1, 0));
    let colIdx = 0;
    let itemIdx = 0;

    while (itemIdx < rowItems.length) {
      const rawCell = rowItems[itemIdx];

      // Skip columns occupied by active spans
      while (colIdx < activeSpans.length && activeSpans[colIdx] > 0) {
        if (isPlaceholderCell(rawCell)) {
          // Placeholder consumed by active span
          itemIdx += 1;
          break;
        }
        colIdx += 1;
      }
      if (itemIdx >= rowItems.length) {
        break;
      }
      // Re-check after consuming placeholder
      if (colIdx < activeSpans.length && activeSpans[colIdx] > 0) {
        continue;
      }

      const cell = rowItems[itemIdx];
      if (isPlaceholderCell(cell)) {
        // Trailing placeholder not consumed by a span - skip it
        itemIdx += 1;
        continue;
      }

      // Expand spans arrays if needed
      while (activeSpans.length <= colIdx) {
        activeSpans.push(0);
        nextSpans.push(0);
      }

      const normalized = normalizeTableCell(cell);
      const colSpan = normalized.colSpan;
      if (normalized.rowSpan > 1) {
        while (activeSpans.length < colIdx + colSpan) {
          activeSpans.push(0);
          nextSpans.push(0);
        }
        for (let spanIndex = 0; spanIndex < colSpan; spanIndex += 1) {
          nextSpans[colIdx + spanIndex] = Math.max(
            nextSpans[colIdx + spanIndex],
            normalized.rowSpan - 1,
          );
        }
      }

      colIdx += colSpan;
      itemIdx += 1;
    }

    // Account for any remaining active spans after processing row items
    while (colIdx < activeSpans.length && activeSpans[colIdx] > 0) {
      colIdx += 1;
    }

    maxColumns = Math.max(maxColumns, activeSpans.length, nextSpans.length, colIdx);
    activeSpans = nextSpans;
  }

  return maxColumns;
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

function isPlaceholderCell(cell: unknown): boolean {
  if (cell === "") {
    return true;
  }
  if (!cell || typeof cell !== "object" || Array.isArray(cell)) {
    return false;
  }
  const obj = asObject(cell);
  return (
    stringValue(obj.text) === "" &&
    arrayValue(obj.runs).length === 0 &&
    (numberValue(obj.row_span) ?? 1) === 1 &&
    (numberValue(obj.col_span) ?? 1) === 1
  );
}

function tableCellPlainText(cell: TableCellValue): string {
  const runs = arrayValue(cell.runs);
  if (runs.length > 0) {
    return runs
      .map((run) => normalizeLineBreaks(stringValue(asObject(run).text)))
      .join("");
  }
  return normalizeLineBreaks(cell.text);
}

type ColumnMetric = {
  header: string;
  maxUnits: number;
  sampleCount: number;
  numericLikeCount: number;
};

const MIN_TABLE_TOTAL_WIDTH_CM = [0, 8.2, 9.8, 11.0, 12.0, 13.2, 14.2];
const MAX_TABLE_TOTAL_WIDTH_CM = [0, 9.8, 11.2, 12.6, 13.8, 14.8, 15.4];
const MAX_COLUMN_INFERENCE_SAMPLE_ROWS = 50;

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

  for (const row of rows.slice(0, MAX_COLUMN_INFERENCE_SAMPLE_ROWS)) {
    const rowItems = arrayValue(row);
    let rowCursor = 0;
    let columnIndex = 0;

    while (columnIndex < columnCount) {
      if (pendingRowSpans[columnIndex] > 0) {
        if (isPlaceholderCell(rowItems[rowCursor])) {
          rowCursor += 1;
        }
        pendingRowSpans[columnIndex] -= 1;
        columnIndex += 1;
        continue;
      }

      const rawCell = rowItems[rowCursor];
      if (rawCell === undefined) {
        break;
      }
      rowCursor += 1;

      const cell = normalizeTableCell(rawCell);
      if (columnIndex + cell.colSpan > columnCount) {
        throw new RenderCliError(
          "TABLE_ROW_SHAPE_INVALID",
          `Table row exceeds logical column count (${columnCount})`,
        );
      }
      for (let spanIndex = 0; spanIndex < cell.colSpan; spanIndex += 1) {
        if (pendingRowSpans[columnIndex + spanIndex] > 0) {
          throw new RenderCliError(
            "TABLE_ROW_SHAPE_INVALID",
            "Table row overlaps active row spans during width inference",
          );
        }
      }
      const colSpan = cell.colSpan;
      if (cell.rowSpan > 1) {
        for (let spanIndex = 0; spanIndex < colSpan; spanIndex += 1) {
          pendingRowSpans[columnIndex + spanIndex] = cell.rowSpan - 1;
        }
      }

      const metricText = tableCellPlainText(cell).trim();
      if (!metricText) {
        columnIndex += colSpan;
        continue;
      }

      const measuredUnits = Math.max(measureTextUnits(metricText) / colSpan, 1.2);
      for (let spanIndex = 0; spanIndex < colSpan; spanIndex += 1) {
        metrics[columnIndex + spanIndex].maxUnits = Math.max(
          metrics[columnIndex + spanIndex].maxUnits,
          measuredUnits,
        );
        metrics[columnIndex + spanIndex].sampleCount += 1;
        if (colSpan === 1 && isNumericLike(metricText)) {
          metrics[columnIndex + spanIndex].numericLikeCount += 1;
        }
      }
      columnIndex += colSpan;
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

  // These ranges are tuned to keep generated tables readable in Word:
  // - remark/comment columns stay noticeably wider
  // - numeric and date-like columns stay compact
  // - all inferred widths remain bounded so the later total-width scaling step can
  //   normalize the table without producing extreme single-column growth
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
