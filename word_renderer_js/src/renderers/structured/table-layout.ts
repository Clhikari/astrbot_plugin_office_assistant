import {
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
              children: [
                new TextRun({
                  text: cell.text,
                  bold: cell.bold ?? (firstColumnBold && columnIndex === 0),
                  color: cell.textColor,
                  size: halfPoint(resolveTableFontSize(tableStyleName, theme, false)),
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

    return new TableRow({ cantSplit: true, children });
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
  return widths.length === columnCount ? widths : [];
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
