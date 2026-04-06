import {
  AlignmentType,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableLayoutType,
  TableRow,
  TextRun,
  WidthType,
} from "docx";

import { RenderCliError } from "../../core/errors";
import { JsonObject } from "../../core/payload";
import { Block, ThemeConfig } from "./types";
import {
  arrayValue,
  asObject,
  halfPoint,
  numberValue,
  stringValue,
} from "./utils";
import {
  buildTableBodyRows,
  normalizeColumnWidths,
  resolveTableCellWidth,
  resolveTableColumnCount,
} from "./table-layout";
import {
  resolveCaptionColor,
  resolveCaptionFill,
  resolveCaptionFontSize,
  resolveDocxTableStyle,
  resolveHeaderBold,
  resolveHeaderFill,
  resolveHeaderTextColor,
  resolveTableAlignment,
  resolveTableBorders,
  resolveTableCellMargin,
  resolveTableFontSize,
} from "./table-style";

export function renderTable(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Table {
  const documentStyle = asObject(metadata.document_style);
  const tableDefaults = asObject(documentStyle.table_defaults);
  const headers = arrayValue(block.headers).map((value) => stringValue(value));
  const rows = arrayValue(block.rows);
  const columnCount = resolveTableColumnCount(headers, rows);
  if (columnCount <= 0) {
    throw new RenderCliError(
      "TABLE_COLUMN_COUNT_INVALID",
      "Table requires at least one column",
    );
  }

  const tableStyleName =
    stringValue(asObject(block.style).table_grid) ||
    stringValue(block.table_style) ||
    stringValue(tableDefaults.preset) ||
    theme.tableStyle;
  const bodyAlignment =
    stringValue(asObject(block.style).cell_align) ||
    stringValue(tableDefaults.cell_align);
  const numericColumns = new Set(
    arrayValue(block.numeric_columns)
      .map((value) => numberValue(value))
      .filter((value): value is number => value !== undefined),
  );
  const columnWidths = normalizeColumnWidths(block, columnCount);

  const tableRows: TableRow[] = [];
  const caption = stringValue(block.caption) || stringValue(block.title);
  if (caption.trim()) {
    tableRows.push(
      new TableRow({
        cantSplit: true,
        children: [
          new TableCell({
            columnSpan: columnCount,
            width: resolveTableCellWidth(columnWidths, 0, columnCount),
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: caption,
                    bold: true,
                    color: resolveCaptionColor(block, tableDefaults, theme),
                    size: halfPoint(
                      resolveCaptionFontSize(block, tableDefaults, theme),
                    ),
                  }),
                ],
              }),
            ],
            shading: {
              fill: resolveCaptionFill(block, tableDefaults, theme),
              color: "auto",
              type: ShadingType.CLEAR,
            },
          }),
        ],
      }),
    );
  }

  const headerGroups = arrayValue(block.header_groups).map((value) => asObject(value));
  if (headerGroups.length > 0) {
    let spanSum = 0;
    let columnCursor = 0;
    const headerFill = resolveHeaderFill(block, tableDefaults, tableStyleName, theme);
    const groupCells = headerGroups.map((group) => {
      const span = numberValue(group.span) ?? 1;
      spanSum += span;
      const width = resolveTableCellWidth(columnWidths, columnCursor, span);
      columnCursor += span;
      return new TableCell({
        columnSpan: span,
        width,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: stringValue(group.title),
                bold: resolveHeaderBold(block),
                color: resolveHeaderTextColor(
                  block,
                  tableDefaults,
                  tableStyleName,
                  theme,
                ),
                size: halfPoint(resolveTableFontSize(tableStyleName, theme, true)),
              }),
            ],
          }),
        ],
        shading: headerFill
          ? {
              fill: headerFill,
              color: "auto",
              type: ShadingType.CLEAR,
            }
          : undefined,
      });
    });
    if (spanSum !== columnCount) {
      throw new RenderCliError(
        "TABLE_HEADER_GROUP_SPAN_INVALID",
        `Header group span total (${spanSum}) does not match column count (${columnCount})`,
      );
    }
    tableRows.push(
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: groupCells,
      }),
    );
  }

  if (headers.length > 0) {
    const headerFill = resolveHeaderFill(block, tableDefaults, tableStyleName, theme);
    tableRows.push(
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: headers.map((header, columnIndex) =>
          new TableCell({
            width: resolveTableCellWidth(columnWidths, columnIndex),
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: header,
                    bold: resolveHeaderBold(block),
                    color: resolveHeaderTextColor(
                      block,
                      tableDefaults,
                      tableStyleName,
                      theme,
                    ),
                    size: halfPoint(resolveTableFontSize(tableStyleName, theme, true)),
                  }),
                ],
              }),
            ],
            shading: headerFill
              ? {
                  fill: headerFill,
                  color: "auto",
                  type: ShadingType.CLEAR,
                }
              : undefined,
          }),
        ),
      }),
    );
  }

  tableRows.push(
    ...buildTableBodyRows(
      block,
      columnCount,
      tableStyleName,
      tableDefaults,
      theme,
      bodyAlignment,
      numericColumns,
      columnWidths,
    ),
  );

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    alignment: resolveTableAlignment(block, tableDefaults),
    layout: columnWidths.length > 0 ? TableLayoutType.FIXED : undefined,
    columnWidths: columnWidths.length > 0 ? columnWidths : undefined,
    style: resolveDocxTableStyle(tableStyleName),
    margins: resolveTableCellMargin(tableStyleName),
    borders: resolveTableBorders(block, tableDefaults, theme, tableStyleName),
    rows: tableRows,
  });
}
