import { AlignmentType, BorderStyle, WidthType } from "docx";

import { JsonObject } from "../../core/payload";
import {
  DEFAULT_DIVIDER_COLOR,
  DEFAULT_LIGHT_TABLE_BORDER_COLOR,
  DEFAULT_LIGHT_TABLE_SPECS,
  DEFAULT_TABLE_BANDED_ROW_FILL,
  DOCX_TABLE_STYLE_MAP,
} from "./constants";
import { Block, DocxBorderSpec, ThemeConfig } from "./types";
import {
  booleanValue,
  numberValue,
  mapAlignment,
  normalizeHexColor,
  stringValue,
} from "./utils";

export function resolveHeaderFill(
  block: Block,
  tableDefaults: JsonObject,
  tableStyleName: string,
  theme: ThemeConfig,
): string | undefined {
  if (booleanValue(block.header_fill_enabled) === false) {
    return undefined;
  }
  const explicit = normalizeHexColor(stringValue(block.header_fill));
  if (explicit) {
    return explicit;
  }
  const defaultFill = normalizeHexColor(stringValue(tableDefaults.header_fill));
  if (defaultFill) {
    return defaultFill;
  }
  return tableStyleName === "minimal" ? theme.accentSoft : theme.accent;
}

export function resolveHeaderTextColor(
  block: Block,
  tableDefaults: JsonObject,
  tableStyleName: string,
  theme: ThemeConfig,
): string {
  return (
    normalizeHexColor(stringValue(block.header_text_color)) ||
    normalizeHexColor(stringValue(tableDefaults.header_text_color)) ||
    (tableStyleName === "minimal" ? theme.accent : "FFFFFF")
  );
}

export function resolveHeaderBold(block: Block): boolean {
  const explicit = booleanValue(block.header_bold);
  return explicit === undefined ? true : explicit;
}

export function resolveBandedRowFill(
  block: Block,
  tableDefaults: JsonObject,
  tableStyleName: string,
  rowIndex: number,
): string | undefined {
  const bandedRows =
    booleanValue(block.banded_rows) ?? booleanValue(tableDefaults.banded_rows);
  const fill =
    normalizeHexColor(stringValue(block.banded_row_fill)) ||
    normalizeHexColor(stringValue(tableDefaults.banded_row_fill)) ||
    DEFAULT_TABLE_BANDED_ROW_FILL;

  if (bandedRows === true) {
    return rowIndex % 2 === 1 ? fill : undefined;
  }
  if (bandedRows === false) {
    return undefined;
  }
  return tableStyleName === "report_grid" && rowIndex % 2 === 1
    ? DEFAULT_TABLE_BANDED_ROW_FILL
    : undefined;
}

export function resolveFirstColumnBold(
  block: Block,
  tableDefaults: JsonObject,
): boolean {
  const explicit = booleanValue(block.first_column_bold);
  return explicit === undefined
    ? booleanValue(tableDefaults.first_column_bold) === true
    : explicit;
}

export function resolveTableAlignment(
  block: Block,
  tableDefaults: JsonObject,
) {
  const align =
    stringValue(block.table_align) || stringValue(tableDefaults.table_align);
  if (align === "left") {
    return AlignmentType.LEFT;
  }
  if (align === "center") {
    return AlignmentType.CENTER;
  }
  return undefined;
}

export function resolveTableBorders(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
  tableStyleName: string,
): DocxBorderSpec | undefined {
  const borderStyle =
    stringValue(block.border_style) || stringValue(tableDefaults.border_style);
  const borderMap: Record<string, { size: number; color: string }> = {
    minimal: { size: 4, color: DEFAULT_DIVIDER_COLOR },
    standard: { size: 8, color: "7A7A7A" },
    strong: { size: 16, color: theme.accent },
  };
  const defaultLightSpec = DEFAULT_LIGHT_TABLE_SPECS[tableStyleName];
  const spec =
    borderMap[borderStyle] ??
    (defaultLightSpec
      ? {
          size: defaultLightSpec.borderSize,
          color: DEFAULT_LIGHT_TABLE_BORDER_COLOR,
        }
      : undefined);
  if (!spec) {
    return undefined;
  }
  return {
    top: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    bottom: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    left: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    right: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    insideHorizontal: {
      style: BorderStyle.SINGLE,
      size: spec.size,
      color: spec.color,
    },
    insideVertical: {
      style: BorderStyle.SINGLE,
      size: spec.size,
      color: spec.color,
    },
  };
}

export function resolveCaptionFill(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
): string {
  const emphasis =
    stringValue(block.caption_emphasis) ||
    stringValue(tableDefaults.caption_emphasis);
  if (emphasis === "strong") {
    return (
      resolveHeaderFill(
        block,
        tableDefaults,
        stringValue(block.table_style) || theme.tableStyle,
        theme,
      ) || theme.accent
    );
  }
  return theme.accentSoft;
}

export function resolveCaptionColor(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
): string {
  const emphasis =
    stringValue(block.caption_emphasis) ||
    stringValue(tableDefaults.caption_emphasis);
  if (emphasis === "strong") {
    return (
      normalizeHexColor(stringValue(block.header_text_color)) ||
      normalizeHexColor(stringValue(tableDefaults.header_text_color)) ||
      "FFFFFF"
    );
  }
  return theme.accent;
}

export function resolveCaptionFontSize(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
): number {
  const emphasis =
    stringValue(block.caption_emphasis) ||
    stringValue(tableDefaults.caption_emphasis);
  const baseSize = Math.max(theme.bodySize, 11);
  return emphasis === "strong" ? baseSize + 1 : baseSize;
}

export function resolveTableFontSize(
  block: Block,
  tableStyleName: string,
  theme: ThemeConfig,
  header: boolean,
  cellFontScale?: number,
): number {
  const baseSize = header ? theme.tableFontSize + 0.5 : theme.tableFontSize;
  const defaultScale =
    numberValue(header ? block.header_font_scale : block.body_font_scale) ?? 1;
  const effectiveScale = cellFontScale ?? defaultScale;
  const scaledBaseSize = baseSize * effectiveScale;
  if (tableStyleName === "metrics_compact") {
    return Math.max(scaledBaseSize - 0.5, 9);
  }
  if (tableStyleName === "minimal" && header) {
    return Math.max(scaledBaseSize, theme.bodySize);
  }
  return scaledBaseSize;
}

export function resolveTableBodyAlignment(
  tableStyleName: string,
  bodyAlignment: string,
  numericColumns: Set<number>,
  columnIndex: number,
) {
  const explicit = mapAlignment(bodyAlignment);
  if (explicit) {
    return explicit;
  }
  if (tableStyleName === "report_grid") {
    return AlignmentType.CENTER;
  }
  if (numericColumns.has(columnIndex)) {
    return AlignmentType.RIGHT;
  }
  if (tableStyleName === "metrics_compact" && columnIndex > 0) {
    return AlignmentType.CENTER;
  }
  return AlignmentType.LEFT;
}

export function resolveTableCellMargin(
  tableStyleName: string,
  block?: Block,
): {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
  marginUnitType?: (typeof WidthType)[keyof typeof WidthType];
} | undefined {
  const spec = DEFAULT_LIGHT_TABLE_SPECS[tableStyleName];
  if (!spec) {
    return undefined;
  }
  const horizontalPadding = numberValue(block?.cell_padding_horizontal_pt);
  const verticalPadding = numberValue(block?.cell_padding_vertical_pt);
  return {
    top:
      verticalPadding !== undefined
        ? Math.round(verticalPadding * 20)
        : spec.verticalMargin,
    bottom:
      verticalPadding !== undefined
        ? Math.round(verticalPadding * 20)
        : spec.verticalMargin,
    left:
      horizontalPadding !== undefined
        ? Math.round(horizontalPadding * 20)
        : spec.horizontalMargin,
    right:
      horizontalPadding !== undefined
        ? Math.round(horizontalPadding * 20)
        : spec.horizontalMargin,
    marginUnitType: WidthType.DXA,
  };
}

export function resolveDocxTableStyle(tableStyleName: string): string | undefined {
  return DOCX_TABLE_STYLE_MAP[tableStyleName];
}

export function resolveTableParagraphSpacing(
  tableStyleName: string,
  header: boolean,
): {
  before?: number;
  after?: number;
  line?: number;
} {
  if (tableStyleName === "report_grid") {
    return header
      ? { before: 10, after: 10, line: 300 }
      : { before: 8, after: 8, line: 312 };
  }
  if (tableStyleName === "metrics_compact") {
    return header
      ? { before: 6, after: 6, line: 264 }
      : { before: 4, after: 4, line: 252 };
  }
  return header
    ? { before: 4, after: 4, line: 252 }
    : { before: 3, after: 3, line: 240 };
}

export function resolveTableRowHeight(
  tableStyleName: string,
  header: boolean,
): { value: number; rule: "atLeast" } | undefined {
  if (tableStyleName !== "report_grid") {
    return undefined;
  }
  return {
    value: header ? 520 : 480,
    rule: "atLeast",
  };
}
