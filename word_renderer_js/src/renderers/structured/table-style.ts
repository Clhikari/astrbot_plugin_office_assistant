import { AlignmentType, BorderStyle, WidthType } from "docx";

import { JsonObject } from "../../core/payload";
import {
  DEFAULT_DIVIDER_COLOR,
  DEFAULT_LIGHT_TABLE_BORDER_COLOR,
  DEFAULT_LIGHT_TABLE_SPECS,
  DEFAULT_TABLE_BANDED_ROW_FILL,
  DOCX_TABLE_STYLE_MAP,
} from "./constants";
import { Block, ThemeConfig } from "./types";
import {
  booleanValue,
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
): JsonObject | undefined {
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
  tableStyleName: string,
  theme: ThemeConfig,
  header: boolean,
): number {
  const baseSize = theme.tableFontSize;
  if (tableStyleName === "metrics_compact") {
    return Math.max(baseSize - 0.5, 9);
  }
  if (tableStyleName === "minimal" && header) {
    return Math.max(baseSize, theme.bodySize);
  }
  return baseSize;
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
  return {
    left: spec.horizontalMargin,
    right: spec.horizontalMargin,
    marginUnitType: WidthType.DXA,
  };
}

export function resolveDocxTableStyle(tableStyleName: string): string | undefined {
  return DOCX_TABLE_STYLE_MAP[tableStyleName];
}
