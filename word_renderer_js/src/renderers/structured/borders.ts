import { BorderStyle } from "docx";

import { DocxBorderSideSpec, DocxBorderSpec } from "./types";
import { asObject, borderSize, normalizeHexColor, stringValue } from "./utils";

const BORDER_STYLE_MAP = {
  single: BorderStyle.SINGLE,
  double: BorderStyle.DOUBLE,
  dashed: BorderStyle.DASHED,
  dotted: BorderStyle.DOTTED,
  none: BorderStyle.NONE,
} as const;
const CELL_BORDER_SIDES = ["top", "bottom", "left", "right"] as const;

function resolveBorderSide(side: unknown): DocxBorderSideSpec | undefined {
  if (!side || typeof side !== "object" || Array.isArray(side)) {
    return undefined;
  }
  const sideObject = asObject(side);
  const styleName = stringValue(sideObject.style) || "single";
  const style = BORDER_STYLE_MAP[styleName as keyof typeof BORDER_STYLE_MAP];
  if (!style) {
    return undefined;
  }

  const color = normalizeHexColor(stringValue(sideObject.color));
  const widthPt =
    typeof sideObject.width_pt === "number" ? sideObject.width_pt : undefined;
  const resolved: DocxBorderSideSpec = {
    style,
    size: borderSize(widthPt, 0.5),
  };
  if (color) {
    resolved.color = color;
  }
  return resolved;
}

export function resolveBorderSpec(border: unknown): DocxBorderSpec | undefined {
  if (!border || typeof border !== "object" || Array.isArray(border)) {
    return undefined;
  }
  const borderObject = asObject(border);
  const top = Object.prototype.hasOwnProperty.call(borderObject, "top")
    ? resolveBorderSide(borderObject.top)
    : undefined;
  const bottom = Object.prototype.hasOwnProperty.call(borderObject, "bottom")
    ? resolveBorderSide(borderObject.bottom)
    : undefined;
  const left = Object.prototype.hasOwnProperty.call(borderObject, "left")
    ? resolveBorderSide(borderObject.left)
    : undefined;
  const right = Object.prototype.hasOwnProperty.call(borderObject, "right")
    ? resolveBorderSide(borderObject.right)
    : undefined;

  if (!top && !bottom && !left && !right) {
    return undefined;
  }

  const resolved: DocxBorderSpec = {};
  if (top) {
    resolved.top = top;
  }
  if (bottom) {
    resolved.bottom = bottom;
  }
  if (left) {
    resolved.left = left;
  }
  if (right) {
    resolved.right = right;
  }
  return resolved;
}

export function mergeBorderSpecs(
  base: DocxBorderSpec | undefined,
  override: DocxBorderSpec | undefined,
): DocxBorderSpec | undefined {
  if (!base) {
    return override;
  }
  if (!override) {
    return base;
  }

  const merged: DocxBorderSpec = {
    ...base,
    ...override,
  };
  return Object.keys(merged).length > 0 ? merged : undefined;
}

function extractCellBorders(
  borders: DocxBorderSpec | undefined,
): DocxBorderSpec | undefined {
  if (!borders) {
    return undefined;
  }

  const resolved: DocxBorderSpec = {};
  for (const side of CELL_BORDER_SIDES) {
    const border = borders[side];
    if (border) {
      resolved[side] = border;
    }
  }
  return Object.keys(resolved).length > 0 ? resolved : undefined;
}

export function resolveTableCellBorders(
  tableBorders: DocxBorderSpec | undefined,
  cellBorder: DocxBorderSpec | undefined,
): DocxBorderSpec | undefined {
  return mergeBorderSpecs(
    extractCellBorders(tableBorders),
    extractCellBorders(cellBorder),
  );
}
