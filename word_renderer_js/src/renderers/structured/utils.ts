import {
  AlignmentType,
  HeadingLevel,
  PageOrientation,
  WidthType,
  convertInchesToTwip,
} from "docx";

import { JsonObject } from "../../core/payload";
import { ThemeConfig } from "./types";

export function mapHeadingLevel(level: number) {
  if (level <= 1) {
    return HeadingLevel.HEADING_1;
  }
  if (level === 2) {
    return HeadingLevel.HEADING_2;
  }
  if (level === 3) {
    return HeadingLevel.HEADING_3;
  }
  if (level === 4) {
    return HeadingLevel.HEADING_4;
  }
  if (level === 5) {
    return HeadingLevel.HEADING_5;
  }
  return HeadingLevel.HEADING_6;
}

export function mapAlignment(value: string | undefined) {
  switch (value) {
    case "center":
      return AlignmentType.CENTER;
    case "right":
      return AlignmentType.RIGHT;
    case "justify":
      return AlignmentType.JUSTIFIED;
    case "left":
      return AlignmentType.LEFT;
    default:
      return undefined;
  }
}

export function mapPageOrientation(value: string | undefined) {
  if (value === "landscape") {
    return PageOrientation.LANDSCAPE;
  }
  if (value === "portrait") {
    return PageOrientation.PORTRAIT;
  }
  return undefined;
}

export function normalizeHexColor(value: string): string | undefined {
  const normalized = value.trim().replace(/^#/, "").toUpperCase();
  if (normalized.length !== 6 || /[^0-9A-F]/.test(normalized)) {
    return undefined;
  }
  return normalized;
}

export function point(value: number | undefined): number | undefined {
  return value === undefined ? undefined : Math.round(value * 20);
}

export function halfPoint(value: number): number {
  return Math.round(value * 2);
}

export function borderSize(valuePt: number | undefined, fallbackPt: number): number {
  return Math.max(2, Math.round((valuePt ?? fallbackPt) * 8));
}

export function cmToTwip(value: number): number {
  return convertInchesToTwip(value / 2.54);
}

export function asObject(value: unknown): JsonObject {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }
  return value as JsonObject;
}

export function arrayValue(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

export function stringValue(value: unknown): string {
  return typeof value === "string" ? value : "";
}

export function numberValue(value: unknown): number | undefined {
  return typeof value === "number" ? value : undefined;
}

export function booleanValue(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

export function resolveBoxPadding(
  layout: JsonObject,
  defaultPaddingPt: number,
): {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
  marginUnitType: (typeof WidthType)[keyof typeof WidthType];
} {
  const top = numberValue(layout.padding_top_pt) ?? defaultPaddingPt;
  const right = numberValue(layout.padding_right_pt) ?? defaultPaddingPt;
  const bottom = numberValue(layout.padding_bottom_pt) ?? defaultPaddingPt;
  const left = numberValue(layout.padding_left_pt) ?? defaultPaddingPt;

  return {
    top: point(top),
    right: point(right),
    bottom: point(bottom),
    left: point(left),
    marginUnitType: WidthType.DXA,
  };
}

export function resolveBoxPaddingEdges(
  layout: JsonObject,
  defaults: { top: number; right: number; bottom: number; left: number },
): {
  top?: number;
  bottom?: number;
  left?: number;
  right?: number;
  marginUnitType: (typeof WidthType)[keyof typeof WidthType];
} {
  const top = numberValue(layout.padding_top_pt) ?? defaults.top;
  const right = numberValue(layout.padding_right_pt) ?? defaults.right;
  const bottom = numberValue(layout.padding_bottom_pt) ?? defaults.bottom;
  const left = numberValue(layout.padding_left_pt) ?? defaults.left;

  return {
    top: point(top),
    right: point(right),
    bottom: point(bottom),
    left: point(left),
    marginUnitType: WidthType.DXA,
  };
}

export function resolveContentWidthDxa(
  theme: ThemeConfig,
  sideInsetCm = 0,
): number {
  const pageWidthCm = 21;
  const usableWidthCm =
    pageWidthCm - theme.margins.leftCm - theme.margins.rightCm - sideInsetCm * 2;
  return Math.max(cmToTwip(usableWidthCm), 4000);
}

export function resolveBold(
  defaultBold: boolean,
  emphasis: string | undefined,
): boolean {
  if (emphasis === "strong") {
    return true;
  }
  return defaultBold;
}

export function resolveTextColor(
  theme: ThemeConfig,
  emphasis: string | undefined,
): string | undefined {
  if (emphasis === "subtle") {
    return theme.accent;
  }
  return undefined;
}

export const BunLike = {
  async writeFile(path: string, data: Buffer): Promise<void> {
    const fs = await import("node:fs/promises");
    await fs.writeFile(path, data);
  },
};
