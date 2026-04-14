import { BorderStyle, Paragraph, Table } from "docx";

import { JsonObject } from "../../core/payload";

export type FileChild = Paragraph | Table;
export type Block = JsonObject & { type: string };
export type HeaderFooterConfig = JsonObject;

export type ThemeConfig = {
  themeName: string;
  accent: string;
  accentSoft: string;
  accentBoxStripColor: string;
  headingBottomBorderColor: string;
  heroBannerDividerColor: string;
  heroBannerDividerSizePt: number;
  fontName: string;
  headingFontName: string;
  tableFontName: string;
  codeFontName: string;
  titleSize: number;
  titleAlign: string;
  titleSpacingAfter: number;
  headingSize: number;
  headingSpaceBefore: number;
  headingSpaceAfter: number;
  headingBottomBorder: boolean;
  headingBottomBorderSizePt: number;
  bodySize: number;
  bodyIndent: number;
  bodySpaceAfter: number;
  bodyLineSpacing: number;
  listSpaceAfter: number;
  tableFontSize: number;
  tableStyle: string;
  summaryFill: string;
  margins: {
    topCm: number;
    rightCm: number;
    bottomCm: number;
    leftCm: number;
  };
};

export type TableCellValue = {
  text: string;
  runs?: JsonObject[];
  rowSpan: number;
  colSpan: number;
  fill?: string;
  textColor?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  align?: string;
  fontName?: string;
  fontScale?: number;
  border?: DocxBorderSpec;
};

export type BorderSideSpec = {
  style?: string;
  color?: string;
  width_pt?: number;
};

export type BorderSpec = {
  top?: BorderSideSpec;
  bottom?: BorderSideSpec;
  left?: BorderSideSpec;
  right?: BorderSideSpec;
};

export type DocxBorderSideSpec = {
  style: (typeof BorderStyle)[keyof typeof BorderStyle];
  size: number;
  color?: string;
};

export type DocxBorderSpec = {
  top?: DocxBorderSideSpec;
  bottom?: DocxBorderSideSpec;
  left?: DocxBorderSideSpec;
  right?: DocxBorderSideSpec;
  insideHorizontal?: DocxBorderSideSpec;
  insideVertical?: DocxBorderSideSpec;
};

export type SectionStartType =
  | "new_page"
  | "continuous"
  | "odd_page"
  | "even_page"
  | "new_column";

export type SectionState = {
  children: FileChild[];
  headerFooter: HeaderFooterConfig;
  startType?: SectionStartType;
  pageOrientation?: string;
  margins?: JsonObject;
  restartPageNumbering: boolean;
  pageNumberStart?: number;
  inheritPreviousHeaderFooter: boolean;
};

export type RunDefaults = {
  fontSize?: number;
  emphasis?: string;
  fontScale?: number;
  fontName?: string;
  codeFontName?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strikethrough?: boolean;
  color?: string;
};
