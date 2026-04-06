import { Paragraph, Table } from "docx";

import { JsonObject } from "../../core/payload";

export type FileChild = Paragraph | Table;
export type Block = JsonObject & { type: string };
export type HeaderFooterConfig = JsonObject;

export type ThemeConfig = {
  accent: string;
  accentSoft: string;
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
  rowSpan: number;
  fill?: string;
  textColor?: string;
  bold?: boolean;
  align?: string;
  fontScale?: number;
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
};
