import { SectionType } from "docx";

import { SectionStartType, ThemeConfig } from "./types";

export const SECTION_TYPE_MAP: Record<
  SectionStartType,
  (typeof SectionType)[keyof typeof SectionType]
> = {
  new_page: SectionType.NEXT_PAGE,
  continuous: SectionType.CONTINUOUS,
  odd_page: SectionType.ODD_PAGE,
  even_page: SectionType.EVEN_PAGE,
  new_column: SectionType.NEXT_COLUMN,
};

export const ORDERED_NUMBERING_REFERENCE = "default-numbering";
export const DEFAULT_TABLE_BANDED_ROW_FILL = "F7FBFF";
export const DEFAULT_DIVIDER_COLOR = "D0D7DE";
export const DEFAULT_LIGHT_TABLE_BORDER_COLOR = "D9E1E8";

export const DOCX_TABLE_STYLE_MAP: Record<string, string> = {
  report_grid: "TableGrid",
  metrics_compact: "TableGrid",
  minimal: "TableGrid",
};

export const EXTERNAL_WORD_STYLES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="table" w:default="1" w:styleId="TableNormal">
    <w:name w:val="Normal Table"/>
    <w:uiPriority w:val="99"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
  <w:style w:type="table" w:styleId="TableGrid">
    <w:name w:val="Table Grid"/>
    <w:basedOn w:val="TableNormal"/>
    <w:uiPriority w:val="59"/>
    <w:pPr>
      <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
    </w:pPr>
    <w:tblPr>
      <w:tblInd w:w="0" w:type="dxa"/>
      <w:tblBorders>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="auto"/>
      </w:tblBorders>
      <w:tblCellMar>
        <w:top w:w="0" w:type="dxa"/>
        <w:left w:w="108" w:type="dxa"/>
        <w:bottom w:w="0" w:type="dxa"/>
        <w:right w:w="108" w:type="dxa"/>
      </w:tblCellMar>
    </w:tblPr>
  </w:style>
</w:styles>`;

export const DEFAULT_LIGHT_TABLE_SPECS: Record<
  string,
  { borderSize: number; horizontalMargin: number; verticalMargin: number }
> = {
  report_grid: { borderSize: 3, horizontalMargin: 108, verticalMargin: 136 },
  metrics_compact: { borderSize: 3, horizontalMargin: 84, verticalMargin: 96 },
  minimal: { borderSize: 2, horizontalMargin: 72, verticalMargin: 84 },
};

export const THEMES: Record<string, ThemeConfig> = {
  business_report: {
    themeName: "business_report",
    accent: "1F4E79",
    accentSoft: "DCE6F1",
    fontName: "Microsoft YaHei",
    headingFontName: "Microsoft YaHei",
    tableFontName: "Microsoft YaHei",
    codeFontName: "Consolas",
    titleSize: 20,
    titleAlign: "center",
    titleSpacingAfter: 18,
    headingSize: 15,
    headingSpaceBefore: 10,
    headingSpaceAfter: 0,
    headingBottomBorder: true,
    bodySize: 11,
    bodyIndent: 0,
    bodySpaceAfter: 4,
    bodyLineSpacing: 1.35,
    listSpaceAfter: 4,
    tableFontSize: 10,
    tableStyle: "report_grid",
    summaryFill: "EEF4FA",
    margins: { topCm: 2.0, rightCm: 2.0, bottomCm: 2.0, leftCm: 2.0 },
  },
  project_review: {
    themeName: "project_review",
    accent: "0F766E",
    accentSoft: "D9F3EE",
    fontName: "Microsoft YaHei",
    headingFontName: "Microsoft YaHei",
    tableFontName: "Microsoft YaHei",
    codeFontName: "Consolas",
    titleSize: 19,
    titleAlign: "center",
    titleSpacingAfter: 18,
    headingSize: 14,
    headingSpaceBefore: 10,
    headingSpaceAfter: 0,
    headingBottomBorder: true,
    bodySize: 10,
    bodyIndent: 0,
    bodySpaceAfter: 4,
    bodyLineSpacing: 1.35,
    listSpaceAfter: 4,
    tableFontSize: 10,
    tableStyle: "metrics_compact",
    summaryFill: "E8F6F3",
    margins: { topCm: 2.0, rightCm: 2.0, bottomCm: 2.0, leftCm: 2.0 },
  },
  executive_brief: {
    themeName: "executive_brief",
    accent: "B45309",
    accentSoft: "FDEBD8",
    fontName: "Microsoft YaHei",
    headingFontName: "Microsoft YaHei",
    tableFontName: "Microsoft YaHei",
    codeFontName: "Consolas",
    titleSize: 18.5,
    titleAlign: "center",
    titleSpacingAfter: 18,
    headingSize: 13,
    headingSpaceBefore: 10,
    headingSpaceAfter: 0,
    headingBottomBorder: true,
    bodySize: 10,
    bodyIndent: 0,
    bodySpaceAfter: 4,
    bodyLineSpacing: 1.35,
    listSpaceAfter: 4,
    tableFontSize: 10,
    tableStyle: "minimal",
    summaryFill: "FFF7ED",
    margins: { topCm: 2.0, rightCm: 2.0, bottomCm: 2.0, leftCm: 2.0 },
  },
};
