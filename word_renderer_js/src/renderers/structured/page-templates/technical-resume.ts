import {
  AlignmentType,
  BorderStyle,
  LineRuleType,
  PageBreak,
  Paragraph,
  TabStopType,
  TextRun,
} from "docx";

import { JsonObject } from "../../../core/payload";
import { buildFontAttributes, normalizeInlineItem } from "../inline";
import { FileChild, ThemeConfig } from "../types";
import {
  arrayValue,
  asObject,
  booleanValue,
  halfPoint,
  point,
  stringValue,
} from "../utils";

const RESUME_FONT_NAME = "CMU Serif";
const RESUME_BODY_COLOR = "1A1A1A";
const RESUME_MUTED_COLOR = "444444";
const RESUME_DATE_COLOR = "444444";
const RESUME_SECTION_COLOR = "1A1A1A";
const RESUME_DIVIDER_COLOR = "000000";
const RESUME_DIVIDER_SIZE = 4;
const RESUME_DATE_TAB_POSITION = 9026;

export function renderTechnicalResume(
  block: JsonObject,
  _metadata: JsonObject,
  _theme: ThemeConfig,
): FileChild[] {
  const data = asObject(block.data);
  const children: FileChild[] = [buildNameParagraph(stringValue(data.name))];
  const headline = stringValue(data.headline);
  if (headline.trim()) {
    children.push(buildHeadlineParagraph(headline));
  }
  children.push(
    buildContactParagraph(stringValue(data.contact_line)),
    buildContactDividerParagraph(),
  );

  for (const rawSection of arrayValue(data.sections)) {
    children.push(...buildResumeSection(asObject(rawSection)));
  }
  if (booleanValue(data.auto_page_break) === true) {
    children.push(
      new Paragraph({
        spacing: { before: point(4) },
        children: [new PageBreak()],
      }),
    );
  }

  return children;
}

function buildResumeSection(section: JsonObject): FileChild[] {
  const children: FileChild[] = [
    new Paragraph({
      spacing: {
        before: point(10),
        after: 0,
      },
      border: {
        bottom: {
          color: RESUME_DIVIDER_COLOR,
          style: BorderStyle.SINGLE,
          size: RESUME_DIVIDER_SIZE,
        },
      },
      children: [
        new TextRun({
          text: stringValue(section.title),
          bold: true,
          color: RESUME_SECTION_COLOR,
          size: halfPoint(12),
          font: buildFontAttributes(RESUME_FONT_NAME),
        }),
      ],
    }),
  ];

  const entries = arrayValue(section.entries);
  if (entries.length > 0) {
    for (const rawEntry of entries) {
      children.push(...buildResumeEntry(asObject(rawEntry)));
    }
  } else {
    for (const rawLine of arrayValue(section.lines)) {
      children.push(buildDetailParagraph(rawLine, { afterPt: 4 }));
    }
  }

  return children;
}

function buildResumeEntry(entry: JsonObject): FileChild[] {
  const headingLineChildren: TextRun[] = [
    new TextRun({
      text: stringValue(entry.heading),
      bold: true,
      color: RESUME_SECTION_COLOR,
      size: halfPoint(11),
      font: buildFontAttributes(RESUME_FONT_NAME),
    }),
    ...(stringValue(entry.date)
      ? [
          new TextRun({
            text: `\t${stringValue(entry.date)}`,
            color: RESUME_DATE_COLOR,
            italics: true,
            size: halfPoint(10.5),
            font: buildFontAttributes(RESUME_FONT_NAME),
          }),
        ]
      : []),
  ];
  const subtitle = stringValue(entry.subtitle);
  const children: FileChild[] = [
    new Paragraph({
      tabStops: [
        {
          type: TabStopType.RIGHT,
          position: RESUME_DATE_TAB_POSITION,
        },
      ],
      spacing: {
        before: point(6),
        after: point(subtitle ? 0.5 : 1.5),
        line: point(11.5),
        lineRule: LineRuleType.EXACT,
      },
      children: headingLineChildren,
    }),
  ];
  if (subtitle) {
    children.push(
      new Paragraph({
        spacing: {
          before: 0,
          after: point(2),
          line: point(10.8),
          lineRule: LineRuleType.EXACT,
        },
        children: [
          new TextRun({
            text: subtitle,
            italics: true,
            color: RESUME_MUTED_COLOR,
            size: halfPoint(10.5),
            font: buildFontAttributes(RESUME_FONT_NAME),
          }),
        ],
      }),
    );
  }

  for (const rawDetail of arrayValue(entry.details)) {
    children.push(buildDetailParagraph(rawDetail));
  }

  return children;
}

function buildDetailParagraph(
  item: unknown,
  options?: { afterPt?: number },
): Paragraph {
  return new Paragraph({
    spacing: {
      before: 0,
      after: point(options?.afterPt ?? 2.5),
    },
    children: normalizeInlineItem(item, RESUME_THEME, {
      fontSize: 10,
      fontName: RESUME_FONT_NAME,
      codeFontName: "Consolas",
    }).children,
  });
}

function buildNameParagraph(name: string): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: {
      before: 0,
      after: point(4),
    },
    children: [
      new TextRun({
        text: name,
        bold: true,
        color: RESUME_BODY_COLOR,
        size: halfPoint(28),
        font: buildFontAttributes(RESUME_FONT_NAME),
      }),
    ],
  });
}

function buildHeadlineParagraph(headline: string): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: {
      before: 0,
      after: 0,
    },
    children: [
      new TextRun({
        text: headline,
        color: RESUME_MUTED_COLOR,
        size: halfPoint(10),
        font: buildFontAttributes(RESUME_FONT_NAME),
      }),
    ],
  });
}

function buildContactParagraph(contactLine: string): Paragraph {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: {
      before: 0,
      after: 0,
    },
    children: [
      new TextRun({
        text: contactLine,
        color: RESUME_BODY_COLOR,
        size: halfPoint(9),
        font: buildFontAttributes(RESUME_FONT_NAME),
      }),
    ],
  });
}

function buildContactDividerParagraph(): Paragraph {
  return new Paragraph({
    spacing: {
      before: point(4),
      after: 0,
    },
    border: {
      bottom: {
        color: RESUME_DIVIDER_COLOR,
        style: BorderStyle.SINGLE,
        size: RESUME_DIVIDER_SIZE,
      },
    },
    children: [new TextRun({ text: "" })],
  });
}

const RESUME_THEME: ThemeConfig = {
  themeName: "technical_resume",
  accent: RESUME_SECTION_COLOR,
  accentSoft: "DCE6F1",
  accentBoxStripColor: RESUME_SECTION_COLOR,
  headingBottomBorderColor: RESUME_DIVIDER_COLOR,
  heroBannerDividerColor: RESUME_DIVIDER_COLOR,
  heroBannerDividerSizePt: 0.75,
  fontName: RESUME_FONT_NAME,
  headingFontName: RESUME_FONT_NAME,
  tableFontName: RESUME_FONT_NAME,
  codeFontName: "Consolas",
  titleSize: 20,
  titleAlign: "center",
  titleSpacingAfter: 0,
  headingSize: 13,
  headingSpaceBefore: 0,
  headingSpaceAfter: 0,
  headingBottomBorder: true,
  headingBottomBorderSizePt: 0.75,
  bodySize: 10,
  bodyIndent: 0,
  bodySpaceAfter: 0,
  bodyLineSpacing: 1.0,
  listSpaceAfter: 0,
  tableFontSize: 10,
  tableStyle: "minimal",
  summaryFill: "FFFFFF",
  margins: { topCm: 1.905, rightCm: 2.2225, bottomCm: 1.905, leftCm: 2.2225 },
};
