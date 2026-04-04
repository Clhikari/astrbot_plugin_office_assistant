import {
  AlignmentType,
  BorderStyle,
  Document,
  Footer,
  Header,
  HeadingLevel,
  Packer,
  PageOrientation,
  Paragraph,
  SectionType,
  ShadingType,
  SimpleField,
  Table,
  TableCell,
  TableLayoutType,
  TableOfContents,
  TableRow,
  TabStopPosition,
  TabStopType,
  TextRun,
  VerticalMergeType,
  WidthType,
  convertInchesToTwip,
} from "docx";

import { RenderCliError } from "../../core/errors";
import { DocumentRenderPayload, JsonObject } from "../../core/payload";

type FileChild = Paragraph | Table | TableOfContents;
type Block = JsonObject & { type: string };
type HeaderFooterConfig = JsonObject;
type ThemeConfig = {
  accent: string;
  accentSoft: string;
  titleSize: number;
  headingSize: number;
  bodySize: number;
  tableStyle: string;
  summaryFill: string;
  margins: {
    topCm: number;
    rightCm: number;
    bottomCm: number;
    leftCm: number;
  };
};
type TableCellValue = {
  text: string;
  rowSpan: number;
  fill?: string;
  textColor?: string;
  bold?: boolean;
  align?: string;
};
type SectionState = {
  children: FileChild[];
  headerFooter: HeaderFooterConfig;
  startType?: keyof typeof SECTION_TYPE_MAP;
  pageOrientation?: string;
  margins?: JsonObject;
  restartPageNumbering: boolean;
  pageNumberStart?: number;
  inheritPreviousHeaderFooter: boolean;
};

const SECTION_TYPE_MAP = {
  new_page: SectionType.NEXT_PAGE,
  continuous: SectionType.CONTINUOUS,
  odd_page: SectionType.ODD_PAGE,
  even_page: SectionType.EVEN_PAGE,
  new_column: SectionType.NEXT_COLUMN,
} as const;

const ORDERED_NUMBERING_REFERENCE = "default-numbering";
const DEFAULT_TABLE_BANDED_ROW_FILL = "F7FBFF";
const DEFAULT_DIVIDER_COLOR = "D0D7DE";
const THEMES: Record<string, ThemeConfig> = {
  business_report: {
    accent: "1F4E79",
    accentSoft: "DCE6F1",
    titleSize: 20,
    headingSize: 14,
    bodySize: 11,
    tableStyle: "report_grid",
    summaryFill: "EEF4FA",
    margins: { topCm: 2.54, rightCm: 2.54, bottomCm: 2.54, leftCm: 2.54 },
  },
  project_review: {
    accent: "0F766E",
    accentSoft: "D9F3EE",
    titleSize: 19,
    headingSize: 13,
    bodySize: 10.5,
    tableStyle: "metrics_compact",
    summaryFill: "E8F6F3",
    margins: { topCm: 2.3, rightCm: 2.3, bottomCm: 2.3, leftCm: 2.3 },
  },
  executive_brief: {
    accent: "B45309",
    accentSoft: "FDEBD8",
    titleSize: 18.5,
    headingSize: 12.5,
    bodySize: 10.5,
    tableStyle: "minimal",
    summaryFill: "FFF7ED",
    margins: { topCm: 2.2, rightCm: 2.2, bottomCm: 2.2, leftCm: 2.2 },
  },
};

export async function renderStructuredDocument(
  payload: DocumentRenderPayload,
  outputPath: string,
): Promise<void> {
  const metadata = asObject(payload.metadata);
  const theme = resolveTheme(metadata);
  const defaultHeaderFooter = asObject(metadata.header_footer);
  const sections: SectionState[] = [
    {
      children: [],
      headerFooter: defaultHeaderFooter,
      restartPageNumbering: false,
      inheritPreviousHeaderFooter: false,
    },
  ];

  const titleParagraph = renderDocumentTitle(metadata, theme);
  if (titleParagraph) {
    sections[0].children.push(titleParagraph);
  }

  let currentHeaderFooter = defaultHeaderFooter;
  for (const rawBlock of payload.blocks) {
    const block = rawBlock as Block;
    if (block.type === "section_break") {
      const inheritHeaderFooter = booleanValue(block.inherit_header_footer) !== false;
      const inheritedHeaderFooter = normalizeInheritedHeaderFooter(currentHeaderFooter);
      const overrideHeaderFooter = asObject(block.header_footer);
      const inheritPreviousHeaderFooter =
        inheritHeaderFooter && !hasHeaderFooterOverride(overrideHeaderFooter);
      const effectiveHeaderFooter = inheritPreviousHeaderFooter
        ? inheritedHeaderFooter
        : inheritHeaderFooter
          ? mergeHeaderFooter(inheritedHeaderFooter, overrideHeaderFooter)
          : overrideHeaderFooter;

      sections.push({
        children: [],
        headerFooter: effectiveHeaderFooter,
        startType: stringValue(block.start_type) as keyof typeof SECTION_TYPE_MAP,
        pageOrientation: stringValue(block.page_orientation) || undefined,
        margins: asObject(block.margins),
        restartPageNumbering: booleanValue(block.restart_page_numbering) === true,
        pageNumberStart: numberValue(block.page_number_start),
        inheritPreviousHeaderFooter,
      });
      currentHeaderFooter = effectiveHeaderFooter;
      continue;
    }

    const section = sections.at(-1);
    if (!section) {
      throw new RenderCliError("SECTION_STATE_INVALID", "Section state is empty");
    }
    section.children.push(...renderBlock(block, metadata, theme));
  }

  const doc = new Document({
    features: {
      updateFields: true,
    },
    evenAndOddHeaderAndFooters: sections.some((section) =>
      usesEvenPageVariants(section.headerFooter),
    ),
    numbering: {
      config: [
        {
          reference: ORDERED_NUMBERING_REFERENCE,
          levels: [
            {
              level: 0,
              format: "decimal",
              text: "%1.",
              alignment: AlignmentType.START,
            },
          ],
        },
      ],
    },
    sections: sections.map((section, index) =>
      buildSection(section, index === 0, theme),
    ),
  });

  const buffer = await Packer.toBuffer(doc);
  await BunLike.writeFile(outputPath, buffer);
}

function buildSection(
  section: SectionState,
  isFirstSection: boolean,
  theme: ThemeConfig,
): {
  properties: JsonObject;
  headers?: { default?: Header; first?: Header; even?: Header };
  footers?: { default?: Footer; first?: Footer; even?: Footer };
  children: FileChild[];
} {
  const properties: JsonObject = {};
  const sectionType = section.startType
    ? SECTION_TYPE_MAP[section.startType] ?? SectionType.NEXT_PAGE
    : undefined;
  if (sectionType) {
    properties.type = sectionType;
  }

  const page: JsonObject = {};
  const orientation = mapPageOrientation(section.pageOrientation);
  if (orientation) {
    page.size = { orientation };
  }
  const margin = buildPageMargins(section.margins, theme);
  if (margin) {
    page.margin = margin;
  }
  if (section.restartPageNumbering) {
    page.pageNumbers = {
      start: section.pageNumberStart ?? 1,
    };
  }
  if (Object.keys(page).length > 0) {
    properties.page = page;
  }
  if (usesFirstPageVariants(section.headerFooter)) {
    properties.titlePage = true;
  }

  return {
    properties,
    headers:
      !section.inheritPreviousHeaderFooter || isFirstSection
        ? buildHeaders(section.headerFooter)
        : undefined,
    footers:
      !section.inheritPreviousHeaderFooter || isFirstSection
        ? buildFooters(section.headerFooter)
        : undefined,
    children: section.children.length > 0 ? section.children : [new Paragraph("")],
  };
}

function renderBlock(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  switch (block.type) {
    case "heading":
      return [renderHeading(block, metadata, theme)];
    case "paragraph":
      return [renderParagraph(block, metadata, theme)];
    case "list":
      return renderList(block, metadata, theme);
    case "table":
      return [renderTable(block, metadata, theme)];
    case "group":
      return arrayValue(block.blocks).flatMap((child) =>
        renderBlock(child as Block, metadata, theme),
      );
    case "columns":
      return renderColumns(block, metadata, theme);
    case "page_break":
      return [new Paragraph({ pageBreakBefore: true, children: [new TextRun("")] })];
    case "toc":
      return [
        new TableOfContents(stringValue(block.title) || "Contents", {
          hyperlink: true,
        }),
      ];
    case "accent_box":
      return [renderAccentBox(block, metadata, theme)];
    case "metric_cards":
      return [renderMetricCards(block, metadata, theme)];
    default:
      throw new RenderCliError(
        "UNSUPPORTED_BLOCK",
        `Unsupported structured block type: ${block.type}`,
      );
  }
}

function renderDocumentTitle(
  metadata: JsonObject,
  theme: ThemeConfig,
): Paragraph | null {
  const title = stringValue(metadata.title);
  if (!title.trim()) {
    return null;
  }
  const documentStyle = asObject(metadata.document_style);
  return new Paragraph({
    alignment: mapAlignment(stringValue(documentStyle.title_align)) ?? AlignmentType.LEFT,
    spacing: {
      after: 220,
    },
    children: [
      new TextRun({
        text: title,
        bold: true,
        color: stringValue(documentStyle.heading_color) || "000000",
        size: halfPoint(theme.titleSize),
      }),
    ],
  });
}

function renderHeading(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Paragraph {
  const documentStyle = asObject(metadata.document_style);
  const level = numberValue(block.level) ?? 1;
  const style = asObject(block.style);
  const layout = asObject(block.layout);
  const color =
    stringValue(block.color) ||
    stringValue(documentStyle[`heading_level_${level}_color`]) ||
    stringValue(documentStyle.heading_color) ||
    "000000";
  const fontScale = numberValue(style.font_scale) ?? 1;
  const baseSize = level <= 1 ? theme.headingSize : Math.max(theme.bodySize + 1, 11.5);
  const border =
    booleanValue(block.bottom_border) === true
      ? {
          bottom: {
            color:
              stringValue(block.bottom_border_color) ||
              stringValue(documentStyle.heading_bottom_border_color) ||
              DEFAULT_DIVIDER_COLOR,
            style: BorderStyle.SINGLE,
            size: Math.max(
              4,
              Math.round(
                (numberValue(block.bottom_border_size_pt) ||
                  numberValue(documentStyle.heading_bottom_border_size_pt) ||
                  0.5) * 8,
              ),
            ),
          },
        }
      : undefined;

  return new Paragraph({
    heading: mapHeadingLevel(level),
    alignment: mapAlignment(stringValue(style.align)) ?? AlignmentType.LEFT,
    border,
    spacing: {
      before: point(numberValue(layout.spacing_before)),
      after: point(numberValue(layout.spacing_after) ?? 10),
    },
    children: [
      new TextRun({
        text: stringValue(block.text),
        bold: true,
        color,
        size: halfPoint(baseSize * fontScale),
      }),
    ],
  });
}

function renderParagraph(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Paragraph {
  const documentStyle = asObject(metadata.document_style);
  const style = asObject(block.style);
  const layout = asObject(block.layout);

  return new Paragraph({
    children: buildRuns(block),
    spacing: {
      before: point(numberValue(layout.spacing_before)),
      after: point(
        numberValue(layout.spacing_after) ??
          numberValue(documentStyle.paragraph_space_after) ??
          6,
      ),
      line: point(
        (numberValue(documentStyle.body_font_size) || theme.bodySize) *
          (numberValue(documentStyle.body_line_spacing) || 1.35),
      ),
    },
    alignment: mapAlignment(stringValue(style.align)),
  });
}

function renderList(
  block: Block,
  metadata: JsonObject,
  _theme: ThemeConfig,
): Paragraph[] {
  const documentStyle = asObject(metadata.document_style);
  const style = asObject(block.style);
  const ordered = booleanValue(block.ordered) === true;

  return arrayValue(block.items).map((item) => {
    const normalized = normalizeInlineItem(item);
    return new Paragraph({
      children: normalized.runs,
      bullet: ordered ? undefined : { level: 0 },
      numbering: ordered
        ? {
            reference: ORDERED_NUMBERING_REFERENCE,
            level: 0,
          }
        : undefined,
      spacing: {
        after: point(numberValue(documentStyle.list_space_after) ?? 4),
      },
      alignment: mapAlignment(stringValue(style.align)),
    });
  });
}

function renderColumns(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  const rendered: FileChild[] = [];
  arrayValue(block.columns).forEach((column, index) => {
    if (index > 0) {
      rendered.push(new Paragraph(""));
    }
    rendered.push(
      ...arrayValue(asObject(column).blocks).flatMap((child) =>
        renderBlock(child as Block, metadata, theme),
      ),
    );
  });
  return rendered;
}

function renderTable(
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
    stringValue(asObject(block.style).cell_align) || stringValue(tableDefaults.cell_align);
  const numericColumns = new Set(
    arrayValue(block.numeric_columns)
      .map((value) => numberValue(value))
      .filter((value): value is number => value !== undefined),
  );

  const tableRows: TableRow[] = [];
  const caption = stringValue(block.caption) || stringValue(block.title);
  if (caption.trim()) {
    tableRows.push(
      new TableRow({
        tableHeader: true,
        cantSplit: true,
        children: [
          new TableCell({
            columnSpan: columnCount,
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: caption,
                    bold: true,
                    color: resolveCaptionColor(block, tableDefaults, theme),
                    size: halfPoint(resolveCaptionFontSize(block, tableDefaults, theme)),
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
    const headerFill = resolveHeaderFill(block, tableDefaults, tableStyleName, theme);
    const groupCells = headerGroups.map((group) => {
      const span = numberValue(group.span) ?? 1;
      spanSum += span;
      return new TableCell({
        columnSpan: span,
        children: [
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({
                text: stringValue(group.title),
                bold: resolveHeaderBold(block),
                color: resolveHeaderTextColor(block, tableDefaults, tableStyleName, theme),
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
        children: headers.map((header) =>
          new TableCell({
            children: [
              new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [
                  new TextRun({
                    text: header,
                    bold: resolveHeaderBold(block),
                    color: resolveHeaderTextColor(block, tableDefaults, tableStyleName, theme),
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
    ),
  );

  const columnWidths = normalizeColumnWidths(block, columnCount);
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    alignment: resolveTableAlignment(block, tableDefaults),
    layout: columnWidths.length > 0 ? TableLayoutType.FIXED : undefined,
    columnWidths: columnWidths.length > 0 ? columnWidths : undefined,
    borders: resolveTableBorders(block, tableDefaults, theme),
    rows: tableRows,
  });
}

function buildTableBodyRows(
  block: Block,
  columnCount: number,
  tableStyleName: string,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
  bodyAlignment: string,
  numericColumns: Set<number>,
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
          children: [
            new Paragraph({
              alignment: mapAlignment(cell.align) ??
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

function renderAccentBox(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): Table {
  const accentColor = stringValue(block.accent_color) || theme.accent;
  const fillColor = stringValue(block.fill_color) || theme.summaryFill;
  const titleColor = stringValue(block.title_color) || accentColor;
  const content: Paragraph[] = [];

  if (stringValue(block.title).trim()) {
    content.push(
      new Paragraph({
        spacing: { after: 80 },
        children: [
          new TextRun({
            text: stringValue(block.title),
            bold: true,
            color: titleColor,
          }),
        ],
      }),
    );
  }

  const items = arrayValue(block.items);
  if (items.length > 0) {
    for (const item of items) {
      const normalized = normalizeInlineItem(item);
      content.push(
        new Paragraph({
          spacing: { after: 40 },
          children: normalized.runs,
        }),
      );
    }
  } else if (stringValue(block.text).trim()) {
    content.push(renderParagraph(block, metadata, theme));
  }

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    rows: [
      new TableRow({
        children: [
          new TableCell({
            children: content.length > 0 ? content : [new Paragraph("")],
            shading: {
              fill: fillColor,
              color: "auto",
              type: ShadingType.CLEAR,
            },
            borders: {
              left: {
                color: accentColor,
                style: BorderStyle.SINGLE,
                size: 18,
              },
            },
          }),
        ],
      }),
    ],
  });
}

function renderMetricCards(
  block: Block,
  _metadata: JsonObject,
  theme: ThemeConfig,
): Table {
  const accentColor = stringValue(block.accent_color) || theme.accent;
  const fillColor = stringValue(block.fill_color) || "F8FAFC";
  const labelColor = stringValue(block.label_color) || "666666";

  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    layout: TableLayoutType.FIXED,
    borders: {
      top: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      bottom: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      left: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      right: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      insideHorizontal: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
      insideVertical: { style: BorderStyle.SINGLE, color: "E5E7EB", size: 4 },
    },
    rows: [
      new TableRow({
        children: arrayValue(block.metrics).map((metric) => {
          const metricObject = asObject(metric);
          const paragraphs: Paragraph[] = [
            new Paragraph({
              children: [
                new TextRun({
                  text: stringValue(metricObject.label),
                  bold: true,
                  color: labelColor,
                }),
              ],
            }),
            new Paragraph({
              spacing: { before: 40, after: 40 },
              children: [
                new TextRun({
                  text: stringValue(metricObject.value),
                  bold: true,
                  color: stringValue(metricObject.value_color) || accentColor,
                  size: 28,
                }),
              ],
            }),
          ];

          if (stringValue(metricObject.delta).trim()) {
            paragraphs.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: stringValue(metricObject.delta),
                    color: stringValue(metricObject.delta_color) || "15803D",
                  }),
                ],
              }),
            );
          }
          if (stringValue(metricObject.note).trim()) {
            paragraphs.push(
              new Paragraph({
                children: [
                  new TextRun({
                    text: stringValue(metricObject.note),
                    color: "666666",
                  }),
                ],
              }),
            );
          }

          return new TableCell({
            children: paragraphs,
            shading: {
              fill: stringValue(metricObject.fill_color) || fillColor,
              color: "auto",
              type: ShadingType.CLEAR,
            },
          });
        }),
      }),
    ],
  });
}

function buildHeaders(
  config: HeaderFooterConfig,
): { default?: Header; first?: Header; even?: Header } {
  const headers: { default?: Header; first?: Header; even?: Header } = {};
  const normalHeader = buildHeaderFooterParagraphs(config, "header", "default");
  if (normalHeader) {
    headers.default = new Header({ children: normalHeader });
  }
  if (usesFirstPageVariants(config)) {
    headers.first = new Header({
      children:
        buildHeaderFooterParagraphs(config, "header", "first") ?? [new Paragraph("")],
    });
  }
  if (usesEvenPageVariants(config)) {
    headers.even = new Header({
      children:
        buildHeaderFooterParagraphs(config, "header", "even") ?? [new Paragraph("")],
    });
  }
  return headers;
}

function buildFooters(
  config: HeaderFooterConfig,
): { default?: Footer; first?: Footer; even?: Footer } {
  const footers: { default?: Footer; first?: Footer; even?: Footer } = {};
  const normalFooter = buildHeaderFooterParagraphs(config, "footer", "default");
  if (normalFooter) {
    footers.default = new Footer({ children: normalFooter });
  }
  if (usesFirstPageVariants(config)) {
    footers.first = new Footer({
      children:
        buildHeaderFooterParagraphs(config, "footer", "first") ?? [new Paragraph("")],
    });
  }
  if (usesEvenPageVariants(config)) {
    footers.even = new Footer({
      children:
        buildHeaderFooterParagraphs(config, "footer", "even") ?? [new Paragraph("")],
    });
  }
  return footers;
}

function buildHeaderFooterParagraphs(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
  variant: "default" | "first" | "even",
): Paragraph[] | null {
  const splitLayout = usesSplitLayout(config, kind);
  let left = resolveHeaderFooterLeft(config, kind, variant);
  let right = resolveHeaderFooterRight(config, kind, variant);
  const showPageNumberSetting = resolveShowPageNumberSetting(config, variant);
  const showPageNumber = showPageNumberSetting === true;

  if (showPageNumberSetting === false) {
    left = suppressPagePlaceholder(left);
    right = suppressPagePlaceholder(right);
  }

  if (showPageNumber && !containsPagePlaceholder(left) && !containsPagePlaceholder(right)) {
    if (kind === "footer") {
      if (splitLayout || left.trim()) {
        right = right.trim() ? `${right} {PAGE}` : "{PAGE}";
      } else {
        right = "{PAGE}";
      }
    }
  }

  if (!(left.trim() || right.trim() || showPageNumber)) {
    return null;
  }

  if (right.trim()) {
    return [
      new Paragraph({
        border: buildHeaderFooterDivider(config, kind),
        tabStops: [
          {
            type: TabStopType.RIGHT,
            position: TabStopPosition.MAX,
          },
        ],
        children: buildSplitHeaderFooterRuns(left, right),
      }),
    ];
  }

  if (kind === "footer" && showPageNumber) {
    if (left.trim()) {
      return [
        new Paragraph({
          border: buildHeaderFooterDivider(config, kind),
          children: buildInlineRuns(left),
        }),
        new Paragraph({
          alignment: resolvePageNumberAlignment(config),
          children: [new SimpleField("PAGE")],
        }),
      ];
    }
    return [
      new Paragraph({
        border: buildHeaderFooterDivider(config, kind),
        alignment: resolvePageNumberAlignment(config),
        children: [new SimpleField("PAGE")],
      }),
    ];
  }

  return [
    new Paragraph({
      border: buildHeaderFooterDivider(config, kind),
      children: buildInlineRuns(left),
    }),
  ];
}

function buildSplitHeaderFooterRuns(left: string, right: string): any[] {
  const children: any[] = [];
  if (left.trim()) {
    children.push(...buildInlineRuns(left));
  }
  if (left.trim() || right.trim()) {
    children.push(new TextRun({ text: "\t" }));
  }
  if (right.trim()) {
    children.push(...buildInlineRuns(right));
  }
  return children.length > 0 ? children : [new TextRun("")];
}

function buildInlineRuns(text: string): any[] {
  if (!text.includes("{PAGE}")) {
    return [new TextRun({ text })];
  }
  const children: any[] = [];
  const parts = text.split("{PAGE}");
  parts.forEach((part, index) => {
    if (part) {
      children.push(new TextRun({ text: part }));
    }
    if (index < parts.length - 1) {
      children.push(new SimpleField("PAGE"));
    }
  });
  return children.length > 0 ? children : [new SimpleField("PAGE")];
}

function suppressPagePlaceholder(text: string): string {
  return containsPagePlaceholder(text) ? "" : text;
}

function normalizeInlineItem(item: unknown): { runs: TextRun[] } {
  if (typeof item === "string") {
    return { runs: [new TextRun(item)] };
  }
  const obj = asObject(item);
  if (arrayValue(obj.runs).length > 0) {
    return { runs: buildRuns(obj) };
  }
  return { runs: [new TextRun(stringValue(obj.text))] };
}

function buildRuns(block: JsonObject): TextRun[] {
  const runs = arrayValue(block.runs);
  if (runs.length === 0) {
    return [new TextRun(stringValue(block.text))];
  }
  return runs.map((rawRun) => {
    const run = asObject(rawRun);
    return new TextRun({
      text: stringValue(run.text),
      bold: booleanValue(run.bold) === true,
      italics: booleanValue(run.italic) === true,
      underline: booleanValue(run.underline) === true ? {} : undefined,
      color: stringValue(run.color) || undefined,
      font: booleanValue(run.code) === true ? "Consolas" : undefined,
    });
  });
}

function normalizeTableCell(cell: unknown): TableCellValue {
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

function resolveTheme(metadata: JsonObject): ThemeConfig {
  const themeName = stringValue(metadata.theme_name);
  const baseTheme = THEMES[themeName] ?? THEMES.business_report;
  const accentColor = normalizeHexColor(stringValue(metadata.accent_color)) || baseTheme.accent;
  return {
    ...baseTheme,
    accent: accentColor,
  };
}

function resolveTableColumnCount(headers: string[], rows: unknown[]): number {
  if (headers.length > 0) {
    return headers.length;
  }
  return rows.reduce<number>(
    (max, row) => Math.max(max, countLogicalColumns(arrayValue(row))),
    0,
  );
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

function resolveHeaderFill(
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

function resolveHeaderTextColor(
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

function resolveHeaderBold(block: Block): boolean {
  const explicit = booleanValue(block.header_bold);
  return explicit === undefined ? true : explicit;
}

function resolveBandedRowFill(
  block: Block,
  tableDefaults: JsonObject,
  tableStyleName: string,
  rowIndex: number,
): string | undefined {
  const bandedRows = booleanValue(block.banded_rows) ?? booleanValue(tableDefaults.banded_rows);
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

function resolveFirstColumnBold(block: Block, tableDefaults: JsonObject): boolean {
  const explicit = booleanValue(block.first_column_bold);
  return explicit === undefined
    ? booleanValue(tableDefaults.first_column_bold) === true
    : explicit;
}

function resolveTableAlignment(
  block: Block,
  tableDefaults: JsonObject,
) {
  const align = stringValue(block.table_align) || stringValue(tableDefaults.table_align);
  if (align === "left") {
    return AlignmentType.LEFT;
  }
  if (align === "center") {
    return AlignmentType.CENTER;
  }
  return undefined;
}

function resolveTableBorders(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
): JsonObject | undefined {
  const borderStyle = stringValue(block.border_style) || stringValue(tableDefaults.border_style);
  const borderMap: Record<string, { size: number; color: string }> = {
    minimal: { size: 4, color: DEFAULT_DIVIDER_COLOR },
    standard: { size: 8, color: "7A7A7A" },
    strong: { size: 16, color: theme.accent },
  };
  const spec = borderMap[borderStyle];
  if (!spec) {
    return undefined;
  }
  return {
    top: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    bottom: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    left: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    right: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    insideHorizontal: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
    insideVertical: { style: BorderStyle.SINGLE, size: spec.size, color: spec.color },
  };
}

function resolveCaptionFill(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
): string {
  const emphasis =
    stringValue(block.caption_emphasis) || stringValue(tableDefaults.caption_emphasis);
  if (emphasis === "strong") {
    return (
      resolveHeaderFill(block, tableDefaults, stringValue(block.table_style) || theme.tableStyle, theme) ||
      theme.accent
    );
  }
  return theme.accentSoft;
}

function resolveCaptionColor(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
): string {
  const emphasis =
    stringValue(block.caption_emphasis) || stringValue(tableDefaults.caption_emphasis);
  if (emphasis === "strong") {
    return (
      normalizeHexColor(stringValue(block.header_text_color)) ||
      normalizeHexColor(stringValue(tableDefaults.header_text_color)) ||
      "FFFFFF"
    );
  }
  return theme.accent;
}

function resolveCaptionFontSize(
  block: Block,
  tableDefaults: JsonObject,
  theme: ThemeConfig,
): number {
  const emphasis =
    stringValue(block.caption_emphasis) || stringValue(tableDefaults.caption_emphasis);
  const baseSize = Math.max(theme.bodySize, 11);
  return emphasis === "strong" ? baseSize + 1 : baseSize;
}

function resolveTableFontSize(
  tableStyleName: string,
  theme: ThemeConfig,
  header: boolean,
): number {
  if (tableStyleName === "metrics_compact") {
    return Math.max(theme.bodySize - 0.5, 9);
  }
  if (tableStyleName === "minimal" && header) {
    return Math.max(theme.bodySize, 10.5);
  }
  return theme.bodySize;
}

function resolveTableBodyAlignment(
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

function normalizeColumnWidths(block: Block, columnCount: number): number[] {
  const widths = arrayValue(block.column_widths)
    .map((value) => numberValue(value))
    .filter((value): value is number => value !== undefined && value > 0)
    .slice(0, columnCount)
    .map((value) => cmToTwip(value));
  return widths.length === columnCount ? widths : [];
}

function buildPageMargins(
  margins: JsonObject | undefined,
  theme: ThemeConfig,
): JsonObject | undefined {
  if (!margins || Object.keys(margins).length === 0) {
    return undefined;
  }
  return {
    top: cmToTwip(numberValue(margins.top_cm) ?? theme.margins.topCm),
    right: cmToTwip(numberValue(margins.right_cm) ?? theme.margins.rightCm),
    bottom: cmToTwip(numberValue(margins.bottom_cm) ?? theme.margins.bottomCm),
    left: cmToTwip(numberValue(margins.left_cm) ?? theme.margins.leftCm),
  };
}

function normalizeInheritedHeaderFooter(config: HeaderFooterConfig): HeaderFooterConfig {
  const next = { ...config };
  if (usesFirstPageVariants(next)) {
    delete next.different_first_page;
    delete next.first_page_header_text;
    delete next.first_page_footer_text;
    delete next.first_page_show_page_number;
  }
  return next;
}

function mergeHeaderFooter(
  baseConfig: HeaderFooterConfig,
  overrideConfig: HeaderFooterConfig,
): HeaderFooterConfig {
  return { ...baseConfig, ...overrideConfig };
}

function hasHeaderFooterOverride(config: HeaderFooterConfig): boolean {
  return Object.keys(config).length > 0;
}

function usesFirstPageVariants(config: HeaderFooterConfig): boolean {
  return (
    booleanValue(config.different_first_page) === true ||
    stringValue(config.first_page_header_text).trim().length > 0 ||
    stringValue(config.first_page_footer_text).trim().length > 0 ||
    booleanValue(config.first_page_show_page_number) !== undefined
  );
}

function usesEvenPageVariants(config: HeaderFooterConfig): boolean {
  return (
    booleanValue(config.different_odd_even) === true ||
    stringValue(config.even_page_header_text).trim().length > 0 ||
    stringValue(config.even_page_footer_text).trim().length > 0 ||
    booleanValue(config.even_page_show_page_number) !== undefined
  );
}

function usesSplitLayout(config: HeaderFooterConfig, kind: "header" | "footer"): boolean {
  return Boolean(
    stringValue(config[`${kind}_left`]).trim() || stringValue(config[`${kind}_right`]).trim(),
  );
}

function resolveHeaderFooterLeft(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
  variant: "default" | "first" | "even",
): string {
  const splitLeft = stringValue(config[`${kind}_left`]);
  if (splitLeft.trim()) {
    return splitLeft;
  }
  if (variant === "first") {
    return stringValue(config[`first_page_${kind}_text`]);
  }
  if (variant === "even") {
    return stringValue(config[`even_page_${kind}_text`]) || stringValue(config[`${kind}_text`]);
  }
  return stringValue(config[`${kind}_text`]);
}

function resolveHeaderFooterRight(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
  _variant: "default" | "first" | "even",
): string {
  return stringValue(config[`${kind}_right`]);
}

function resolveShowPageNumberSetting(
  config: HeaderFooterConfig,
  variant: "default" | "first" | "even",
): boolean | undefined {
  if (variant === "first") {
    return (
      booleanValue(config.first_page_show_page_number) ??
      booleanValue(config.show_page_number)
    );
  }
  if (variant === "even") {
    return (
      booleanValue(config.even_page_show_page_number) ??
      booleanValue(config.show_page_number)
    );
  }
  return booleanValue(config.show_page_number);
}

function resolvePageNumberAlignment(config: HeaderFooterConfig) {
  return mapAlignment(stringValue(config.page_number_align)) ?? AlignmentType.RIGHT;
}

function containsPagePlaceholder(text: string): boolean {
  return text.includes("{PAGE}");
}

function buildHeaderFooterDivider(
  config: HeaderFooterConfig,
  kind: "header" | "footer",
): JsonObject | undefined {
  const enabled =
    kind === "header"
      ? booleanValue(config.header_border_bottom) === true
      : booleanValue(config.footer_border_top) === true;
  if (!enabled) {
    return undefined;
  }
  const color =
    normalizeHexColor(
      stringValue(
        config[
          kind === "header" ? "header_border_color" : "footer_border_color"
        ],
      ),
    ) || DEFAULT_DIVIDER_COLOR;

  return kind === "header"
    ? {
        bottom: {
          color,
          style: BorderStyle.SINGLE,
          size: 4,
        },
      }
    : {
        top: {
          color,
          style: BorderStyle.SINGLE,
          size: 4,
        },
      };
}

function mapHeadingLevel(level: number) {
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

function mapAlignment(value: string | undefined) {
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

function mapPageOrientation(value: string | undefined) {
  if (value === "landscape") {
    return PageOrientation.LANDSCAPE;
  }
  if (value === "portrait") {
    return PageOrientation.PORTRAIT;
  }
  return undefined;
}

function normalizeHexColor(value: string): string | undefined {
  const normalized = value.trim().replace(/^#/, "").toUpperCase();
  if (normalized.length !== 6 || /[^0-9A-F]/.test(normalized)) {
    return undefined;
  }
  return normalized;
}

function point(value: number | undefined): number | undefined {
  return value === undefined ? undefined : Math.round(value * 20);
}

function halfPoint(value: number): number {
  return Math.round(value * 2);
}

function cmToTwip(value: number): number {
  return convertInchesToTwip(value / 2.54);
}

function asObject(value: unknown): JsonObject {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return {};
  }
  return value as JsonObject;
}

function arrayValue(value: unknown): unknown[] {
  return Array.isArray(value) ? value : [];
}

function stringValue(value: unknown): string {
  return typeof value === "string" ? value : "";
}

function numberValue(value: unknown): number | undefined {
  return typeof value === "number" ? value : undefined;
}

function booleanValue(value: unknown): boolean | undefined {
  return typeof value === "boolean" ? value : undefined;
}

const BunLike = {
  async writeFile(path: string, data: Buffer): Promise<void> {
    const fs = await import("node:fs/promises");
    await fs.writeFile(path, data);
  },
};
