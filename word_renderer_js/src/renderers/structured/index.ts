import {
  AlignmentType,
  Document,
  Footer,
  Header,
  Packer,
  PageBreak,
  Paragraph,
  SectionType,
} from "docx";

import { RenderCliError } from "../../core/errors";
import { DocumentRenderPayload, JsonObject } from "../../core/payload";
import {
  renderDocumentTitle,
  renderHeadingBlock,
  renderList,
  renderParagraph,
  renderSummaryCard,
  renderToc,
} from "./blocks";
import { renderAccentBox, renderMetricCards } from "./cards";
import { renderHeroBannerBlock } from "./hero-banner";
import {
  EXTERNAL_WORD_STYLES_XML,
  ORDERED_NUMBERING_REFERENCE,
  SECTION_TYPE_MAP,
} from "./constants";
import {
  buildFooters,
  buildHeaders,
  hasHeaderFooterOverride,
  mergeHeaderFooter,
  normalizeInheritedHeaderFooter,
  usesEvenPageVariants,
  usesFirstPageVariants,
} from "./header-footer";
import { createSpacingParagraph } from "./layout-spacing";
import {
  renderPageTemplate,
} from "./page-templates";
import { renderTable } from "./table";
import { buildPageMargins, resolveTheme } from "./theme";
import {
  Block,
  FileChild,
  HeaderFooterConfig,
  SectionState,
  ThemeConfig,
} from "./types";
import {
  arrayValue,
  asObject,
  booleanValue,
  mapPageOrientation,
  numberValue,
  stringValue,
  writeBufferToFile,
} from "./utils";
import { buildBusinessReviewFooterNote } from "./page-templates/business-review-cover";

export async function renderStructuredDocument(
  payload: DocumentRenderPayload,
  outputPath: string,
): Promise<void> {
  const metadata = asObject(payload.metadata);
  const theme = resolveTheme(metadata);
  const structuredBlocks = payload.blocks.map((block) => block as Block);
  const defaultHeaderFooter = asObject(metadata.header_footer);
  const defaultMargins = resolveDefaultSectionMargins(structuredBlocks);
  const sections: SectionState[] = [
    {
      children: [],
      headerFooter: defaultHeaderFooter,
      margins: defaultMargins,
      restartPageNumbering: false,
      inheritPreviousHeaderFooter: false,
    },
  ];

  const titleParagraph = shouldRenderDocumentTitle(structuredBlocks)
    ? renderDocumentTitle(metadata, theme)
    : null;
  if (titleParagraph) {
    sections[0].children.push(titleParagraph);
  }

  appendBlocksToSections(
    structuredBlocks,
    metadata,
    theme,
    sections,
    defaultHeaderFooter,
  );

  const doc = new Document({
    externalStyles: EXTERNAL_WORD_STYLES_XML,
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
  await writeBufferToFile(outputPath, buffer);
}

function resolveDefaultSectionMargins(blocks: Block[]): JsonObject | undefined {
  const firstBlock = blocks[0];
  if (
    firstBlock?.type === "page_template" &&
    stringValue(firstBlock.template) === "technical_resume"
  ) {
    return {
      top_cm: 1.905,
      right_cm: 2.2225,
      bottom_cm: 1.905,
      left_cm: 2.2225,
    };
  }
  return undefined;
}

function shouldRenderDocumentTitle(blocks: Block[]): boolean {
  const firstBlockType = blocks[0]?.type;
  return firstBlockType !== "page_template" && firstBlockType !== "hero_banner";
}

function appendBlocksToSections(
  blocks: Block[],
  metadata: JsonObject,
  theme: ThemeConfig,
  sections: SectionState[],
  currentHeaderFooter: HeaderFooterConfig,
): HeaderFooterConfig {
  let activeHeaderFooter = currentHeaderFooter;
  const deferredTrailingChildren: FileChild[] = [];

  for (const [index, rawBlock] of blocks.entries()) {
    const pageTemplateOverride =
      rawBlock.type === "page_template" && index < blocks.length - 1
        ? prepareTrailingPageTemplate(rawBlock, theme, deferredTrailingChildren)
        : null;
    const block = pageTemplateOverride ?? rawBlock;

    if (block.type === "section_break") {
      activeHeaderFooter = appendSectionBreak(
        block,
        sections,
        activeHeaderFooter,
      );
      continue;
    }

    if (block.type === "group") {
      activeHeaderFooter = appendBlocksToSections(
        arrayValue(block.blocks).map((child) => child as Block),
        metadata,
        theme,
        sections,
        activeHeaderFooter,
      );
      continue;
    }

    if (block.type === "columns") {
      const columns = arrayValue(block.columns).map((column) => asObject(column));
      columns.forEach((column, index) => {
        if (index > 0) {
          const section = sections.at(-1);
          if (!section) {
            throw new RenderCliError(
              "SECTION_STATE_INVALID",
              "Section state is empty",
            );
          }
          section.children.push(createSpacingParagraph({ afterPt: 3 }));
        }

        activeHeaderFooter = appendBlocksToSections(
          arrayValue(column.blocks).map((child) => child as Block),
          metadata,
          theme,
          sections,
          activeHeaderFooter,
        );
      });
      continue;
    }

    const section = sections.at(-1);
    if (!section) {
      throw new RenderCliError(
        "SECTION_STATE_INVALID",
        "Section state is empty",
      );
    }
    section.children.push(...renderBlock(block, metadata, theme));
    const interBlockSpacingAfterPt = resolveInterBlockSpacingAfterPt(block, metadata);
    if (interBlockSpacingAfterPt > 0) {
      section.children.push(createSpacingParagraph({ afterPt: interBlockSpacingAfterPt }));
    }
  }

  if (deferredTrailingChildren.length > 0) {
    const lastSection = sections.at(-1);
    if (!lastSection) {
      throw new RenderCliError(
        "SECTION_STATE_INVALID",
        "Section state is empty",
      );
    }
    lastSection.children.push(...deferredTrailingChildren);
  }

  return activeHeaderFooter;
}

function resolveInterBlockSpacingAfterPt(block: Block, metadata: JsonObject): number {
  const layout = asObject(block.layout);
  const explicitSpacingAfter = numberValue(layout.spacing_after);
  if (explicitSpacingAfter !== undefined) {
    return explicitSpacingAfter;
  }
  if (stringValue(metadata.theme_name) !== "business_report") {
    return 0;
  }
  switch (block.type) {
    case "hero_banner":
      return 6;
    case "accent_box":
      return 5;
    case "metric_cards":
      return 7;
    case "table":
      return 6;
    default:
      return 0;
  }
}

function prepareTrailingPageTemplate(
  block: Block,
  theme: ThemeConfig,
  deferredTrailingChildren: FileChild[],
): Block | null {
  if (block.type !== "page_template") {
    return null;
  }
  const data = asObject(block.data);
  const templateName = stringValue(block.template);
  const footerNote = stringValue(data.footer_note).trim();
  if (templateName === "business_review_cover" && footerNote) {
    deferredTrailingChildren.push(...buildBusinessReviewFooterNote(footerNote, theme));
  }

  return {
    ...block,
    data: {
      ...data,
      auto_page_break: false,
      footer_note: "",
    },
  } as Block;
}

function appendSectionBreak(
  block: Block,
  sections: SectionState[],
  currentHeaderFooter: HeaderFooterConfig,
): HeaderFooterConfig {
  const inheritHeaderFooter = booleanValue(block.inherit_header_footer) !== false;
  const inheritedHeaderFooter = normalizeInheritedHeaderFooter(
    currentHeaderFooter,
  );
  const overrideHeaderFooter = asObject(block.header_footer);
  const inheritPreviousHeaderFooter =
    inheritHeaderFooter && !hasHeaderFooterOverride(overrideHeaderFooter);
  const effectiveHeaderFooter = inheritPreviousHeaderFooter
    ? inheritedHeaderFooter
    : inheritHeaderFooter
      ? mergeHeaderFooter(inheritedHeaderFooter, overrideHeaderFooter)
      : hasHeaderFooterOverride(overrideHeaderFooter)
        ? overrideHeaderFooter
        : buildClearedHeaderFooterOverride(inheritedHeaderFooter);

  sections.push({
    children: [],
    headerFooter: effectiveHeaderFooter,
    startType: stringValue(block.start_type) as SectionState["startType"],
    pageOrientation: stringValue(block.page_orientation) || undefined,
    margins: asObject(block.margins),
    restartPageNumbering: booleanValue(block.restart_page_numbering) === true,
    pageNumberStart: numberValue(block.page_number_start),
    inheritPreviousHeaderFooter,
  });

  return effectiveHeaderFooter;
}

function buildClearedHeaderFooterOverride(
  inheritedHeaderFooter: HeaderFooterConfig,
): HeaderFooterConfig {
  const cleared: HeaderFooterConfig = {
    header_text: "",
    footer_text: "",
    show_page_number: false,
  };
  if (usesFirstPageVariants(inheritedHeaderFooter)) {
    cleared.different_first_page = true;
    cleared.first_page_header_text = "";
    cleared.first_page_footer_text = "";
    cleared.first_page_show_page_number = false;
  }
  if (usesEvenPageVariants(inheritedHeaderFooter)) {
    cleared.different_odd_even = true;
    cleared.even_page_header_text = "";
    cleared.even_page_footer_text = "";
    cleared.even_page_show_page_number = false;
  }
  return cleared;
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

  const page: JsonObject = {
    margin: buildPageMargins(section.margins, theme),
  };
  const orientation = mapPageOrientation(section.pageOrientation);
  if (orientation) {
    page.size = { orientation };
  }
  if (section.restartPageNumbering) {
    page.pageNumbers = {
      start: section.pageNumberStart ?? 1,
    };
  }
  properties.page = page;

  if (usesFirstPageVariants(section.headerFooter)) {
    properties.titlePage = true;
  }

  return {
    properties,
    headers:
      !section.inheritPreviousHeaderFooter || isFirstSection
        ? buildHeaders(section.headerFooter, theme)
        : undefined,
    footers:
      !section.inheritPreviousHeaderFooter || isFirstSection
        ? buildFooters(section.headerFooter, theme)
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
    case "page_template":
      return renderPageTemplate(block, metadata, theme);
    case "hero_banner":
      return renderHeroBannerBlock(block, theme);
    case "heading":
      return renderHeadingBlock(block, metadata, theme);
    case "paragraph":
      return renderParagraph(block, metadata, theme);
    case "list":
      return renderList(block, metadata, theme);
    case "table":
      return [renderTable(block, metadata, theme)];
    case "page_break":
      return [new Paragraph({ children: [new PageBreak()] })];
    case "toc":
      return renderToc(block, metadata, theme);
    case "accent_box":
      return [renderAccentBox(block, metadata, theme)];
    case "metric_cards":
      return [renderMetricCards(block, metadata, theme)];
    case "summary_card":
      return renderSummaryCard(block, metadata, theme);
    default:
      throw new RenderCliError(
        "UNSUPPORTED_BLOCK",
        `Unsupported structured block type: ${block.type}`,
      );
  }
}
