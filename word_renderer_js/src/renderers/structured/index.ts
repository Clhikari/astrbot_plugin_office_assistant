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
  renderHeading,
  renderList,
  renderParagraph,
  renderSummaryCard,
  renderToc,
} from "./blocks";
import { renderAccentBox, renderMetricCards } from "./cards";
import { renderHeroBanner } from "./hero-banner";
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
  BunLike,
  arrayValue,
  asObject,
  booleanValue,
  mapPageOrientation,
  numberValue,
  stringValue,
} from "./utils";

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

  appendBlocksToSections(
    payload.blocks.map((block) => block as Block),
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
  await BunLike.writeFile(outputPath, buffer);
}

function appendBlocksToSections(
  blocks: Block[],
  metadata: JsonObject,
  theme: ThemeConfig,
  sections: SectionState[],
  currentHeaderFooter: HeaderFooterConfig,
): HeaderFooterConfig {
  let activeHeaderFooter = currentHeaderFooter;

  for (const block of blocks) {
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
          section.children.push(new Paragraph(""));
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
  }

  return activeHeaderFooter;
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
      : overrideHeaderFooter;

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
    case "hero_banner":
      return [renderHeroBanner(block, theme)];
    case "heading":
      return [renderHeading(block, metadata, theme)];
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
