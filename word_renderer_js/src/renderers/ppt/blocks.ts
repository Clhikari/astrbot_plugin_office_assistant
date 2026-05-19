import pptxgen from "pptxgenjs";

import { JsonObject } from "../../core/payload";
import { RenderCliError } from "../../core/errors";
import { PptTheme } from "./theme";

export function renderSlideBlock(
  pres: pptxgen,
  block: JsonObject,
  _metadata: JsonObject,
  theme: PptTheme,
): void {
  const blockType = block.type as string;
  switch (blockType) {
    case "title_slide":
      renderTitleSlide(pres, block, theme);
      break;
    case "content_slide":
      renderContentSlide(pres, block, theme);
      break;
    case "table_slide":
      renderTableSlide(pres, block, theme);
      break;
    case "image_slide":
      renderImageSlide(pres, block, theme);
      break;
    default:
      throw new RenderCliError(
        "UNSUPPORTED_BLOCK",
        `Unsupported PPT block type: ${blockType}`,
      );
  }
}

function renderTitleSlide(
  pres: pptxgen,
  block: JsonObject,
  theme: PptTheme,
): void {
  const slide = pres.addSlide();
  slide.background = { color: theme.backgroundColor };

  const title = (block.title as string) || "";
  const subtitle = (block.subtitle as string) || "";

  slide.addText(title, {
    x: 0.5,
    y: 2.0,
    w: 9.0,
    h: 1.2,
    fontSize: 36,
    bold: true,
    color: theme.titleColor,
    fontFace: theme.titleFontFace,
    align: "center",
    valign: "bottom",
  });

  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.5,
      y: 3.3,
      w: 9.0,
      h: 0.8,
      fontSize: 18,
      color: theme.bodyColor,
      fontFace: theme.fontFace,
      align: "center",
      valign: "top",
    });
  }
}

function renderContentSlide(
  pres: pptxgen,
  block: JsonObject,
  theme: PptTheme,
): void {
  const slide = pres.addSlide();
  slide.background = { color: theme.backgroundColor };

  const title = (block.title as string) || "";
  const bullets = (block.bullets as string[]) || [];

  slide.addText(title, {
    x: 0.5,
    y: 0.3,
    w: 9.0,
    h: 0.8,
    fontSize: 24,
    bold: true,
    color: theme.titleColor,
    fontFace: theme.titleFontFace,
    align: "left",
    valign: "bottom",
  });

  const bulletItems = bullets.map((text) => ({
    text,
    options: { bullet: true as const, color: theme.bodyColor },
  }));

  slide.addText(bulletItems, {
    x: 0.5,
    y: 1.3,
    w: 9.0,
    h: 3.8,
    fontSize: 16,
    color: theme.bodyColor,
    fontFace: theme.fontFace,
    valign: "top",
    paraSpaceAfter: 6,
  });
}

function renderTableSlide(
  pres: pptxgen,
  block: JsonObject,
  theme: PptTheme,
): void {
  const slide = pres.addSlide();
  slide.background = { color: theme.backgroundColor };

  const title = (block.title as string) || "";
  const headers = (block.headers as string[]) || [];
  const rows = (block.rows as string[][]) || [];

  let tableY = 0.5;
  if (title) {
    slide.addText(title, {
      x: 0.5,
      y: 0.3,
      w: 9.0,
      h: 0.7,
      fontSize: 22,
      bold: true,
      color: theme.titleColor,
      fontFace: theme.titleFontFace,
      align: "left",
      valign: "bottom",
    });
    tableY = 1.2;
  }

  const headerRow: pptxgen.TableCell[] = headers.map((h) => ({
    text: h,
    options: {
      bold: true,
      color: "FFFFFF",
      fill: { color: theme.accentColor },
      fontSize: 12,
      fontFace: theme.fontFace,
    },
  }));

  const dataRows: pptxgen.TableCell[][] = rows.map((row) =>
    row.map((cell) => ({
      text: cell,
      options: {
        fontSize: 11,
        color: theme.bodyColor,
        fontFace: theme.fontFace,
      },
    })),
  );

  slide.addTable([headerRow, ...dataRows], {
    x: 0.5,
    y: tableY,
    w: 9.0,
    border: { type: "solid", pt: 0.5, color: "D1D5DB" },
    colW: headers.map(() => 9.0 / headers.length),
    autoPage: true,
  });
}

function renderImageSlide(
  pres: pptxgen,
  block: JsonObject,
  theme: PptTheme,
): void {
  const slide = pres.addSlide();
  slide.background = { color: theme.backgroundColor };

  const title = (block.title as string) || "";
  const imagePath = (block.image_path as string) || "";
  const caption = (block.caption as string) || "";

  let imageY = 0.5;
  if (title) {
    slide.addText(title, {
      x: 0.5,
      y: 0.3,
      w: 9.0,
      h: 0.7,
      fontSize: 22,
      bold: true,
      color: theme.titleColor,
      fontFace: theme.titleFontFace,
      align: "left",
      valign: "bottom",
    });
    imageY = 1.2;
  }

  const imageH = caption ? 3.5 : 4.0;
  slide.addImage({
    path: imagePath,
    x: 1.0,
    y: imageY,
    w: 8.0,
    h: imageH,
    sizing: { type: "contain", w: 8.0, h: imageH },
  });

  if (caption) {
    slide.addText(caption, {
      x: 0.5,
      y: imageY + imageH + 0.2,
      w: 9.0,
      h: 0.5,
      fontSize: 12,
      color: theme.bodyColor,
      fontFace: theme.fontFace,
      align: "center",
      italic: true,
    });
  }
}
