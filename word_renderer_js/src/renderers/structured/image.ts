import fs from "node:fs";
import path from "node:path";

import { AlignmentType, ImageRun, Paragraph, TextRun } from "docx";
import { imageSize } from "image-size";

import { RenderCliError } from "../../core/errors";
import { Block, FileChild, ThemeConfig } from "./types";

export function renderImageBlock(
  block: Block,
  theme: ThemeConfig,
  workspaceDir: string,
): FileChild[] {
  const imagePath = (block.path as string) || "";
  const caption = (block.caption as string) || "";
  const widthPx = (block.width_px as number) || 0;

  if (!imagePath) {
    throw new RenderCliError(
      "MISSING_IMAGE_PATH",
      "image block requires a non-empty path",
    );
  }

  if (!imagePath.startsWith("images/")) {
    throw new RenderCliError(
      "INVALID_IMAGE_PATH",
      `image block path must start with "images/", got: ${imagePath}`,
    );
  }

  const segments = imagePath.replace(/\\/g, "/").split("/");
  if (segments.includes("..")) {
    throw new RenderCliError(
      "INVALID_IMAGE_PATH",
      `image block path must not contain directory traversal (..): ${imagePath}`,
    );
  }

  const resolvedPath = path.resolve(workspaceDir, imagePath);
  if (!fs.existsSync(resolvedPath)) {
    throw new RenderCliError(
      "MISSING_IMAGE_FILE",
      `Image file does not exist: ${resolvedPath}`,
    );
  }

  const imageData = fs.readFileSync(resolvedPath);

  const maxWidthPx = 580;
  let displayWidth: number;
  let displayHeight: number;

  const dimensions = readImageDimensions(imageData, imagePath);
  const naturalWidth = dimensions.width || 580;
  const naturalHeight = dimensions.height || 580;

  type DocxImageType = "png" | "jpg" | "gif" | "bmp";
  const typeMap: Record<string, DocxImageType> = {
    png: "png",
    jpg: "jpg",
    gif: "gif",
    bmp: "bmp",
  };
  const docxType = dimensions.type ? typeMap[dimensions.type] : undefined;
  if (!docxType) {
    throw new RenderCliError(
      "UNSUPPORTED_IMAGE_TYPE",
      `Unsupported image type for embedding: ${dimensions.type ?? "unknown"} (${imagePath}). Supported: png, jpg, gif, bmp.`,
    );
  }

  if (widthPx && widthPx > 0) {
    displayWidth = Math.min(widthPx, maxWidthPx);
  } else {
    displayWidth = Math.min(naturalWidth, maxWidthPx);
  }
  const scale = displayWidth / naturalWidth;
  displayHeight = Math.round(naturalHeight * scale);

  const elements: FileChild[] = [];

  elements.push(
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new ImageRun({
          type: docxType,
          data: imageData,
          transformation: {
            width: displayWidth,
            height: displayHeight,
          },
          altText: {
            title: caption || path.basename(imagePath),
            description: caption || imagePath,
            name: path.basename(imagePath),
          },
        }),
      ],
    }),
  );

  if (caption) {
    elements.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 80 },
        children: [
          new TextRun({
            text: caption,
            italics: true,
            size: theme.bodySize - 2,
            font: theme.fontName,
          }),
        ],
      }),
    );
  }

  return elements;
}

function readImageDimensions(imageData: Uint8Array, imagePath: string) {
  try {
    return imageSize(imageData);
  } catch (error) {
    const reason = error instanceof Error ? error.message : String(error);
    throw new RenderCliError(
      "UNSUPPORTED_IMAGE_TYPE",
      `Unsupported image data for embedding: ${imagePath}. ${reason}`,
    );
  }
}
