import { RenderCliError } from "../../../core/errors";
import { JsonObject } from "../../../core/payload";
import { Block, FileChild, ThemeConfig } from "../types";
import { stringValue } from "../utils";
import { renderBusinessReviewCover } from "./business-review-cover";
import { renderTechnicalResume } from "./technical-resume";

export function renderPageTemplate(
  block: Block,
  metadata: JsonObject,
  theme: ThemeConfig,
): FileChild[] {
  const template = stringValue(block.template);

  switch (template) {
    case "business_review_cover":
      return renderBusinessReviewCover(block, metadata, theme);
    case "technical_resume":
      return renderTechnicalResume(block, metadata, theme);
    default:
      throw new RenderCliError(
        "UNSUPPORTED_PAGE_TEMPLATE",
        `Unsupported page template: ${template || "unknown"}`,
      );
  }
}
