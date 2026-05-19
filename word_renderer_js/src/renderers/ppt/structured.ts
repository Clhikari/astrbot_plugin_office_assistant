import pptxgen from "pptxgenjs";

import { DocumentRenderPayload } from "../../core/payload";
import { renderSlideBlock } from "./blocks";
import { resolveTheme } from "./theme";

export async function renderPptStructuredDocument(
  payload: DocumentRenderPayload,
  outputPath: string,
): Promise<void> {
  const pres = new pptxgen();
  pres.title = (payload.metadata.title as string) || "";
  pres.layout = "LAYOUT_16x9";

  const theme = resolveTheme(payload.metadata.theme_name as string | undefined);

  for (const block of payload.blocks) {
    renderSlideBlock(pres, block, payload.metadata, theme);
  }

  await pres.writeFile({ fileName: outputPath });
}
