import { RenderCliError } from "../../core/errors";
import { DocumentRenderPayload } from "../../core/payload";

export async function renderPptStructuredDocument(
  _payload: DocumentRenderPayload,
  _outputPath: string,
): Promise<void> {
  throw new RenderCliError(
    "FORMAT_NOT_IMPLEMENTED",
    "PPT structured renderer is reserved for the JS pipeline but is not implemented yet",
  );
}
