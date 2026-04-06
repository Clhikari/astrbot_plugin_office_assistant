import { RenderCliError } from "./errors";
import { DocumentRenderPayload } from "./payload";
import { renderPptStructuredDocument } from "../renderers/ppt/structured";
import { renderWordStructuredDocument } from "../renderers/word/structured";

type Renderer = (payload: DocumentRenderPayload, outputPath: string) => Promise<void>;

const RENDERERS: Record<string, Record<string, Renderer>> = {
  word: {
    structured: renderWordStructuredDocument,
  },
  ppt: {
    structured: renderPptStructuredDocument,
  },
};

function normalizeDocumentFormat(payload: DocumentRenderPayload): string {
  const candidate =
    typeof payload.format === "string" && payload.format.trim()
      ? payload.format.trim().toLowerCase()
      : "word";
  return candidate;
}

export async function renderDocumentPayload(
  payload: DocumentRenderPayload,
  outputPath: string,
): Promise<void> {
  const documentFormat = normalizeDocumentFormat(payload);
  const formatRenderers = RENDERERS[documentFormat];
  if (!formatRenderers) {
    throw new RenderCliError(
      "UNSUPPORTED_FORMAT",
      `Unsupported document format: ${documentFormat}`,
    );
  }

  const renderer = formatRenderers[payload.render_mode];
  if (!renderer) {
    throw new RenderCliError(
      "UNSUPPORTED_RENDER_MODE",
      `Unsupported render mode for ${documentFormat}: ${payload.render_mode}`,
    );
  }

  await renderer(payload, outputPath);
}
