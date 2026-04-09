import path from "node:path";

import { RenderCliError } from "./core/errors";
import { readPayload } from "./core/payload";
import { renderDocumentPayload } from "./core/registry";

async function main(): Promise<void> {
  const [, , inputPath, outputPath] = process.argv;
  if (!inputPath || !outputPath) {
    throw new RenderCliError(
      "INVALID_ARGUMENTS",
      "Usage: node dist/cli.js <input_json_path> <output_docx_path>",
    );
  }

  const payload = readPayload(path.resolve(inputPath));
  await renderDocumentPayload(payload, path.resolve(outputPath));
  process.stdout.write(
    JSON.stringify({
      success: true,
      document_id: payload.document_id,
      output_path: path.resolve(outputPath),
    }),
  );
}

main().catch((error: unknown) => {
  const normalized =
    error instanceof RenderCliError
      ? error
      : new RenderCliError(
          "UNEXPECTED_ERROR",
          error instanceof Error ? error.message : String(error),
          error,
        );
  process.stderr.write(
    JSON.stringify({
      success: false,
      code: normalized.code,
      message: normalized.message,
      details: normalized.details,
    }),
  );
  process.exit(1);
});
