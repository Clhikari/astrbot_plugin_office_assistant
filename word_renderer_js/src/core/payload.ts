import fs from "node:fs";

import Ajv2020 from "ajv/dist/2020";

import schema from "../schema/document_payload.schema.json";
import { RenderCliError } from "./errors";

export type JsonObject = Record<string, unknown>;

export interface DocumentRenderPayload extends JsonObject {
  version: "v1";
  render_mode: "structured";
  document_id: string;
  format?: string;
  metadata: JsonObject;
  blocks: Array<JsonObject>;
}

const ajv = new Ajv2020({ allErrors: true, strict: false });
const validatePayload = ajv.compile<DocumentRenderPayload>(schema);

export function readPayload(inputPath: string): DocumentRenderPayload {
  let rawText: string;
  try {
    rawText = fs.readFileSync(inputPath, "utf8");
  } catch (error) {
    throw new RenderCliError(
      "INPUT_READ_FAILED",
      `Cannot read payload: ${inputPath}`,
      error,
    );
  }

  let payload: unknown;
  try {
    payload = JSON.parse(rawText);
  } catch (error) {
    throw new RenderCliError(
      "INPUT_PARSE_FAILED",
      `Payload is not valid JSON: ${inputPath}`,
      error,
    );
  }

  if (!validatePayload(payload)) {
    throw new RenderCliError(
      "SCHEMA_VALIDATION_FAILED",
      ajv.errorsText(validatePayload.errors, { separator: "; " }),
      validatePayload.errors ?? undefined,
    );
  }
  return payload;
}
