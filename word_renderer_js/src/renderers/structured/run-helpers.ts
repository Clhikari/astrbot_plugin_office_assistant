import { TextRun } from "docx";

import { RenderCliError } from "../../core/errors";
import { readSharedContract } from "../../core/shared-contracts";

type HyperlinkUrlContract = {
  allowed_schemes: string[];
  schemes_requiring_authority: string[];
  schemes_requiring_path: string[];
  error_message: string;
};

const HYPERLINK_URL_CONTRACT =
  readSharedContract<HyperlinkUrlContract>("hyperlink_url.json");
const SUPPORTED_HYPERLINK_PROTOCOLS = new Set(
  HYPERLINK_URL_CONTRACT.allowed_schemes.map((scheme) => `${scheme}:`),
);
const HYPERLINK_PROTOCOLS_REQUIRING_HOST = new Set(
  HYPERLINK_URL_CONTRACT.schemes_requiring_authority.map(
    (scheme) => `${scheme}:`,
  ),
);
const HYPERLINK_PROTOCOLS_REQUIRING_PATH = new Set(
  HYPERLINK_URL_CONTRACT.schemes_requiring_path.map(
    (scheme) => `${scheme}:`,
  ),
);
const HYPERLINK_URL_INVALID_CODE = "HYPERLINK_URL_INVALID";
const HYPERLINK_URL_INVALID_MESSAGE = `Hyperlink ${HYPERLINK_URL_CONTRACT.error_message}`;

export const DEFAULT_HYPERLINK_COLOR = "0563C1";

export function normalizeLineBreaks(text: string): string {
  return text.replace(/\\n/g, "\n").replace(/\r\n/g, "\n");
}

export function buildTextRuns(
  text: string,
  options: Record<string, unknown>,
): TextRun[] {
  const normalizedText = normalizeLineBreaks(text);
  const segments = normalizedText.split("\n");
  if (segments.length === 1) {
    return [new TextRun({ ...options, text: normalizedText })];
  }
  return segments.map((segment, index) =>
    new TextRun({
      ...options,
      text: segment,
      break: index === 0 ? undefined : 1,
    }),
  );
}

function throwInvalidHyperlinkUrl(target: string): never {
  throw new RenderCliError(
    HYPERLINK_URL_INVALID_CODE,
    `${HYPERLINK_URL_INVALID_MESSAGE}: ${target}`,
  );
}

function formatHyperlinkTargetForError(value: unknown): string {
  if (typeof value === "string") {
    return value;
  }
  if (
    typeof value === "number" ||
    typeof value === "boolean" ||
    typeof value === "bigint"
  ) {
    return String(value);
  }
  if (value === null) {
    return "null";
  }
  try {
    const serialized = JSON.stringify(value);
    if (typeof serialized === "string") {
      return serialized;
    }
  } catch {
    // Fall through to String(value) when JSON serialization fails.
  }
  return String(value);
}

export function normalizeHyperlinkTarget(value: unknown): string | undefined {
  if (value === undefined) {
    return undefined;
  }
  if (typeof value !== "string") {
    throwInvalidHyperlinkUrl(formatHyperlinkTargetForError(value));
  }

  const target = value.trim();
  if (!target) {
    return undefined;
  }

  let parsed: URL;
  try {
    parsed = new URL(target);
  } catch {
    throwInvalidHyperlinkUrl(target);
  }

  if (!SUPPORTED_HYPERLINK_PROTOCOLS.has(parsed.protocol)) {
    throwInvalidHyperlinkUrl(target);
  }

  if (HYPERLINK_PROTOCOLS_REQUIRING_HOST.has(parsed.protocol) && !parsed.host) {
    throwInvalidHyperlinkUrl(target);
  }

  if (
    HYPERLINK_PROTOCOLS_REQUIRING_PATH.has(parsed.protocol) &&
    !parsed.pathname
  ) {
    throwInvalidHyperlinkUrl(target);
  }

  return parsed.toString();
}
