import { LineRuleType, Paragraph, TextRun } from "docx";

import { point } from "./utils";

export function createSpacingParagraph({
  beforePt = 0,
  afterPt = 0,
}: {
  beforePt?: number;
  afterPt?: number;
}): Paragraph {
  return new Paragraph({
    spacing: {
      before: point(beforePt),
      after: point(afterPt),
      line: 1,
      lineRule: LineRuleType.EXACT,
    },
    children: [
      new TextRun({
        text: " ",
        size: 2,
      }),
    ],
  });
}
