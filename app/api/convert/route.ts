import { Document, Packer, PageBreak, Paragraph, TextRun } from "docx";
import pdfParse from "pdf-parse";

export const runtime = "nodejs";
export const maxDuration = 60;

const MAX_FILE_SIZE = 25 * 1024 * 1024;
const PAGE_MARKER_START = "__PDF_TO_WORD_PAGE_START__";
const PAGE_MARKER_END = "__PDF_TO_WORD_PAGE_END__";

type PdfTextStyle = {
  fontFamily?: string;
};

type PdfTextItem = {
  str: string;
  transform?: number[];
  width?: number;
  height?: number;
  fontName?: string;
};

type PdfTextContent = {
  items: PdfTextItem[];
  styles?: Record<string, PdfTextStyle>;
};

type PdfPageData = {
  pageNumber?: number;
  getTextContent: (options?: {
    normalizeWhitespace?: boolean;
    disableCombineTextItems?: boolean;
  }) => Promise<PdfTextContent>;
};

type StyledRun = {
  text: string;
  x: number;
  width: number;
  font?: string;
  size: number;
  bold: boolean;
  italics: boolean;
};

type StyledLine = {
  y: number;
  x: number;
  runs: StyledRun[];
};

type StyledPage = {
  pageNumber: number;
  lines: StyledLine[];
};

function sanitizeFileName(name: string) {
  return name.replace(/[^\w.-]+/g, "_").replace(/\.pdf$/i, "") || "converted";
}

function normalizeFontName(fontName: string | undefined) {
  if (!fontName) {
    return undefined;
  }

  return fontName
    .replace(/^[A-Z]{6}\+/, "")
    .replace(/[-_](Bold|Italic|Oblique|Regular|Medium|Light|SemiBold).*$/i, "")
    .replace(/[-_]+/g, " ")
    .trim();
}

function isBoldFont(fontName: string | undefined) {
  return /bold|black|heavy|semibold|demibold/i.test(fontName ?? "");
}

function isItalicFont(fontName: string | undefined) {
  return /italic|oblique/i.test(fontName ?? "");
}

function getItemX(item: PdfTextItem) {
  return item.transform?.[4] ?? 0;
}

function getItemY(item: PdfTextItem) {
  return item.transform?.[5] ?? 0;
}

function getItemSize(item: PdfTextItem) {
  const [, b = 0, , d = item.height ?? 12] = item.transform ?? [];
  const size = Math.hypot(b, d) || item.height || 12;

  return Math.max(7, Math.min(72, Math.round(size)));
}

function getWordSize(points: number) {
  return Math.max(14, Math.min(144, Math.round(points * 2)));
}

function getLineTolerance(size: number) {
  return Math.max(3, size * 0.45);
}

async function renderStyledPage(pageData: unknown) {
  const page = pageData as PdfPageData;
  const content = await page.getTextContent({
    normalizeWhitespace: false,
    disableCombineTextItems: false,
  });

  const lines: StyledLine[] = [];
  const sortedItems = content.items
    .filter((item) => item.str.trim())
    .sort((a, b) => getItemY(b) - getItemY(a) || getItemX(a) - getItemX(b));

  for (const item of sortedItems) {
    const text = item.str.replace(/\s+/g, " ");

    if (!text.trim()) {
      continue;
    }

    const x = getItemX(item);
    const y = getItemY(item);
    const size = getItemSize(item);
    const styleFont = item.fontName ? content.styles?.[item.fontName]?.fontFamily : undefined;
    const rawFont = styleFont || item.fontName;
    const font = normalizeFontName(rawFont);
    const existingLine = lines.find((line) => Math.abs(line.y - y) <= getLineTolerance(size));

    const run: StyledRun = {
      text,
      x,
      width: item.width ?? 0,
      font,
      size,
      bold: isBoldFont(rawFont),
      italics: isItalicFont(rawFont),
    };

    if (existingLine) {
      existingLine.x = Math.min(existingLine.x, x);
      existingLine.runs.push(run);
      continue;
    }

    lines.push({ y, x, runs: [run] });
  }

  for (const line of lines) {
    line.runs = line.runs
      .filter((run) => run.text.trim())
      .sort((a, b) => a.x - b.x)
      .map((run, index, runs) => {
        if (index === 0 || /^\s/.test(run.text) || /\s$/.test(runs[index - 1].text)) {
          return run;
        }

        const previous = runs[index - 1];
        const gap = run.x - (previous.x + previous.width);

        if (gap <= Math.max(1, run.size * 0.2)) {
          return run;
        }

        return {
          ...run,
          text: ` ${run.text}`,
        };
      });
  }

  const styledPage: StyledPage = {
    pageNumber: page.pageNumber ?? 0,
    lines,
  };

  return `${PAGE_MARKER_START}${JSON.stringify(styledPage)}${PAGE_MARKER_END}`;
}

function parseStyledPages(text: string) {
  const pages: StyledPage[] = [];
  const pattern = new RegExp(`${PAGE_MARKER_START}([\\s\\S]*?)${PAGE_MARKER_END}`, "g");
  let match = pattern.exec(text);

  while (match) {
    try {
      pages.push(JSON.parse(match[1]) as StyledPage);
    } catch {
      // Ignore a malformed page payload and let the plain-text fallback handle it.
    }

    match = pattern.exec(text);
  }

  return pages;
}

function buildStyledParagraphs(pages: StyledPage[], fallbackText: string) {
  if (!pages.length) {
    return buildFallbackParagraphs(fallbackText);
  }

  const paragraphs: Paragraph[] = [];

  pages.forEach((page, pageIndex) => {
    if (pageIndex > 0) {
      paragraphs.push(
        new Paragraph({
          children: [new PageBreak()],
        }),
      );
    }

    if (!page.lines.length) {
      paragraphs.push(
        new Paragraph({
          spacing: { after: 180 },
          children: [new TextRun("No selectable text was found on this page.")],
        }),
      );
      return;
    }

    page.lines.forEach((line) => {
      const averageSize =
        line.runs.reduce((total, run) => total + run.size, 0) / Math.max(1, line.runs.length);

      paragraphs.push(
        new Paragraph({
          spacing: { after: Math.max(80, Math.round(averageSize * 8)) },
          indent: { left: Math.max(0, Math.round(line.x * 10)) },
          children: line.runs.map(
            (run) =>
              new TextRun({
                text: run.text,
                font: run.font,
                size: getWordSize(run.size),
                bold: run.bold,
                italics: run.italics,
              }),
          ),
        }),
      );
    });
  });

  return paragraphs;
}

function buildFallbackParagraphs(text: string) {
  const normalized = text.replace(/\r\n/g, "\n").replace(/\r/g, "\n").trim();

  if (!normalized) {
    return [
      new Paragraph({
        children: [new TextRun("No selectable text was found in this PDF.")],
      }),
    ];
  }

  return normalized.split(/\n{2,}/).map((block) => {
    const lines = block
      .split("\n")
      .map((line) => line.trim())
      .filter(Boolean);

    return new Paragraph({
      spacing: { after: 180 },
      children: [new TextRun(lines.join(" "))],
    });
  });
}

export async function POST(request: Request) {
  const formData = await request.formData();
  const file = formData.get("file");

  if (!(file instanceof File)) {
    return Response.json({ message: "Upload a PDF file." }, { status: 400 });
  }

  if (file.size > MAX_FILE_SIZE) {
    return Response.json({ message: "Use a PDF smaller than 25 MB." }, { status: 400 });
  }

  if (file.type !== "application/pdf" && !file.name.toLowerCase().endsWith(".pdf")) {
    return Response.json({ message: "Only PDF files can be converted." }, { status: 400 });
  }

  const arrayBuffer = await file.arrayBuffer();
  const pdfBuffer = Buffer.from(arrayBuffer);
  const parsed = await pdfParse(pdfBuffer, {
    pagerender: renderStyledPage,
  });
  const pages = parseStyledPages(parsed.text);

  const doc = new Document({
    creator: "PDF to Word",
    title: sanitizeFileName(file.name),
    sections: [
      {
        properties: {},
        children: buildStyledParagraphs(pages, parsed.text),
      },
    ],
  });

  const output = await Packer.toBuffer(doc);
  const body = new Uint8Array(output);
  const outputName = `${sanitizeFileName(file.name)}.docx`;

  return new Response(body, {
    headers: {
      "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition": `attachment; filename="${outputName}"`,
    },
  });
}
