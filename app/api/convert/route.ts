import { Document, Packer, Paragraph, TextRun } from "docx";
import pdfParse from "pdf-parse";

export const runtime = "nodejs";
export const maxDuration = 60;

const MAX_FILE_SIZE = 25 * 1024 * 1024;

function sanitizeFileName(name: string) {
  return name.replace(/[^\w.-]+/g, "_").replace(/\.pdf$/i, "") || "converted";
}

function buildParagraphs(text: string) {
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
  const parsed = await pdfParse(pdfBuffer);

  const doc = new Document({
    creator: "PDF to Word",
    title: sanitizeFileName(file.name),
    sections: [
      {
        properties: {},
        children: buildParagraphs(parsed.text),
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
