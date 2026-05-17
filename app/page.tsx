"use client";

import { ChangeEvent, DragEvent, FormEvent, useMemo, useRef, useState } from "react";

type ConvertState = "idle" | "ready" | "converting" | "done" | "error";

const MAX_FILE_SIZE = 25 * 1024 * 1024;

function formatBytes(bytes: number) {
  if (bytes < 1024 * 1024) {
    return `${Math.max(1, Math.round(bytes / 1024))} KB`;
  }

  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`;
}

export default function Home() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [file, setFile] = useState<File | null>(null);
  const [state, setState] = useState<ConvertState>("idle");
  const [message, setMessage] = useState("");
  const [isDragging, setIsDragging] = useState(false);

  const canConvert = useMemo(() => file && state !== "converting", [file, state]);

  function pickFile(nextFile: File | undefined) {
    if (!nextFile) {
      return;
    }

    if (nextFile.type !== "application/pdf" && !nextFile.name.toLowerCase().endsWith(".pdf")) {
      setState("error");
      setMessage("Choose a PDF file.");
      setFile(null);
      return;
    }

    if (nextFile.size > MAX_FILE_SIZE) {
      setState("error");
      setMessage("Use a PDF smaller than 25 MB.");
      setFile(null);
      return;
    }

    setFile(nextFile);
    setState("ready");
    setMessage("");
  }

  function onInputChange(event: ChangeEvent<HTMLInputElement>) {
    pickFile(event.target.files?.[0]);
  }

  function onDrop(event: DragEvent<HTMLLabelElement>) {
    event.preventDefault();
    setIsDragging(false);
    pickFile(event.dataTransfer.files?.[0]);
  }

  async function convert(event: FormEvent<HTMLFormElement>) {
    event.preventDefault();

    if (!file) {
      return;
    }

    setState("converting");
    setMessage("Reading PDF layout, fonts, and text styles...");

    const formData = new FormData();
    formData.append("file", file);

    try {
      const response = await fetch("/api/convert", {
        method: "POST",
        body: formData,
      });

      if (!response.ok) {
        const error = await response.json().catch(() => null);
        throw new Error(error?.message ?? "Conversion failed.");
      }

      const blob = await response.blob();
      const url = URL.createObjectURL(blob);
      const link = document.createElement("a");
      const outputName = file.name.replace(/\.pdf$/i, "") || "converted";

      link.href = url;
      link.download = `${outputName}.docx`;
      document.body.appendChild(link);
      link.click();
      link.remove();
      URL.revokeObjectURL(url);

      setState("done");
      setMessage("Your styled Word document is ready.");
    } catch (error) {
      setState("error");
      setMessage(error instanceof Error ? error.message : "Conversion failed.");
    }
  }

  return (
    <main className="page">
      <section className="workspace" aria-label="PDF to Word converter">
        <div className="intro">
          <p className="eyebrow">Editable document converter</p>
          <h1>PDF to Word</h1>
          <p className="lede">
            Upload a PDF and get a `.docx` file with editable text and preserved styling.
          </p>
        </div>

        <form className="converter" onSubmit={convert}>
          <input
            ref={inputRef}
            className="sr-only"
            id="pdf-file"
            type="file"
            accept="application/pdf,.pdf"
            onChange={onInputChange}
          />

          <label
            className={`dropzone ${isDragging ? "dragging" : ""}`}
            htmlFor="pdf-file"
            onDragEnter={() => setIsDragging(true)}
            onDragLeave={() => setIsDragging(false)}
            onDragOver={(event) => event.preventDefault()}
            onDrop={onDrop}
          >
            <span className="uploadIcon" aria-hidden="true">
              <svg viewBox="0 0 24 24" role="img">
                <path d="M12 3l4.5 4.5-1.4 1.4-2.1-2.08V15h-2V6.82L8.9 8.9 7.5 7.5 12 3z" />
                <path d="M5 14h2v4h10v-4h2v6H5v-6z" />
              </svg>
            </span>
            <span className="dropTitle">{file ? file.name : "Drop a PDF here"}</span>
            <span className="dropMeta">
              {file ? `${formatBytes(file.size)} selected` : "or click to browse, up to 25 MB"}
            </span>
          </label>

          <div className="actions">
            <button type="button" className="secondary" onClick={() => inputRef.current?.click()}>
              Choose PDF
            </button>
            <button type="submit" disabled={!canConvert}>
              {state === "converting" ? "Converting..." : "Convert to Word"}
            </button>
          </div>

          <div className={`status ${state}`} role="status" aria-live="polite">
            {message || "Ready for a PDF."}
          </div>
        </form>
      </section>
    </main>
  );
}
