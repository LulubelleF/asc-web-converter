import React, { useCallback, useEffect, useRef, useState } from "react";
import styles from "../styles/Converter.module.css";

import { saveAs } from "file-saver";
import { createWorkbookFromPages } from "../utils/makeExcel";
import {
  parseDisbursementPageFromText,
  parseDisbursementPageFromCanvas,
} from "../utils/parse";

// The structure you push into `pagesOut`
type PageResult = any;

export default function Home() {
  const pdfjsRef = useRef<any>(null);

  const [pdfReady, setPdfReady] = useState(false);
  const [working, setWorking] = useState({
    loading: false,
    ocr: false,
    parse: false,
    excel: false,
  });

  const [logs, setLogs] = useState<string[]>([]);
  const [pagesOut, setPagesOut] = useState<PageResult[]>([]);

  const log = useCallback((msg: string) => {
    setLogs((prev) => [...prev, msg]);
    console.log(msg);
  }, []);

  /** ---------------------------------------
   *  PDF.js loader (pdfjs-dist v4 with local worker)
   * -------------------------------------- */
  useEffect(() => {
    let cancelled = false;

    const boot = async () => {
      if (typeof window === "undefined") return;

      try {
        // pdf.js library
        const pdfjsLib: any = await import("pdfjs-dist");
        // use the worker file you placed into /public
        pdfjsLib.GlobalWorkerOptions.workerSrc = "/pdf.worker.min.mjs";

        if (cancelled) return;
        pdfjsRef.current = pdfjsLib;
        setPdfReady(true);
        log("PDF engine ready (local worker).");
      } catch (err) {
        console.error(err);
        log("Failed to load PDF engine. See console for details.");
      }
    };

    boot();
    return () => {
      cancelled = true;
    };
  }, [log]);

  /** ---------------------------------------
   *  Helpers: log & reset
   * -------------------------------------- */
  const clearLog = () => setLogs([]);
  const resetAll = () => {
    setLogs([]);
    setPagesOut([]);
    setWorking({ loading: false, ocr: false, parse: false, excel: false });
  };

  /** ---------------------------------------
   *  File handlers
   * -------------------------------------- */
  const onChooseFile = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    if (!pdfReady) {
      alert("The PDF engine is still loading. Please wait a moment and try again.");
      return;
    }
    await handleFile(file);
  };

  const onDrop = async (ev: React.DragEvent<HTMLDivElement>) => {
    ev.preventDefault();
    if (!pdfReady) {
      alert("The PDF engine is still loading. Please wait a moment and try again.");
      return;
    }
    const file = ev.dataTransfer.files?.[0];
    if (!file) return;
    await handleFile(file);
  };

  const onDragOver = (ev: React.DragEvent<HTMLDivElement>) => {
    ev.preventDefault();
  };

  /** ---------------------------------------
   *  Pipeline: PDF -> text/ocr -> parse -> Excel
   * -------------------------------------- */
  const handleFile = async (file: File) => {
    try {
      // fresh run
      setLogs([]);
      setPagesOut([]);
      log(`Selected file: ${file.name}`);

      setWorking((w) => ({ ...w, loading: true }));

      const data = new Uint8Array(await file.arrayBuffer());
      const pdfjs = pdfjsRef.current;

      const loadingTask = pdfjs.getDocument({ data });
      const doc = await loadingTask.promise;
      log(`PDF loaded. Pages: ${doc.numPages}`);
      setWorking((w) => ({ ...w, loading: false }));

      const results: PageResult[] = [];

      for (let i = 1; i <= doc.numPages; i++) {
        log(`Rendering page ${i}/${doc.numPages}…`);
        setWorking((w) => ({ ...w, ocr: true }));

        const page = await doc.getPage(i);
        const viewport = page.getViewport({ scale: 2.0 });

        // Render to canvas (needed for OCR fallback)
        const canvas = document.createElement("canvas");
        canvas.width = Math.ceil(viewport.width);
        canvas.height = Math.ceil(viewport.height);
        const ctx = canvas.getContext("2d", { willReadFrequently: true })!;
        await page.render({ canvasContext: ctx as any, viewport }).promise;

        // Try PDF-native text extraction first
        const textContent = await page.getTextContent();
        const rawText = textContent.items
          .map((item: any) => ("str" in item ? item.str : ""))
          .join("\n")
          .replace(/\u00A0/g, " ")
          .replace(/[ \t]+/g, " ")
          .trim();

        setWorking((w) => ({ ...w, parse: true }));

        let pageResult: PageResult;
        if (rawText && rawText.length > 50) {
          // parse from text
          pageResult = await parseDisbursementPageFromText(rawText, { pageIndex: i });
        } else {
          // if text layer is empty, OCR the canvas
          pageResult = await parseDisbursementPageFromCanvas(canvas as HTMLCanvasElement, {
            pageIndex: i,
          });
        }

        results.push(pageResult);

        setWorking((w) => ({ ...w, ocr: false, parse: false }));
        log(`Page ${i} processed.`);
      }

      setPagesOut(results);

      // Build & download Excel
      log("Building Excel workbook…");
      setWorking((w) => ({ ...w, excel: true }));

      const workbook = await createWorkbookFromPages(results);
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const outName = file.name.replace(/\.pdf$/i, "") + ".xlsx";
      saveAs(blob, outName);

      log("Done. Excel downloaded.");
    } catch (err: any) {
      console.error(err);
      log("Error: " + (err?.message || String(err)));
    } finally {
      setWorking((w) => ({ loading: false, ocr: false, parse: false, excel: false }));
    }
  };

  const disabled = !pdfReady;

  /** ---------------------------------------
   *  UI
   * -------------------------------------- */
  return (
    <div className={styles.page}>
      <header className={styles.header}>
        <span className={styles.brandDot} />
        <div className={styles.title}>ASC Disbursement Converter</div>
        <div className={styles.subtitle}>PDF → Excel (one worksheet per page)</div>
      </header>

      <main className={styles.container}>
        <section className={styles.card}>
          <div
            className={styles.dropzone}
            data-disabled={!pdfReady}
            onDrop={onDrop}
            onDragOver={onDragOver}
          >
            <div className={styles.dropTitle}>Drag &amp; drop your PDF here</div>
            <div className={styles.dropHint}>or</div>

            <div className={styles.actions}>
              <label className={styles.primaryBtn}>
                {pdfReady ? "Choose PDF…" : "Loading PDF engine…"}
                <input
                  className={styles.hiddenInput}
                    type="file"
                    accept="application/pdf"
                    onChange={onChooseFile}
                    disabled={disabled}
                  />
                </label>
              </div>

            {/* Status chips + new buttons */}
            <div className={styles.progressRow}>
              <span className={`${styles.badge} ${working.loading ? styles.badgeActive : ""}`}>
                Load
              </span>
              <span className={`${styles.badge} ${working.ocr ? styles.badgeActive : ""}`}>OCR</span>
              <span className={`${styles.badge} ${working.parse ? styles.badgeActive : ""}`}>
                Parse
              </span>
              <span className={`${styles.badge} ${working.excel ? styles.badgeActive : ""}`}>
                Excel
              </span>
              <span className={styles.badge}>
                {pagesOut.length > 0 ? `${pagesOut.length} page(s)` : "0 page"}
              </span>

              {/* NEW buttons */}
              <button type="button" className={styles.ghostBtn} onClick={clearLog}>
                Clear log
              </button>
              <button
                type="button"
                className={styles.ghostBtn}
                onClick={resetAll}
                title="Clear log + results"
              >
                Reset
              </button>
            </div>

            <div className={styles.logCard}>
              <div className={styles.logTitle}>Log</div>
              {logs.length === 0 ? (
                <div className={styles.logEmpty}>Upload a PDF to begin.</div>
              ) : (
                <ul className={styles.logList}>
                  {logs.map((l, idx) => (
                    <li key={idx}>{l}</li>
                  ))}
                </ul>
              )}
            </div>
          </div>
        </section>
      </main>

      <footer className={styles.footer}>
        Built for ASC • PDF.js v4 (text → OCR fallback) → Excel
      </footer>
    </div>
  );
}
