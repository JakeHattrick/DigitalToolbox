// electron/wordWorker.js
const path = require('path');
const fs = require('fs');
const os = require('os');

async function mergePdfs(pdfPaths, outPath) {
  const { PDFDocument } = require('pdf-lib');
  const merged = await PDFDocument.create();
  for (const p of pdfPaths) {
    const bytes = fs.readFileSync(p);
    const src = await PDFDocument.load(bytes);
    const pages = await merged.copyPages(src, src.getPageIndices());
    pages.forEach(pg => merged.addPage(pg));
  }
  const outBytes = await merged.save();
  fs.mkdirSync(path.dirname(outPath), { recursive: true });
  fs.writeFileSync(outPath, outBytes);
}

function saveAsPdf(doc, outPath) {
  // Prefer ExportAsFixedFormat; fallback to SaveAs2/SaveAs (17 = PDF)
  if (typeof doc.ExportAsFixedFormat === 'function') {
    doc.ExportAsFixedFormat(outPath, 17);
  } else if (typeof doc.SaveAs2 === 'function') {
    doc.SaveAs2(outPath, 17);
  } else if (typeof doc.SaveAs === 'function') {
    doc.SaveAs(outPath, 17);
  } else {
    throw new Error('No PDF export method available on this Word build');
  }
}

async function run({ inputFiles, outputPath, mergeDocuments }) {
  let winax, word;
  // temp folder for per-file PDFs when merging
  const tempRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'docx2pdf-'));
  const tempPdfs = [];

  try {
    winax = require('winax');
    word = new winax.Object('Word.Application');
    word.Visible = false;
    word.DisplayAlerts = 0; // wdAlertsNone
    word.Options.ConfirmConversions = false;

    const docs = word.Documents;

    if (mergeDocuments) {
      // 1) Convert each DOCX to its own PDF in a temp folder (highest fidelity)
      for (let i = 0; i < inputFiles.length; i++) {
        const file = inputFiles[i];
        const pdfOut = path.join(
          tempRoot,
          path.basename(file, path.extname(file)) + `.${i}.pdf`
        );
        const doc = docs.Open(file, false, true); // ReadOnly
        saveAsPdf(doc, pdfOut);
        doc.Close(false);
        tempPdfs.push(pdfOut);
        if (process.send) process.send({ progress: Math.round(((i + 1) / inputFiles.length) * 80) });
      }

      // 2) Merge PDFs (purely visual concat; no formatting loss)
      await mergePdfs(tempPdfs, outputPath);
      if (process.send) process.send({ progress: 100 });

    } else {
      // Convert separately (unchanged)
      fs.mkdirSync(outputPath, { recursive: true });
      for (let i = 0; i < inputFiles.length; i++) {
        const file = inputFiles[i];
        const out = path.join(
          outputPath,
          path.basename(file, path.extname(file)) + '.pdf'
        );
        const doc = docs.Open(file, false, true);
        saveAsPdf(doc, out);
        doc.Close(false);
        if (process.send) process.send({ progress: Math.round(((i + 1) / inputFiles.length) * 100) });
      }
    }

    return { success: true };
  } catch (err) {
    return { success: false, error: err && err.message ? err.message : String(err) };
  } finally {
    if (word) { try { word.Quit(); } catch {} }
    // Clean up temp PDFs
    try {
      for (const p of tempPdfs) { try { fs.unlinkSync(p); } catch {} }
      try { fs.rmdirSync(tempRoot); } catch {}
    } catch {}
  }
}

// Fork protocol
if (process.send) {
  process.on('message', async (msg) => {
    const result = await run(msg);
    process.send(result);
  });
} else {
  (async () => {
    const [, , json] = process.argv;
    const result = await run(JSON.parse(json));
    console.log(JSON.stringify(result));
  })();
}
