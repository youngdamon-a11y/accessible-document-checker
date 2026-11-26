// Renders one PDF page to a PNG, for local vision models that cannot read PDFs.
// Runs in a hidden window with no access to Node; the main process sends the bytes.
import * as pdfjs from '../../node_modules/pdfjs-dist/legacy/build/pdf.mjs';

pdfjs.GlobalWorkerOptions.workerSrc = new URL('../../node_modules/pdfjs-dist/legacy/build/pdf.worker.mjs', import.meta.url).href;

window.renderPdfPage = async (base64, pageIndex = 0, maxWidth = 1400) => {
  const data = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
  const task = pdfjs.getDocument({ data, isEvalSupported: false });
  const doc = await task.promise;
  try {
    const page = await doc.getPage(Math.min(doc.numPages, pageIndex + 1));
    const base = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({ scale: Math.min(3, maxWidth / base.width) });
    const canvas = document.createElement('canvas');
    canvas.width = Math.ceil(viewport.width); canvas.height = Math.ceil(viewport.height);
    const ctx = canvas.getContext('2d');
    ctx.fillStyle = '#fff'; ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, canvas, viewport }).promise;
    return canvas.toDataURL('image/png').split(',')[1];
  } finally {
    await task.destroy();
  }
};
window.rendererReady = true;
