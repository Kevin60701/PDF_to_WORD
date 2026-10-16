(function () {
  "use strict";

  // ---------- 函式庫載入（多個來源自動備援）----------
  const PDF_SOURCES = [
    { js: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js",
      worker: "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js" },
    { js: "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.min.js",
      worker: "https://cdn.jsdelivr.net/npm/pdfjs-dist@3.11.174/build/pdf.worker.min.js" },
    { js: "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.min.js",
      worker: "https://unpkg.com/pdfjs-dist@3.11.174/build/pdf.worker.min.js" }
  ];
  const DOCX_SOURCES = [
    "https://cdn.jsdelivr.net/npm/docx@8.5.0/build/index.umd.js",
    "https://unpkg.com/docx@8.5.0/build/index.umd.js"
  ];

  function loadScript(url) {
    return new Promise((resolve, reject) => {
      const s = document.createElement("script");
      s.src = url;
      s.onload = resolve;
      s.onerror = () => reject(new Error("無法載入 " + url));
      document.head.appendChild(s);
    });
  }

  let libsPromise = null;
  // 回傳載入失敗的函式庫名稱陣列（空陣列代表全部成功）
  function ensureLibs() {
    if (libsPromise) return libsPromise;
    libsPromise = (async () => {
      const failed = [];
      if (!window.pdfjsLib) {
        let ok = false;
        for (const src of PDF_SOURCES) {
          try {
            await loadScript(src.js);
            if (window.pdfjsLib) {
              pdfjsLib.GlobalWorkerOptions.workerSrc = src.worker;
              ok = true;
              break;
            }
          } catch (e) { /* 換下一個來源 */ }
        }
        if (!ok) failed.push("pdf.js（讀取 PDF）");
      }
      if (!window.docx) {
        let ok = false;
        for (const url of DOCX_SOURCES) {
          try {
            await loadScript(url);
            if (window.docx) { ok = true; break; }
          } catch (e) { /* 換下一個來源 */ }
        }
        if (!ok) failed.push("docx（產生 Word）");
      }
      return failed;
    })();
    return libsPromise;
  }
  ensureLibs(); // 開啟頁面時先預載

  const $ = (s) => document.querySelector(s);
  const drop = $("#drop"), fileInput = $("#fileInput"), fileInfo = $("#fileInfo");
  const convertBtn = $("#convertBtn"), progressBox = $("#progressBox");
  const barFill = $("#barFill"), statusEl = $("#status"), msgEl = $("#msg");
  const resultCard = $("#result"), downloadLink = $("#downloadLink"), previewEl = $("#preview");
  const textOpts = $("#textOpts");

  let currentFile = null;
  let downloadUrl = null;

  // ---------- 小工具 ----------
  const CJK = /[\u3000-\u303f\u3400-\u9fff\uff00-\uffef]/;
  const wait = () => new Promise((r) => setTimeout(r, 0));
  const fmtSize = (n) => n > 1048576 ? (n / 1048576).toFixed(2) + " MB" : (n / 1024).toFixed(1) + " KB";
  const getMode = () => document.querySelector('input[name="mode"]:checked').value;

  function showMsg(text, type) {
    msgEl.textContent = text;
    msgEl.className = type || "";
    msgEl.style.display = text ? "block" : "none";
  }
  function setProgress(p, text) {
    barFill.style.width = Math.round(p * 100) + "%";
    statusEl.textContent = text || "";
  }
  function median(arr) {
    if (!arr.length) return 12;
    const s = arr.slice().sort((a, b) => a - b);
    return s[Math.floor(s.length / 2)];
  }
  function joinText(a, b) {
    if (!a) return b;
    if (!b) return a;
    const last = a.slice(-1), first = b[0];
    if (CJK.test(last) || CJK.test(first)) return a + b;
    if (/\s$/.test(a) || /^\s/.test(b)) return a + b;
    if (last === "-") return a.slice(0, -1) + b;
    return a + " " + b;
  }

  // ---------- 檔案選擇 ----------
  function setFile(f) {
    showMsg("");
    resultCard.style.display = "none";
    if (!f) return;
    const isPdf = f.type === "application/pdf" || /\.pdf$/i.test(f.name);
    if (!isPdf) {
      currentFile = null;
      convertBtn.disabled = true;
      fileInfo.style.display = "none";
      showMsg("請選擇 PDF 檔案（.pdf）。", "error");
      return;
    }
    currentFile = f;
    fileInfo.innerHTML = "已選擇：<b></b>（" + fmtSize(f.size) + "）";
    fileInfo.querySelector("b").textContent = f.name;
    fileInfo.style.display = "block";
    convertBtn.disabled = false;
  }

  drop.addEventListener("click", () => fileInput.click());
  drop.addEventListener("keydown", (e) => { if (e.key === "Enter" || e.key === " ") fileInput.click(); });
  fileInput.addEventListener("change", () => setFile(fileInput.files[0]));
  ["dragenter", "dragover"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach((ev) =>
    drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", (e) => setFile(e.dataTransfer.files[0]));

  function syncModeOptions() {
    textOpts.style.display = getMode() === "text" ? "block" : "none";
  }
  document.querySelectorAll('input[name="mode"]').forEach((r) =>
    r.addEventListener("change", syncModeOptions));
  syncModeOptions();

  // ---------- 文字擷取 ----------
  async function extractPage(page, keepBreaks) {
    const tc = await page.getTextContent();
    const items = tc.items
      .filter((i) => i.str && i.str.length && i.transform)
      .map((i) => ({
        s: i.str, x: i.transform[4], y: i.transform[5], w: i.width,
        size: Math.hypot(i.transform[0], i.transform[1]) || i.height || 12,
        fontName: i.fontName
      }));
    if (!items.length) return [];

    // 1) 依 y 由上到下排序，歸成「行」
    items.sort((a, b) => b.y - a.y || a.x - b.x);
    const lines = [];
    for (const it of items) {
      const last = lines[lines.length - 1];
      if (last && Math.abs(it.y - last.y) <= Math.max(it.size, last.size) * 0.4) {
        last.items.push(it);
        last.size = Math.max(last.size, it.size);
      } else {
        lines.push({ y: it.y, size: it.size, items: [it] });
      }
    }

    // 2) 每行依 x 排序並組成文字
    for (const ln of lines) {
      ln.items.sort((a, b) => a.x - b.x);
      let text = "", prev = null;
      for (const it of ln.items) {
        if (prev) {
          const gap = it.x - (prev.x + prev.w);
          if (gap > it.size * 0.3 && !/\s$/.test(text) && !/^\s/.test(it.s)) text += " ";
        }
        text += it.s;
        prev = it;
      }
      ln.text = text.replace(/\s+$/, "").replace(/^\s+/, "");
      ln.x = Math.min(...ln.items.map((it) => it.x));
      ln.right = Math.max(...ln.items.map((it) => it.x + it.w));
    }

    const pageViewport = page.getViewport({ scale: 1 });
    const pageHeight = pageViewport.height;
    const pageWidth = pageViewport.width;

    // 3) 行 → 原生段落（保留來源行的垂直位置）
    const sizes = [];
    lines.forEach((l) => { for (let k = 0; k < Math.min(l.text.length, 40); k++) sizes.push(l.size); });
    const med = median(sizes);

    const paras = [];
    let cur = null, prevLine = null;
    for (const ln of lines) {
      if (!ln.text) continue;
      const heading = ln.size > med * 1.2 && ln.text.length < 120;
      const gap = prevLine ? prevLine.y - ln.y : 0;
      const startsQuestion = /^(?:Q\s*\d+(?:\s*[~～至-]\s*Q?\s*\d+)?\s*[:：]|[（(]?\d+[)）.、])/i.test(ln.text.trim());
      const startNew =
        !cur ||
        keepBreaks ||
        startsQuestion ||
        heading !== cur.heading ||
        (heading && Math.abs(ln.size - cur.size) > 0.5) ||
        gap > Math.max(ln.size, prevLine.size) * 1.7;
      if (startNew) {
        let level = 0;
        if (heading) {
          const ratio = ln.size / med;
          level = ratio >= 1.8 ? 1 : ratio >= 1.5 ? 2 : 3;
        }
        const beforePt = prevLine
          ? Math.max(0, gap - prevLine.size * 1.25)
          : Math.max(0, pageHeight - ln.y - ln.size * 0.25);
        const centerX = (ln.x + ln.right) / 2;
        const topFont = ln.items.find((it) => it.s.trim())?.fontName;
        cur = {
          kind: "text",
          text: ln.text, heading, level, size: ln.size, ratio: ln.size / med,
          fontFamily: (topFont && tc.styles[topFont] && tc.styles[topFont].fontFamily) || "Times New Roman",
          preserveLine: keepBreaks,
          x: ln.x,
          topPt: Math.max(0, pageHeight - ln.y - ln.size * 0.85),
          firstBaselineY: ln.y,
          firstX: ln.x,
          leftX: ln.x,
          heightPt: ln.size * 1.25,
          beforePt,
          centered: ln.text.trim().length <= 40 && Math.abs(centerX - pageWidth / 2) <= Math.max(5, ln.size * 0.35)
        };
        paras.push(cur);
      } else {
        cur.text = joinText(cur.text, ln.text);
        // Same native Word paragraph spans these source PDF lines; count its full height
        // when positioning following images/paragraphs to prevent cumulative drift.
        cur.leftX = Math.min(cur.leftX, ln.x);
        cur.heightPt = Math.max(cur.heightPt, cur.firstBaselineY - ln.y + ln.size * 1.25);
      }
      prevLine = ln;
    }
    return paras;
  }

  async function extractPageImages(page) {
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    const opList = await page.getOperatorList();
    const OPS = pdfjsLib.OPS;
    const stack = [];
    let ctm = [1, 0, 0, 1, 0, 0];
    const blocks = [];
    const seen = new Set();
    for (let i = 0; i < opList.fnArray.length; i++) {
      const fn = opList.fnArray[i], args = opList.argsArray[i] || [];
      if (fn === OPS.save) stack.push(ctm.slice());
      else if (fn === OPS.restore) ctm = stack.pop() || [1, 0, 0, 1, 0, 0];
      else if (fn === OPS.transform && args.length >= 6) ctm = pdfjsLib.Util.transform(ctm, args);
      else if (fn === OPS.paintImageXObject || fn === OPS.paintJpegXObject) {
        const m = pdfjsLib.Util.transform(viewport.transform, ctm);
        const pts = [[m[4], m[5]], [m[0] + m[4], m[1] + m[5]], [m[2] + m[4], m[3] + m[5]], [m[0] + m[2] + m[4], m[1] + m[3] + m[5]]];
        const xs = pts.map((pt) => pt[0]), ys = pts.map((pt) => pt[1]);
        const sx = Math.max(0, Math.floor(Math.min(...xs))), sy = Math.max(0, Math.floor(Math.min(...ys)));
        const ex = Math.min(canvas.width, Math.ceil(Math.max(...xs))), ey = Math.min(canvas.height, Math.ceil(Math.max(...ys)));
        if (ex - sx < 8 || ey - sy < 8) continue;
        const key = [sx, sy, ex, ey].join(":");
        if (seen.has(key)) continue;
        seen.add(key);
        const crop = document.createElement("canvas");
        crop.width = ex - sx; crop.height = ey - sy;
        crop.getContext("2d").drawImage(canvas, sx, sy, crop.width, crop.height, 0, 0, crop.width, crop.height);
        const blob = await new Promise((resolve) => crop.toBlob(resolve, "image/png"));
        if (!blob) continue;
        blocks.push({
          kind: "image", data: new Uint8Array(await blob.arrayBuffer()),
          x: sx / 2, topPt: sy / 2, widthPt: crop.width / 2, heightPt: crop.height / 2
        });
        crop.width = crop.height = 0;
      }
    }
    canvas.width = canvas.height = 0;
    return blocks;
  }

  // ---------- 建立 Word ----------
  const FONT = { ascii: "Times New Roman", hAnsi: "Times New Roman", eastAsia: "新細明體", cs: "Times New Roman" };

  function buildTextParagraphs(pages, pageBreaks, pageWidthPt) {
    const { Paragraph, TextRun, ImageRun } = docx;
    const pageInsetPt = 72;
    const contentWidthPt = pageWidthPt - pageInsetPt * 2;
    const children = [];
    let currentListInstance = 0;
    let lastListNumber = 0;
    pages.forEach((paras, pi) => {
      let first = true;
      if (!paras.length && pageBreaks && pi > 0) {
        children.push(new Paragraph({ children: [], pageBreakBefore: true }));
        return;
      }
      for (const p of paras) {
        const listMatch = p.kind === "text" && p.text.match(/^\s*[（(](\d+)[）)]\s*/);
        let listInstance = null;
        if (listMatch) {
          const listNumber = Number(listMatch[1]);
          if (listNumber === 1 || lastListNumber === 0 || listNumber !== lastListNumber + 1) currentListInstance++;
          lastListNumber = listNumber;
          listInstance = currentListInstance;
          p.text = p.text.replace(/^\s*[（(]\d+[）)]\s*/, "");
        } else if (p.kind === "text" && p.text.trim()) {
          lastListNumber = 0;
        }
        const opts = {
          alignment: p.centered ? docx.AlignmentType.CENTER : docx.AlignmentType.LEFT,
          indentation: listMatch ? undefined : (p.centered ? {} : {
            left: Math.round(Math.max(0, (p.leftX ?? p.x) - pageInsetPt) * 20),
            firstLine: Math.round(((p.firstX ?? p.x) - (p.leftX ?? p.x)) * 20)
          }),
          spacing: {
            before: Math.round((p.flowBeforePt ?? p.beforePt ?? 0) * 20),
            after: 0,
            line: p.preserveLine || p.kind === "image" ? 240 : 300
          },
          numbering: listMatch ? { reference: "pdf-question-items", level: 0, instance: listInstance } : undefined,
          children: p.kind === "image"
            ? [new ImageRun({
                data: p.data, type: "png",
                transformation: (() => {
                  const xFromMargin = Math.max(0, (p.x || 0) - pageInsetPt);
                  const scale = Math.min(1, Math.max(12, contentWidthPt - xFromMargin) / p.widthPt);
                  return { width: Math.round(p.widthPt * scale * 96 / 72), height: Math.round(p.heightPt * scale * 96 / 72) };
                })()
              })]
            : [new TextRun({
                text: p.text,
                font: { ...FONT, ascii: p.fontFamily || FONT.ascii, hAnsi: p.fontFamily || FONT.hAnsi },
                bold: p.heading,
                size: Math.max(16, Math.round(p.size * 2))
              })]
        };
        if (first && pi > 0 && pageBreaks) opts.pageBreakBefore = true;
        children.push(new Paragraph(opts));
        first = false;
      }
    });
    return children;
  }

  async function renderPageImage(page) {
    const viewport = page.getViewport({ scale: 2 });
    const canvas = document.createElement("canvas");
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const ctx = canvas.getContext("2d");
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, viewport }).promise;
    const blob = await new Promise((r) => canvas.toBlob(r, "image/jpeg", 0.85));
    const data = new Uint8Array(await blob.arrayBuffer());
    const ratio = canvas.width / canvas.height;
    canvas.width = canvas.height = 0;
    return { data, ratio };
  }

  // ---------- 精確版面原生段落模式 ----------
  // 非文字視覺內容以整頁背景圖保留並擦除原文字；文字則用一般 Word 段落重建，
  // 以段落縮排與段前距貼近原座標，同時保有 Word 自動換行與段落編輯能力。
  const EXACT_SCALE = 3; // 背景圖解析度倍率（3 = 216 dpi）
  const CJK_RE = /[　-〿㐀-鿿豈-﫿＀-￯]/;
  const CJK_FAMILIES = ["標楷體", "新細明體", "微軟正黑體", "細明體"];
  const twip = (pt) => Math.round(pt * 20);
  const hex2 = (n) => (n < 16 ? "0" : "") + n.toString(16);

  function mapFontName(raw) {
    const n = String(raw || "").replace(/^[A-Z]{6}\+/, "");
    const bold = /bold|black|heavy/i.test(n);
    const italic = /italic|oblique/i.test(n);
    let family;
    if (/TimesNewRoman|^Times/i.test(n)) family = "Times New Roman";
    else if (/^Arial|^Helvetica/i.test(n)) family = "Arial";
    else if (/CambriaMath/i.test(n)) family = "Cambria Math";
    else if (/^Cambria/i.test(n)) family = "Cambria";
    else if (/^Calibri/i.test(n)) family = "Calibri";
    else if (/DFKai|KaiU|Kai-?SB|標楷/i.test(n)) family = "標楷體";
    else if (/PMingLiU|MingLiU|細明/i.test(n)) family = "新細明體";
    else if (/JhengHei|正黑/i.test(n)) family = "微軟正黑體";
    else if (/^Courier|monospace/i.test(n)) family = "Courier New";
    else if (/^Consolas/i.test(n)) family = "Consolas";
    else if (/^Symbol/i.test(n)) family = "Symbol";
    else if (/^serif$/i.test(n)) family = "Times New Roman";
    else if (/^sans-serif$/i.test(n)) family = "Arial";
    else {
      family = n.split(/[-,]/)[0].replace(/(PSMT|MT|PS)$/, "").replace(/([a-z])([A-Z])/g, "$1 $2").trim() || "Arial";
    }
    return { family, bold, italic };
  }

  function fontInfoFor(page, fontId, styles, cache) {
    if (cache[fontId]) return cache[fontId];
    let name = "", fb = false, fi = false;
    try {
      if (page.commonObjs.has(fontId)) {
        const f = page.commonObjs.get(fontId);
        name = f.name || f.fallbackName || "";
        fb = !!f.bold; fi = !!f.italic;
      }
    } catch (e) { /* 取不到就用備援 */ }
    if (!name && styles && styles[fontId]) name = styles[fontId].fontFamily || "";
    const info = mapFontName(name);
    info.bold = info.bold || fb;
    info.italic = info.italic || fi;
    cache[fontId] = info;
    return info;
  }

  // 分析文字所在區域：背景色、文字色、是否有底線
  function analyzeRegion(ctx, cw, ch, x0, y0, x1, y1) {
    x0 = Math.max(0, Math.floor(x0)); y0 = Math.max(0, Math.floor(y0));
    x1 = Math.min(cw, Math.ceil(x1)); y1 = Math.min(ch, Math.ceil(y1));
    const w = x1 - x0, h = y1 - y0;
    if (w < 1 || h < 1) return null;
    const d = ctx.getImageData(x0, y0, w, h).data;
    const counts = new Map();
    for (let i = 0; i < d.length; i += 4) {
      const key = ((d[i] >> 4) << 8) | ((d[i + 1] >> 4) << 4) | (d[i + 2] >> 4);
      counts.set(key, (counts.get(key) || 0) + 1);
    }
    let bestKey = 0, bestCount = -1;
    counts.forEach((c, k) => { if (c > bestCount) { bestCount = c; bestKey = k; } });
    let sr = 0, sg = 0, sb = 0, sn = 0;
    for (let i = 0; i < d.length; i += 4) {
      const key = ((d[i] >> 4) << 8) | ((d[i + 1] >> 4) << 4) | (d[i + 2] >> 4);
      if (key === bestKey) { sr += d[i]; sg += d[i + 1]; sb += d[i + 2]; sn++; }
    }
    const bg = [Math.round(sr / sn), Math.round(sg / sn), Math.round(sb / sn)];
    let maxDist = 0, tc = null;
    const rowInk = new Float32Array(h);
    for (let y = 0; y < h; y++) {
      let ink = 0;
      for (let x = 0; x < w; x++) {
        const i = (y * w + x) * 4;
        const dr = d[i] - bg[0], dg = d[i + 1] - bg[1], db = d[i + 2] - bg[2];
        const dist = dr * dr + dg * dg + db * db;
        if (dist > maxDist) { maxDist = dist; tc = [d[i], d[i + 1], d[i + 2]]; }
        if (dist > 8100) ink++;
      }
      rowInk[y] = ink / w;
    }
    return {
      bg: bg,
      text: maxDist > 3600 ? tc : null,
      rowInk: rowInk,
      y0: y0
    };
  }

  async function buildExactSection(page, ctxInfo) {
    const D = docx;
    const E = (o, k, fb) => (o && o[k] !== undefined ? o[k] : fb);
    const LINE_EXACT = E(D.LineRuleType, "EXACT", "exact");
    const HEIGHT_EXACT = E(D.HeightRule, "EXACT", "exact");
    const H_PAGE = E(D.HorizontalPositionRelativeFrom, "PAGE", "page");
    const V_PAGE = E(D.VerticalPositionRelativeFrom, "PAGE", "page");
    const TW_NONE = E(D.TextWrappingType, "NONE", 0);
    const ORI_LAND = E(D.PageOrientation, "LANDSCAPE", "landscape");
    const ORI_PORT = E(D.PageOrientation, "PORTRAIT", "portrait");
    const vp1 = page.getViewport({ scale: 1 });
    const vp = page.getViewport({ scale: EXACT_SCALE });
    const W = vp1.width, H = vp1.height;

    // 先載入字型資訊，再取文字
    try { await page.getOperatorList(); } catch (e) { /* 忽略 */ }
    const tc = await page.getTextContent();
    const styles = tc.styles || {};

    const canvas = document.createElement("canvas");
    canvas.width = Math.round(vp.width);
    canvas.height = Math.round(vp.height);
    const ctx = canvas.getContext("2d", { willReadFrequently: true });
    ctx.fillStyle = "#fff";
    ctx.fillRect(0, 0, canvas.width, canvas.height);
    await page.render({ canvasContext: ctx, canvas: canvas, viewport: vp }).promise;
    const S = canvas.width / W;

    // 1) 整理文字項目（只處理水平文字；旋轉文字保留在背景圖中）
    const items = [];
    for (const it of tc.items) {
      if (!it.str || !it.transform || !it.width) continue;
      const t = pdfjsLib.Util.transform(vp1.transform, it.transform);
      const size = Math.hypot(t[2], t[3]);
      if (!(size > 0.5)) continue;
      if (Math.abs(t[1]) > size * 0.2 || Math.abs(t[2]) > size * 0.2) continue;
      const fi = fontInfoFor(page, it.fontName, styles, ctxInfo.fontCache);
      if (CJK_FAMILIES.indexOf(fi.family) >= 0) ctxInfo.cjkFamily = fi.family;
      const x = t[4], base = t[5], w = it.width * (vp1.scale || 1);
      const blank = !it.str.trim();
      if (blank && w > size * 0.8) continue; // 很寬的空白＝版面上的空隙，不當成文字
      const reg = blank ? null : analyzeRegion(ctx, canvas.width, canvas.height,
        (x - 0.5) * S, (base - size * 0.92) * S, (x + w + 0.5) * S, (base + size * 0.34) * S);
      let underline = false;
      if (reg) {
        const from = Math.max(0, Math.floor((base + size * 0.04) * S - reg.y0));
        const to = Math.min(reg.rowInk.length - 1, Math.ceil((base + size * 0.32) * S - reg.y0));
        for (let r = from; r <= to; r++) if (reg.rowInk[r] > 0.9) { underline = true; break; }
      }
      items.push({
        s: it.str, x: x, y: base, w: w, size: size, fi: fi,
        color: reg && reg.text ? hex2(reg.text[0]) + hex2(reg.text[1]) + hex2(reg.text[2]) : "000000",
        bg: reg ? reg.bg : null, underline: underline, blank: blank,
        eraseBottom: underline ? 0.34 : 0.25
      });
    }

    // 2) 把文字從背景圖擦掉（用該處的背景色填滿）
    for (const it of items) {
      if (it.blank || !it.bg) continue;
      ctx.fillStyle = "rgb(" + it.bg[0] + "," + it.bg[1] + "," + it.bg[2] + ")";
      ctx.fillRect((it.x - 0.5) * S, (it.y - it.size * 0.92) * S,
        (it.w + 1) * S, it.size * (0.92 + it.eraseBottom) * S);
    }
    const blob = await new Promise((r) => canvas.toBlob(r, "image/png"));
    const bgData = new Uint8Array(await blob.arrayBuffer());
    canvas.width = canvas.height = 0;

    // 3) 把 PDF 的續行合併成一般 Word 段落。每段可在 Word 內正常編輯和自動換行。
    const nativeParas = await extractPage(page, false);
    const paragraphs = nativeParas.map((p) => {
      const family = p.fontFamily || "Times New Roman";
      const east = CJK_RE.test(p.text) && CJK_FAMILIES.indexOf(family) < 0 ? ctxInfo.cjkFamily : family;
      return new D.Paragraph({
        alignment: p.centered ? D.AlignmentType.CENTER : D.AlignmentType.LEFT,
        indentation: p.centered ? undefined : {
          left: twip(Math.max(0, p.leftX || p.x || 0)),
          firstLine: twip((p.firstX || p.x || 0) - (p.leftX || p.x || 0))
        },
        spacing: {
          before: twip(p.beforePt || 0), after: 0,
          line: twip(Math.max(8, p.size * 1.3)), lineRule: LINE_EXACT
        },
        children: [new D.TextRun({
          text: p.text,
          font: { ascii: family, hAnsi: family, eastAsia: east, cs: family },
          size: Math.round(p.size * 2),
          bold: p.heading
        })]
      });
    });

    // 5) 背景圖：放在最後一個一般段落內，並設為「文字後方」的浮動圖片
    const anchor = new D.Paragraph({
      spacing: { before: 0, after: 0, line: 20, lineRule: LINE_EXACT },
      children: [new D.ImageRun({
        type: "png",
        data: bgData,
        transformation: { width: Math.round(W * 96 / 72), height: Math.round(H * 96 / 72) },
        floating: {
          horizontalPosition: { relative: H_PAGE, offset: 0 },
          verticalPosition: { relative: V_PAGE, offset: 0 },
          behindDocument: true,
          allowOverlap: true,
          wrap: { type: TW_NONE }
        }
      })]
    });

    const landscape = W > H;
    return {
      paragraphCount: paragraphs.length,
      section: {
        properties: {
          page: {
            size: {
              width: twip(Math.min(W, H)), height: twip(Math.max(W, H)),
              orientation: landscape ? ORI_LAND : ORI_PORT
            },
            margin: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0, gutter: 0 }
          }
        },
        children: [anchor].concat(paragraphs)
      }
    };
  }

  // ---------- 主流程 ----------
  async function convert() {
    if (!currentFile) return;
    convertBtn.disabled = true;
    showMsg("");
    progressBox.style.display = "block";
    setProgress(0, "載入必要元件…");
    const failedLibs = await ensureLibs();
    if (failedLibs.length) {
      libsPromise = null; // 下次按鈕會重新嘗試載入
      progressBox.style.display = "none";
      convertBtn.disabled = false;
      showMsg("無法載入：" + failedLibs.join("、") +
        "。請確認網路連線正常、沒有被防火牆或廣告攔截擴充功能擋住（cdnjs.cloudflare.com、cdn.jsdelivr.net、unpkg.com），" +
        "並建議直接用 Chrome／Edge 開啟檔案或 GitHub Pages 網址，再按一次「開始轉換」。", "error");
      return;
    }
    convertBtn.disabled = true;
    showMsg("");
    resultCard.style.display = "none";
    progressBox.style.display = "block";
    setProgress(0, "讀取檔案中…");

    const mode = getMode();
    const keepBreaks = $("#keepBreaks").checked;
    const pageBreaks = $("#pageBreaks").checked;

    try {
      const buf = await currentFile.arrayBuffer();
      const pdf = await pdfjsLib.getDocument({ data: buf, fontExtraProperties: true }).promise;
      const total = pdf.numPages;
      const { Document, Packer, Paragraph, ImageRun, AlignmentType } = docx;
      let doc, previewHtml = "";

      if (mode === "exact") {
        const sections = [];
        const ctxInfo = { fontCache: {}, cjkFamily: "新細明體" };
        let paragraphTotal = 0;
        for (let i = 1; i <= total; i++) {
          setProgress((i - 1) / total, "還原版面：第 " + i + " / " + total + " 頁");
          const page = await pdf.getPage(i);
          const res = await buildExactSection(page, ctxInfo);
          sections.push(res.section);
          paragraphTotal += res.paragraphCount;
          await wait();
        }
        setProgress(0.95, "建立 Word 檔案…");
        doc = new Document({ creator: "PDF 轉 Word 小工具", sections: sections });
        previewHtml = "<p>版面精確原生段落模式：共處理 " + total + " 頁、建立 " + paragraphTotal +
          " 個可自動換行的 Word 段落。頁面圖形與表格線保留在背景；文字可像一般段落編輯。</p>";
      } else if (mode === "text") {
        const pages = [];
        let chars = 0;
        let pageSize = null;
        for (let i = 1; i <= total; i++) {
          setProgress((i - 1) / total, "擷取文字：第 " + i + " / " + total + " 頁");
          const page = await pdf.getPage(i);
          if (i === 1) {
            const viewport = page.getViewport({ scale: 1 });
            pageSize = { width: Math.round(viewport.width * 20), height: Math.round(viewport.height * 20) };
          }
          const paras = await extractPage(page, keepBreaks);
          const images = await extractPageImages(page);
          const blocks = paras.concat(images).sort((a, b) => a.topPt - b.topPt || (a.x || 0) - (b.x || 0));
          let flowBottom = 0;
          for (const block of blocks) {
            block.flowBeforePt = Math.max(0, block.topPt - flowBottom);
            flowBottom = Math.max(flowBottom, block.topPt + block.heightPt);
          }
          paras.forEach((p) => (chars += p.text.length));
          pages.push(blocks);
          await wait();
        }
        if (chars < 20) {
          showMsg("幾乎找不到可擷取的文字，這份 PDF 可能是掃描檔。請改用「圖片模式」。", "warn");
          progressBox.style.display = "none";
          convertBtn.disabled = false;
          return;
        }
        setProgress(0.95, "建立 Word 檔案…");
        doc = new Document({
          creator: "PDF 轉 Word 小工具",
          numbering: {
            config: [{
              reference: "pdf-question-items",
              levels: [{
                level: 0,
                format: docx.LevelFormat.DECIMAL,
                text: "(%1)",
                alignment: docx.AlignmentType.START,
                style: { paragraph: { indent: { left: 880, hanging: 360 } } }
              }]
            }]
          },
          sections: [{
            properties: { page: { size: pageSize, margin: { top: 0, right: 1440, bottom: 0, left: 1440 } } },
            children: buildTextParagraphs(pages, pageBreaks, pageSize.width / 20)
          }]
        });
        // 預覽
        let count = 0;
        outer: for (const paras of pages) {
          for (const p of paras) {
            if (p.kind === "image") continue;
            const el = p.heading ? "h4" : "p";
            const safe = p.text.replace(/[&<>]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;" }[c]));
            previewHtml += "<" + el + ">" + safe + "</" + el + ">";
            if (++count >= 12) break outer;
          }
        }
      } else {
        const children = [];
        for (let i = 1; i <= total; i++) {
          setProgress((i - 1) / total, "轉成圖片：第 " + i + " / " + total + " 頁");
          const page = await pdf.getPage(i);
          const { data, ratio } = await renderPageImage(page);
          let w = 700, h = w / ratio;
          if (h > 1000) { h = 1000; w = h * ratio; }
          children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            pageBreakBefore: i > 1,
            children: [new ImageRun({
              data: data, type: "jpg",
              transformation: { width: Math.round(w), height: Math.round(h) }
            })]
          }));
          await wait();
        }
        setProgress(0.95, "建立 Word 檔案…");
        doc = new Document({
          creator: "PDF 轉 Word 小工具",
          sections: [{
            properties: { page: { margin: { top: 567, bottom: 567, left: 567, right: 567 } } },
            children: children
          }]
        });
        previewHtml = "<p>圖片模式：共 " + total + " 頁，每頁已轉成圖片放入 Word。</p>";
      }

      const blob = await Packer.toBlob(doc);
      if (downloadUrl) URL.revokeObjectURL(downloadUrl);
      downloadUrl = URL.createObjectURL(blob);
      const outName = currentFile.name.replace(/\.pdf$/i, "") + ".docx";
      downloadLink.href = downloadUrl;
      downloadLink.download = outName;
      downloadLink.textContent = "⬇ 下載 " + outName + "（" + fmtSize(blob.size) + "）";
      previewEl.innerHTML = previewHtml;
      $("#previewTitle").style.display = "block";
      setProgress(1, "完成，共 " + total + " 頁");
      resultCard.style.display = "block";
      resultCard.scrollIntoView({ behavior: "smooth", block: "center" });
    } catch (err) {
      console.error(err);
      progressBox.style.display = "none";
      if (err && err.name === "PasswordException") {
        showMsg("這份 PDF 有密碼保護，請先解除密碼再轉換。", "error");
      } else if (err && err.name === "InvalidPDFException") {
        showMsg("檔案不是有效的 PDF，或檔案已損毀。", "error");
      } else {
        showMsg("轉換失敗：" + (err && err.message ? err.message : "未知錯誤"), "error");
      }
    } finally {
      convertBtn.disabled = false;
    }
  }

  convertBtn.addEventListener("click", convert);
})();

