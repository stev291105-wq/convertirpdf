/* =====================================================================
   ConvertirPDF · office.js — Conversión FIEL de Word, Excel y PowerPoint
   ---------------------------------------------------------------------
   Se descarga solo cuando alguien convierte un documento de Office.
   Idea general (todo ocurre en el navegador, nada se sube):
     1. El documento se dibuja como páginas HTML con sus medidas reales
        (fuentes, interlineado, márgenes, tablas, colores e imágenes).
     2. Cada página se convierte en una imagen nítida (≈200 ppp).
     3. El PDF lleva esa imagen + una capa de texto invisible, así el
        texto se puede seleccionar, copiar y buscar.
   Motores: docx-preview 0.4.1 (Apache-2.0) · html2canvas-pro 2.4.5 (MIT)
   Fuentes de respaldo con las mismas medidas que las de Office (OFL).
   ===================================================================== */
(function () {
  "use strict";

  window.CPDFOffice = function (K) {
    var VENDOR = K.VENDOR;
    var PX2PT = 0.75;                 // 1 px CSS = 0,75 pt
    var SCALE = 2;                    // resolución de las páginas (2 × 96 ppp ≈ 192 ppp)
    var O = {};

    /* ---------------------------------------------------------------
       1. Motores y utilidades
       --------------------------------------------------------------- */
    function h2c() {
      return K.once("html2canvas", function () {
        return K.loadScript(VENDOR + "html2canvas-pro.min.js").then(function () {
          if (!window.html2canvas) throw new Error("load html2canvas");
          return window.html2canvas;
        });
      });
    }
    function docxLib() {
      return K.Engines.jszip().then(function () {
        return K.once("docxpreview", function () {
          return K.loadScript(VENDOR + "docx-preview.min.js").then(function () {
            if (!window.docx || !window.docx.renderAsync) throw new Error("load docx-preview");
            return window.docx;
          });
        });
      });
    }
    function wait(ms) { return new Promise(function (ok) { setTimeout(ok, ms); }); }
    function nextFrame() { return new Promise(function (ok) { (window.requestAnimationFrame || setTimeout)(function () { ok(); }); setTimeout(ok, 60); }); }

    /* Contenedor fuera de la pantalla donde se dibujan las páginas */
    function makeHost() {
      var host = document.createElement("div");
      host.setAttribute("aria-hidden", "true");
      host.className = "cpdf-render";
      host.style.cssText = "position:absolute;left:-30000px;top:0;width:auto;pointer-events:none;contain:layout style;";
      document.body.appendChild(host);
      return host;
    }

    /* Espera a que carguen imágenes y fuentes antes de medir */
    function settle(root) {
      void root.offsetHeight;   // fuerza el cálculo de estilos para que se pidan las fuentes
      var imgs = Array.prototype.slice.call(root.querySelectorAll("img"));
      return Promise.all(imgs.map(function (im) {
        if (im.complete && im.naturalWidth) return null;
        return new Promise(function (ok) { im.onload = im.onerror = ok; setTimeout(ok, 4000); });
      })).then(function () {
        return document.fonts && document.fonts.ready ? Promise.race([document.fonts.ready, wait(5000)]) : null;
      }).then(nextFrame).then(function () {
        return document.fonts && document.fonts.ready ? Promise.race([document.fonts.ready, wait(3000)]) : null;
      });
    }

    /* ---------------------------------------------------------------
       2. Fuentes: si el equipo no tiene la fuente de Office, se usa una
          equivalente con las MISMAS medidas (así no cambia la paginación)
       --------------------------------------------------------------- */
    var FALLBACK = {
      calibre: "CPDF Calibre", cambre: "CPDF Cambre", sans: "CPDF Sans", serif: "CPDF Serif", mono: "CPDF Mono"
    };
    var FONT_FILES = { "CPDF Calibre": "CPDFCalibre", "CPDF Cambre": "CPDFCambre", "CPDF Sans": "CPDFSans", "CPDF Serif": "CPDFSerif", "CPDF Mono": "CPDFMono" };
    function ensureFontFaces() {
      if (document.getElementById("cpdf-fonts")) return;
      var css = "";
      Object.keys(FONT_FILES).forEach(function (fam) {
        [["Regular", 400, "normal"], ["Bold", 700, "normal"], ["Italic", 400, "italic"], ["BoldItalic", 700, "italic"]].forEach(function (v) {
          css += '@font-face{font-family:"' + fam + '";src:url("' + VENDOR + "fonts/" + FONT_FILES[fam] + "-" + v[0] + '.woff2") format("woff2");font-weight:' + v[1] + ";font-style:" + v[2] + ";font-display:block}\n";
        });
      });
      // Aísla las páginas de los estilos del sitio (modo oscuro, tipografía, etc.)
      css += ".cpdf-render{color:#000;font:16px/normal \"Times New Roman\",\"CPDF Serif\",serif;letter-spacing:normal;text-align:left;-webkit-font-smoothing:auto;}\n" +
        ".cpdf-render *,.cpdf-render *::before,.cpdf-render *::after{box-sizing:content-box;text-wrap:wrap;}\n" +
        ".cpdf-render img,.cpdf-render svg{display:inline;max-width:none;}\n";
      var st = document.createElement("style");
      st.id = "cpdf-fonts"; st.textContent = css;
      document.head.appendChild(st);
    }
    var GENERIC = /^(serif|sans-serif|monospace|cursive|fantasy|system-ui|inherit|initial|unset)$/i;
    function fallbackFor(name) {
      var n = String(name || "").toLowerCase().replace(/["']/g, "").trim();
      if (!n) return FALLBACK.sans;
      if (/calibri|aptos|segoe|carlito|candara|corbel/.test(n)) return FALLBACK.calibre;
      if (/cambria|caladea|constantia/.test(n)) return FALLBACK.cambre;
      if (/courier|consolas|mono|lucida console|menlo/.test(n)) return FALLBACK.mono;
      if (/times|georgia|garamond|book antiqua|palatino|bookman|century schoolbook|serif|tinos|baskerville|minion/.test(n) && !/sans/.test(n)) return FALLBACK.serif;
      return FALLBACK.sans;
    }
    /* «Calibri» → «Calibri, "CPDF Calibre"» (la de respaldo solo se descarga si falta la original) */
    function fontStack(value) {
      if (!value || /CPDF /.test(value) || /^\s*var\(/.test(value)) return value;
      var parts = value.split(",").map(function (s) { return s.trim(); }).filter(Boolean);
      if (!parts.length) return value;
      var fb = '"' + fallbackFor(parts[0]) + '"';
      var gi = -1;
      for (var i = 0; i < parts.length; i++) if (GENERIC.test(parts[i].replace(/["']/g, ""))) { gi = i; break; }
      if (gi < 0) parts.push(fb); else parts.splice(gi, 0, fb);
      return parts.join(", ");
    }

    /* Viñetas de Word escritas con fuentes de símbolos (Symbol, Wingdings) → Unicode */
    var PUA = {
      "": "•", "": "▪", "": "➢", "": "✓", "": "❖", "": "■", "": "●",
      "": "□", "": "➢", "": "➢", "": "◆", "": "❑", "": "◆", "": "◆",
      "": "–", "": "—", "": "➔", "": "➔", "": "◇", "": "•", "": "○"
    };
    function fixPua(s) { return s.replace(/[-]/g, function (c) { return PUA[c] || "•"; }); }

    /* ---------------------------------------------------------------
       3. Imagen + capa de texto invisible de cada página
       --------------------------------------------------------------- */
    function rasterize(el, bg) {
      return h2c().then(function (html2canvas) {
        return html2canvas(el, { scale: SCALE, backgroundColor: bg || "#ffffff", logging: false, useCORS: true, imageTimeout: 8000 });
      });
    }
    /* Recorta el rectángulo visible de un elemento según sus contenedores con overflow oculto */
    function clipRectOf(el, root, cache) {
      if (!el || el === root) return null;
      if (cache.has(el)) return cache.get(el);
      var parent = clipRectOf(el.parentElement, root, cache), own = null;
      var cs = getComputedStyle(el);
      if (cs.overflow !== "visible" || cs.overflowX !== "visible" || cs.overflowY !== "visible") own = el.getBoundingClientRect();
      var r = parent;
      if (own) r = parent ? { left: Math.max(parent.left, own.left), top: Math.max(parent.top, own.top), right: Math.min(parent.right, own.right), bottom: Math.min(parent.bottom, own.bottom) } : own;
      cache.set(el, r);
      return r;
    }
    function textLayer(root) {
      var R = root.getBoundingClientRect(), out = [], cache = new Map();
      var walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
      var range = document.createRange(), node;
      while ((node = walker.nextNode())) {
        var txt = node.nodeValue;
        if (!txt || !/\S/.test(txt)) continue;
        var pe = node.parentElement;
        if (!pe) continue;
        var cs = getComputedStyle(pe);
        if (cs.visibility === "hidden" || cs.display === "none") continue;
        var fs = parseFloat(cs.fontSize) || 12, clip = clipRectOf(pe, root, cache);
        var re = /\S+/g, m, line = null;
        while ((m = re.exec(txt))) {
          range.setStart(node, m.index); range.setEnd(node, m.index + m[0].length);
          var rr = range.getClientRects()[0];
          if (!rr || rr.width < 0.5) continue;
          var cx = rr.left + rr.width / 2, cy = rr.top + rr.height / 2;
          if (clip && (cx < clip.left || cx > clip.right || cy < clip.top || cy > clip.bottom)) continue;
          if (cx < R.left || cx > R.right || cy < R.top || cy > R.bottom) continue;
          var box = { x: rr.left - R.left, y: rr.top - R.top, w: rr.width, h: rr.height, fs: fs, t: m[0] };
          if (line && Math.abs(line.y - box.y) < 1.5 && box.x - (line.x + line.w) < fs * 2 && box.x >= line.x) {
            line.t += " " + box.t; line.w = box.x + box.w - line.x; line.h = Math.max(line.h, box.h);
          } else { if (line) out.push(line); line = box; }
        }
        if (line) out.push(line);
      }
      return out;
    }

    /* ---------------------------------------------------------------
       Compresión de cada página: si la página tiene pocos colores (texto,
       tablas, rellenos planos) se guarda como PNG de paleta, que pesa ~3-4
       veces menos que un JPEG y no tiene manchas alrededor de las letras.
       Si tiene fotos o degradados, se usa JPEG.
       --------------------------------------------------------------- */
    var CRC = (function () { var t = new Uint32Array(256); for (var n = 0; n < 256; n++) { var c = n; for (var k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1; t[n] = c >>> 0; } return t; })();
    function crc32(u8, start, end) { var c = 0xFFFFFFFF; for (var i = start; i < end; i++) c = CRC[(c ^ u8[i]) & 255] ^ (c >>> 8); return (c ^ 0xFFFFFFFF) >>> 0; }
    function deflate(u8) {
      if (typeof CompressionStream === "undefined") return Promise.reject(new Error("sin CompressionStream"));
      return new Response(new Blob([u8]).stream().pipeThrough(new CompressionStream("deflate"))).arrayBuffer().then(function (b) { return new Uint8Array(b); });
    }
    function pngIndexed(w, h, idx, pal) {
      var raw = new Uint8Array(h * (w + 1));
      for (var y = 0; y < h; y++) raw.set(idx.subarray(y * w, (y + 1) * w), y * (w + 1) + 1);
      return deflate(raw).then(function (z) {
        var chunks = [];
        function chunk(type, data) {
          var c = new Uint8Array(12 + data.length), dv = new DataView(c.buffer);
          dv.setUint32(0, data.length);
          for (var i = 0; i < 4; i++) c[4 + i] = type.charCodeAt(i);
          c.set(data, 8);
          dv.setUint32(8 + data.length, crc32(c, 4, 8 + data.length));
          chunks.push(c);
        }
        var ihdr = new Uint8Array(13), dv = new DataView(ihdr.buffer);
        dv.setUint32(0, w); dv.setUint32(4, h); ihdr[8] = 8; ihdr[9] = 3; ihdr[10] = 0; ihdr[11] = 0; ihdr[12] = 0;
        chunk("IHDR", ihdr); chunk("PLTE", pal); chunk("IDAT", z); chunk("IEND", new Uint8Array(0));
        var len = 8 + chunks.reduce(function (s, c) { return s + c.length; }, 0), out = new Uint8Array(len), o = 8;
        out.set([137, 80, 78, 71, 13, 10, 26, 10], 0);
        chunks.forEach(function (c) { out.set(c, o); o += c.length; });
        return out;
      });
    }
    /* Paleta exacta si la página usa ≤ 256 colores (con 6 o 5 bits por canal) */
    function paletteOf(data, n) {
      var tries = [6, 5];
      for (var t = 0; t < tries.length; t++) {
        var bits = tries[t], sh = 8 - bits, size = 1 << (bits * 3), map = new Int16Array(size).fill(-1), count = 0, ok = true;
        for (var i = 0; i < n; i++) {
          var p = i * 4, key = ((data[p] >> sh) << (bits * 2)) | ((data[p + 1] >> sh) << bits) | (data[p + 2] >> sh);
          if (map[key] < 0) { if (count >= 256) { ok = false; break; } map[key] = count++; }
        }
        if (!ok) continue;
        var idx = new Uint8Array(n), sum = new Float64Array(count * 3), cnt = new Float64Array(count);
        for (i = 0; i < n; i++) {
          p = i * 4; key = ((data[p] >> sh) << (bits * 2)) | ((data[p + 1] >> sh) << bits) | (data[p + 2] >> sh);
          var k = map[key]; idx[i] = k; sum[k * 3] += data[p]; sum[k * 3 + 1] += data[p + 1]; sum[k * 3 + 2] += data[p + 2]; cnt[k]++;
        }
        var pal = new Uint8Array(count * 3);
        for (k = 0; k < count; k++) { pal[k * 3] = Math.round(sum[k * 3] / cnt[k]); pal[k * 3 + 1] = Math.round(sum[k * 3 + 1] / cnt[k]); pal[k * 3 + 2] = Math.round(sum[k * 3 + 2] / cnt[k]); }
        return { idx: idx, pal: pal };
      }
      // Muchos colores: si casi todos los píxeles caen en 256 tonos (bordes suavizados
      // del texto, por ejemplo), se usan esos 256 y el resto va al tono más cercano.
      var B = 5, S5 = 3, hist = new Uint32Array(1 << 15);
      for (i = 0; i < n; i++) { p = i * 4; hist[((data[p] >> S5) << 10) | ((data[p + 1] >> S5) << 5) | (data[p + 2] >> S5)]++; }
      var used = [];
      for (k = 0; k < hist.length; k++) if (hist[k]) used.push(k);
      used.sort(function (a, b) { return hist[b] - hist[a]; });
      var top = used.slice(0, 256), cover = 0;
      top.forEach(function (q) { cover += hist[q]; });
      if (cover / n < 0.985) return null;   // fotos o degradados: mejor JPEG
      var slot = new Int16Array(1 << 15).fill(-1);
      top.forEach(function (q, j) { slot[q] = j; });
      function rgbOf(q) { return [((q >> 10) & 31) * 8 + 4, ((q >> 5) & 31) * 8 + 4, (q & 31) * 8 + 4]; }
      var tops = top.map(rgbOf);
      used.slice(256).forEach(function (q) {
        var c = rgbOf(q), best = 0, bd = Infinity;
        for (var j = 0; j < tops.length; j++) { var t = tops[j], d = (t[0] - c[0]) * (t[0] - c[0]) * 3 + (t[1] - c[1]) * (t[1] - c[1]) * 4 + (t[2] - c[2]) * (t[2] - c[2]) * 2; if (d < bd) { bd = d; best = j; } }
        slot[q] = best;
      });
      var cnt2 = top.length, idx2 = new Uint8Array(n), sum2 = new Float64Array(cnt2 * 3), c2 = new Float64Array(cnt2);
      for (i = 0; i < n; i++) {
        p = i * 4; var q = ((data[p] >> S5) << 10) | ((data[p + 1] >> S5) << 5) | (data[p + 2] >> S5), j = slot[q];
        idx2[i] = j;
        if (hist[q] && slot[q] === j && top[j] === q) { sum2[j * 3] += data[p]; sum2[j * 3 + 1] += data[p + 1]; sum2[j * 3 + 2] += data[p + 2]; c2[j]++; }
      }
      var pal2 = new Uint8Array(cnt2 * 3);
      for (k = 0; k < cnt2; k++) { var cc = c2[k] || 1; pal2[k * 3] = Math.round(sum2[k * 3] / cc); pal2[k * 3 + 1] = Math.round(sum2[k * 3 + 1] / cc); pal2[k * 3 + 2] = Math.round(sum2[k * 3 + 2] / cc); }
      return { idx: idx2, pal: pal2 };
    }
    function encodePage(canvas) {
      var w = canvas.width, h = canvas.height;
      try {
        var ctx = canvas.getContext("2d", { willReadFrequently: true });
        var data = ctx.getImageData(0, 0, w, h).data;
        var p = paletteOf(data, w * h);
        if (p && typeof CompressionStream !== "undefined") {
          return pngIndexed(w, h, p.idx, p.pal).then(function (png) { return { type: "PNG", data: png }; })
            .catch(function () { return { type: "JPEG", data: canvas.toDataURL("image/jpeg", 0.85) }; });
        }
      } catch (e) { /* lienzo no legible: se usa JPEG */ }
      return Promise.resolve({ type: "JPEG", data: canvas.toDataURL("image/jpeg", 0.85) });
    }

    /* Crea el PDF: una imagen por página + texto invisible encima */
    function PdfWriter(sampleText) {
      var self = this;
      this.n = 0;
      this.ready = K.pdfEngines(sampleText).then(function (eng) { self.eng = eng; });
    }
    PdfWriter.prototype.add = function (canvas, wPx, hPx, texts) {
      var wPt = wPx * PX2PT, hPt = hPx * PX2PT, orient = wPt > hPt ? "landscape" : "portrait";
      if (this.n === 0) {
        // El documento se crea con la medida exacta de la primera página
        var made = K.newPdf(this.eng[0], this.eng[1], { format: [wPt, hPt], orientation: orient });
        this.doc = made.doc; this.FONT = made.FONT;
      } else this.doc.addPage([wPt, hPt], orient);
      var doc = this.doc, FONT = this.FONT;
      this.n++;
      return encodePage(canvas).then(function (img) {
      canvas.width = canvas.height = 0;   // libera memoria
      doc.addImage(img.data, img.type, 0, 0, wPt, hPt, undefined, "FAST");
      if (texts && texts.length) {
        doc.setFont(FONT, "normal");
        texts.forEach(function (t) {
          var str = FONT === "helvetica" ? t.t.replace(K.WINANSI_G, "?") : t.t;
          var size = Math.max(1, t.fs * PX2PT);
          doc.setFontSize(size);
          var natural = doc.getTextWidth(str);
          var hs = natural > 0 ? (t.w * PX2PT) / natural : 1;
          var base = (t.y + t.h / 2 + t.fs * 0.32) * PX2PT;
          try {
            doc.text(str, t.x * PX2PT, base, { renderingMode: "invisible", horizontalScale: Math.max(0.2, Math.min(5, hs)), baseline: "alphabetic" });
          } catch (e) { /* un carácter raro no debe romper el documento */ }
        });
      }
      });
    };
    PdfWriter.prototype.blob = function () {
      this.doc.setProperties({ creator: K.BRAND.name, producer: K.BRAND.name });
      return new Blob([this.doc.output("arraybuffer")], { type: "application/pdf" });
    };

    /* Recorre páginas ya construidas: cada una se dibuja, se pasa al PDF y se libera */
    function pagesToPdf(pages, onProgress, from, to) {
      var sample = pages.map(function (p) { return p.sample || ""; }).join(" ");
      var w = new PdfWriter(sample), i = 0;
      return w.ready.then(function step() {
        if (i >= pages.length) return w.blob();
        var p = pages[i++];
        return p.render().then(function (r) {
          return w.add(r.canvas, r.w, r.h, r.texts);
        }).then(function () {
          if (onProgress) onProgress(from + (to - from) * (i / pages.length));
        }).then(K.yieldUI).then(step);
      });
    }
    function pagesToImages(pages, onProgress, from, to, type, quality) {
      var out = [], i = 0;
      return (function step() {
        if (i >= pages.length) return Promise.resolve(out);
        var p = pages[i++];
        return p.render({ noText: true }).then(function (r) {
          return new Promise(function (ok) { r.canvas.toBlob(ok, type, quality); }).then(function (b) {
            out.push({ blob: b, w: r.canvas.width, h: r.canvas.height });
            r.canvas.width = r.canvas.height = 0;
            if (onProgress) onProgress(from + (to - from) * (i / pages.length));
          });
        }).then(K.yieldUI).then(step);
      })();
    }

    /* ---------------------------------------------------------------
       4. WORD (.docx) → páginas
       --------------------------------------------------------------- */
    var W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    var TOK_PAGE = "⟦P⟧", TOK_TOTAL = "⟦N⟧";

    /* Los números de página de encabezados y pies son campos de Word: se
       cambian por una marca que luego se reemplaza por el número real. */
    function markPageFields(xmlText) {
      if (!/PAGE|NUMPAGES|SECTIONPAGES/.test(xmlText)) return null;
      var doc = new DOMParser().parseFromString(xmlText, "application/xml");
      if (doc.getElementsByTagName("parsererror").length) return null;
      var changed = false;
      function tokenFor(instr) {
        var m = /^\s*(PAGE|NUMPAGES|SECTIONPAGES)\b/i.exec(instr || "");
        return m ? (m[1].toUpperCase() === "PAGE" ? TOK_PAGE : TOK_TOTAL) : null;
      }
      function makeRun(rPr, text) {
        var r = doc.createElementNS(W_NS, "w:r");
        if (rPr) r.appendChild(rPr.cloneNode(true));
        var t = doc.createElementNS(W_NS, "w:t");
        t.setAttribute("xml:space", "preserve"); t.textContent = text;
        r.appendChild(t);
        return r;
      }
      // Campos simples
      Array.prototype.slice.call(doc.getElementsByTagNameNS(W_NS, "fldSimple")).forEach(function (f) {
        var tok = tokenFor(f.getAttributeNS(W_NS, "instr") || f.getAttribute("w:instr"));
        if (!tok) return;
        var rPr = f.getElementsByTagNameNS(W_NS, "rPr")[0] || null;
        f.parentNode.replaceChild(makeRun(rPr, tok), f);
        changed = true;
      });
      // Campos complejos: begin · instrText · separate · resultado · end
      var runs = Array.prototype.slice.call(doc.getElementsByTagNameNS(W_NS, "r"));
      var stack = [];
      runs.forEach(function (r) {
        var fc = r.getElementsByTagNameNS(W_NS, "fldChar")[0];
        var type = fc ? (fc.getAttributeNS(W_NS, "fldCharType") || fc.getAttribute("w:fldCharType")) : null;
        if (type === "begin") { stack.push({ instr: "", result: [], phase: "instr", rPr: r.getElementsByTagNameNS(W_NS, "rPr")[0] || null, last: r }); return; }
        var top = stack[stack.length - 1];
        if (!top) return;
        if (type === "separate") { top.phase = "result"; top.last = r; return; }
        if (type === "end") {
          stack.pop();
          var tok = tokenFor(top.instr);
          if (!tok || stack.length) return;
          top.result.forEach(function (x) { if (x.parentNode) x.parentNode.removeChild(x); });
          r.parentNode.insertBefore(makeRun(top.rPr, tok), r);
          changed = true;
          return;
        }
        var it = r.getElementsByTagNameNS(W_NS, "instrText")[0];
        if (top.phase === "instr" && it) { top.instr += it.textContent; if (!top.rPr) top.rPr = r.getElementsByTagNameNS(W_NS, "rPr")[0] || null; }
        else if (top.phase === "result") top.result.push(r);
      });
      return changed ? new XMLSerializer().serializeToString(doc) : null;
    }
    function prepareDocx(buf) {
      return K.Engines.jszip().then(function (JSZip) { return JSZip.loadAsync(buf); }).then(function (zip) {
        if (!zip.file("word/document.xml")) throw K.UserError("No es un documento de Word válido", "El archivo no contiene un documento .docx. Si es un .doc antiguo, ábrelo y guárdalo como .docx.");
        var parts = Object.keys(zip.files).filter(function (n) { return /^word\/(header|footer)\d*\.xml$/.test(n); });
        return Promise.all(parts.map(function (n) {
          return zip.file(n).async("string").then(function (s) { var out = markPageFields(s); if (out) zip.file(n, out); });
        })).then(function () { return zip.generateAsync({ type: "arraybuffer" }); });
      });
    }

    /* Ajustes al HTML que genera docx-preview para que se parezca a Word */
    function polishDocxStyles(host) {
      Array.prototype.forEach.call(host.querySelectorAll("style"), function (st) {
        var css = st.textContent;
        css = css.replace(/(--docx-[\w-]+-font\s*:\s*)([^;}]+)/g, function (m, a, v) { return a + fontStack(v.trim()); });
        css = css.replace(/(font-family\s*:\s*)([^;}]+)/g, function (m, a, v) {
          if (/^\s*(Symbol|Wingdings[^;]*|Webdings)\s*$/i.test(v)) return a + "inherit";
          return a + fontStack(v.trim());
        });
        // Viñetas y numeración: el texto empieza en la sangría francesa, como en Word
        var hang = {};
        css.replace(/([^{}]*docx-num-[\w-]+)\s*\{([^}]*)\}/g, function (m, sel, body) {
          var ti = /text-indent:\s*(-?[\d.]+)pt/.exec(body);
          if (ti && !/:before/.test(sel)) hang[sel.trim()] = -parseFloat(ti[1]);
          return m;
        });
        css = css.replace(/([^{}]*docx-num-[\w-]+):before\s*\{([^}]*)\}/g, function (m, sel, body) {
          var h = hang[sel.trim()];
          body = fixPua(body).replace(/\\9\s*"/g, '"').replace(/\\9/g, "");
          var extra = h > 0 ? "display:inline-block;min-width:" + h + "pt;text-indent:0;" : "padding-right:.6em;";
          return sel + ":before {" + body + ";" + extra + "}";
        });
        st.textContent = css;
      });
      Array.prototype.forEach.call(host.querySelectorAll("[style*='font-family']"), function (el) {
        var ff = el.style.fontFamily;
        if (/^\s*["']?(Symbol|Wingdings|Webdings)/i.test(ff)) el.style.fontFamily = "";
        else if (ff) el.style.fontFamily = fontStack(ff);
      });
      // Imágenes «delante/detrás del texto» (marcas de agua, logos del membrete): docx-preview
      // las mete en una caja de 0×0 alineada a la línea base, y la imagen estiraba la línea
      // hasta cientos de píxeles (el contenido bajaba y el documento se partía en decenas de
      // páginas). Se alinean arriba para que no ocupen espacio en el texto, como en Word.
      Array.prototype.forEach.call(host.querySelectorAll("div[style*='inline-block'], div[style*='position: relative']"), function (d) {
        var st = d.style;
        if ((st.width === "0px" || st.height === "0px") && (st.position === "relative" || st.display === "inline-block")) {
          st.verticalAlign = "top"; st.lineHeight = "0"; st.overflow = "visible";
        }
      });
      // Símbolos sueltos dentro del texto (casillas, flechas de Wingdings…)
      var walker = document.createTreeWalker(host, NodeFilter.SHOW_TEXT, null), n;
      while ((n = walker.nextNode())) if (/[-]/.test(n.nodeValue)) n.nodeValue = fixPua(n.nodeValue);
    }

    /* Imágenes y formas «flotantes» (marca de agua, logotipos del membrete, cuadros
       de texto): en Word NO ocupan espacio. docx-preview las mete en una caja de
       tamaño 0 que, en algunos navegadores, alinea su línea base con el borde inferior
       de la imagen y empuja el texto cientos de píxeles hacia abajo. Se sacan del flujo. */
    function fixFloatingDrawings(root) {
      Array.prototype.forEach.call(root.querySelectorAll("div, span"), function (d) {
        var st = d.style;
        if (st.width !== "0px" || st.height !== "0px") return;
        st.verticalAlign = "top"; st.lineHeight = "0"; st.overflow = "visible";
        if (!st.position || st.position === "static") st.position = "relative";
        Array.prototype.forEach.call(d.children, function (c) {
          var p = c.style.position;
          if (!p || p === "relative" || p === "static") {
            c.style.position = "absolute";
            if (!c.style.left) c.style.left = "0";
            if (!c.style.top) c.style.top = "0";
          }
        });
      });
      // Cuadros de texto VML (Word los guarda así para compatibilidad): docx-preview los
      // dibuja como <svg><foreignObject>, que pierde fuentes y colores y recorta el texto.
      // Se convierten en cajas HTML normales, en la misma posición y tamaño.
      Array.prototype.forEach.call(root.querySelectorAll("svg"), function (svg) {
        var fo = svg.querySelector("foreignObject");
        if (!fo || !/\S/.test(fo.textContent || "")) return;
        var box = document.createElement("div");
        var st = svg.getAttribute("style") || "";
        // Solo propiedades de posición y tamaño (las «mso-…» no son CSS)
        st.split(";").forEach(function (decl) {
          var i = decl.indexOf(":");
          if (i < 0) return;
          var k = decl.slice(0, i).trim().toLowerCase(), v = decl.slice(i + 1).trim();
          if (/^(position|margin-left|margin-top|left|top|width|height|z-index|rotation)$/.test(k) && k !== "rotation") box.style.setProperty(k, k === "z-index" ? "auto" : v);
        });
        if (!box.style.position) box.style.position = "relative";
        box.style.boxSizing = "border-box";
        box.style.padding = "3.6pt 7.2pt";   // márgenes internos por defecto de un cuadro de texto de Word
        box.style.overflow = "visible";
        box.style.lineHeight = "normal";
        while (fo.firstChild) box.appendChild(fo.firstChild);
        svg.parentNode.replaceChild(box, svg);
        // La posición del cuadro se mide desde el inicio del párrafo (como en Word),
        // no desde el lugar donde el navegador lo dejaría dentro de la línea.
        var holder = box.parentElement, para = box.closest("p");
        if (holder && holder.tagName === "DIV" && !holder.style.position && para) {
          holder.style.position = "absolute"; holder.style.left = "0"; holder.style.top = "0";
          holder.style.width = "0"; holder.style.height = "0"; holder.style.lineHeight = "0";
          if (getComputedStyle(para).position === "static") para.style.position = "relative";
        }
      });
      // Líneas y formas VML: misma regla (posición desde el inicio del párrafo)
      Array.prototype.forEach.call(root.querySelectorAll("svg"), function (svg) {
        if (svg.style.position !== "absolute") return;
        // Una línea VML no trae tamaño: se deja ver completa
        // (se le da el tamaño justo para contenerla: el dibujo de la página recorta lo que sobresale)
        if (!svg.style.width) {
          var mx = 1, my = 1;
          Array.prototype.forEach.call(svg.querySelectorAll("line"), function (l) {
            mx = Math.max(mx, pxFromCss(l.getAttribute("x1")), pxFromCss(l.getAttribute("x2")));
            my = Math.max(my, pxFromCss(l.getAttribute("y1")), pxFromCss(l.getAttribute("y2")));
          });
          svg.style.width = (mx + 2) + "px"; svg.style.height = (my + 2) + "px"; svg.style.overflow = "visible";
        }
        Array.prototype.forEach.call(svg.querySelectorAll("line"), function (l) { if (!l.getAttribute("stroke") || l.getAttribute("stroke") === "null") l.setAttribute("stroke", "#000"); });
        var host = svg.parentElement, para = svg.closest("p");
        if (host && host.tagName === "DIV" && !host.style.position && para) {
          host.style.position = "absolute"; host.style.left = "0"; host.style.top = "0";
          host.style.width = "0"; host.style.height = "0"; host.style.lineHeight = "0";
          if (getComputedStyle(para).position === "static") para.style.position = "relative";
        }
      });
    }

    /* Altura de una línea «sencilla» en Word según la fuente (proporción del tamaño).
       Word calcula el interlineado con estas medidas; el navegador no, por eso se ajusta. */
    var LINE = { "calibri": 1.2207, "calibri light": 1.2207, "carlito": 1.2207, "cambria": 1.1724, "caladea": 1.1724,
      "times new roman": 1.15, "times": 1.15, "arial": 1.15, "helvetica": 1.15, "arial narrow": 1.15, "liberation sans": 1.15,
      "liberation serif": 1.15, "courier new": 1.1328, "georgia": 1.1362, "verdana": 1.2153, "tahoma": 1.2075,
      "segoe ui": 1.33, "aptos": 1.2, "aptos display": 1.2, "garamond": 1.1206, "century gothic": 1.2251, "trebuchet ms": 1.1611,
      "book antiqua": 1.1709, "palatino linotype": 1.1709, "comic sans ms": 1.3936, "arial black": 1.41, "franklin gothic book": 1.13,
      "century": 1.2, "bookman old style": 1.17, "lucida sans": 1.16, "consolas": 1.17 };
    function firstFamily(ff) { return String(ff || "").split(",")[0].replace(/["']/g, "").trim().toLowerCase(); }
    function styleKey(p) { return (p.className || "").split(/\s+/).filter(function (c) { return c && c.indexOf("-num-") < 0; }).sort().join(" "); }
    function wordParagraphs(root, host) {
      // Los párrafos sin estilo usan el estilo «Normal» del documento (docx-preview no lo aplica a su texto)
      var css = Array.prototype.map.call(host.querySelectorAll("style"), function (st) { return st.textContent; }).join("\n");
      var def = /\.docx p,\s*p\.(docx_[\w-]+) span/.exec(css);
      if (def) Array.prototype.forEach.call(root.querySelectorAll("p"), function (p) {
        if (!/(^|\s)docx_/.test(p.className)) p.classList.add(def[1]);
      });
      // «No agregar espacio entre párrafos del mismo estilo» (listas, por ejemplo)
      Array.prototype.forEach.call(root.querySelectorAll("p"), function (p) {
        var q = p.nextElementSibling;
        if (!q || q.tagName !== "P" || styleKey(p) !== styleKey(q)) return;
        if (getComputedStyle(p).getPropertyValue("--cpdf-ctx").trim() === "1") p.style.marginBottom = "0";
        if (getComputedStyle(q).getPropertyValue("--cpdf-ctx").trim() === "1") q.style.marginTop = "0";
      });
      Array.prototype.forEach.call(root.querySelectorAll("p"), function (p) {
        // Fuente dominante del párrafo (la de mayor tamaño)
        var best = null, bestSize = 0;
        Array.prototype.forEach.call(p.querySelectorAll("span"), function (sp) {
          if (!sp.firstChild || sp.firstChild.nodeType !== 3 || !/\S/.test(sp.textContent)) return;
          var cs = getComputedStyle(sp), fsz = parseFloat(cs.fontSize) || 0;
          if (fsz > bestSize) { bestSize = fsz; best = cs.fontFamily; }
        });
        if (!best) best = getComputedStyle(p).fontFamily;
        else {
          // Como en Word, la altura de la línea la marca el texto (no el tamaño del párrafo vacío)
          p.style.fontSize = bestSize + "px";
          p.style.fontFamily = best;
        }
        p.style.setProperty("--cpdf-lh", String(LINE[firstFamily(best)] || 1.17));
      });
    }

    /* Puntos de corte de una sección que no cabe en una página.
       Reglas de Word: no cortar dentro de una línea, «conservar con el siguiente»
       (títulos), «conservar líneas juntas» y control de líneas viudas y huérfanas. */
    function computeCuts(article, limit) {
      var A = article.getBoundingClientRect(), total = article.scrollHeight;
      if (total <= limit + 1) return [0, total];
      var atoms = [], info = new Map();
      function paraInfo(p) {
        if (!p) return { kn: false, kl: false };
        if (info.has(p)) return info.get(p);
        var cs = getComputedStyle(p);
        var r = { kn: cs.getPropertyValue("--cpdf-kn").trim() === "1" || /docx_(heading|title|subtitle)/i.test(p.className),
                  kl: cs.getPropertyValue("--cpdf-kl").trim() === "1" };
        info.set(p, r);
        return r;
      }
      var walker = document.createTreeWalker(article, NodeFilter.SHOW_TEXT, null), range = document.createRange(), node;
      while ((node = walker.nextNode())) {
        if (!/\S/.test(node.nodeValue)) continue;
        range.selectNodeContents(node);
        var para = node.parentElement && node.parentElement.closest("p, li, td, th");
        Array.prototype.forEach.call(range.getClientRects(), function (r) {
          if (r.height > 0) atoms.push({ top: r.top - A.top, bottom: r.bottom - A.top, para: para });
        });
      }
      Array.prototype.forEach.call(article.querySelectorAll("img, svg, canvas, hr"), function (el) {
        var r = el.getBoundingClientRect();
        if (r.height > 0) atoms.push({ top: r.top - A.top, bottom: r.bottom - A.top, para: el.closest("p, td") || el });
      });
      atoms.sort(function (a, b) { return a.top - b.top; });
      function linesOf(p, from, to) {
        var tops = {};
        atoms.forEach(function (a) { if (a.para === p && a.top >= from - 0.5 && a.top < to - 0.5) tops[Math.round(a.top)] = a.top; });
        return Object.keys(tops).map(function (k) { return tops[k]; }).sort(function (a, b) { return a - b; });
      }
      function firstTop(p) { var t = Infinity; atoms.forEach(function (a) { if (a.para === p) t = Math.min(t, a.top); }); return t; }
      var cuts = [0], start = 0, guard = 0;
      while (total - start > limit + 1 && guard++ < 3000) {
        var lim = start + limit, cut = lim, minCut = start + limit * 0.3;
        atoms.forEach(function (a) { if (a.top < lim - 0.5 && a.bottom > lim + 0.5 && a.top > start + 1) cut = Math.min(cut, a.top); });
        // Párrafo partido por el corte
        var after = atoms.filter(function (a) { return a.top >= cut - 0.5; })[0];
        var p = after && after.para;
        if (p && p.tagName === "P") {
          var pi = paraInfo(p), linesBefore = linesOf(p, start, cut), linesAfter = linesOf(p, cut, Infinity);
          if (linesBefore.length) {
            if (pi.kl && linesBefore[0] > minCut) cut = linesBefore[0];                         // conservar líneas juntas
            else if (linesBefore.length === 1 && linesBefore[0] > minCut) cut = linesBefore[0]; // huérfana
            else if (linesAfter.length === 1 && linesBefore.length >= 3) cut = linesBefore[linesBefore.length - 1]; // viuda
          }
        }
        // «Conservar con el siguiente»: los títulos bajan junto con el párrafo que los sigue
        for (var k = 0; k < 6; k++) {
          var prev = atoms.filter(function (a) { return a.bottom <= cut + 0.5 && a.top >= start; }).pop();
          if (!prev || !prev.para || !paraInfo(prev.para).kn) break;
          var t = firstTop(prev.para);
          if (t <= minCut || t >= cut) break;
          cut = t;
        }
        if (cut <= start + 1) cut = lim;
        cuts.push(cut); start = cut;
      }
      cuts.push(total);
      return cuts;
    }

    function pxFromCss(v) {
      var m = /([\d.]+)\s*(pt|px|in|cm|mm)?/.exec(v || "");
      if (!m) return 0;
      var n = parseFloat(m[1]), u = m[2] || "px";
      return u === "pt" ? n * 96 / 72 : u === "in" ? n * 96 : u === "cm" ? n * 96 / 2.54 : u === "mm" ? n * 96 / 25.4 : n;
    }

    /* Prepara la lista de páginas de un .docx (cada una sabe dibujarse sola) */
    function docxPages(buf, onProgress, dbg) {
      ensureFontFaces();
      var host = makeHost(), styleBox = document.createElement("div");
      host.appendChild(styleBox);
      var body = document.createElement("div");
      host.appendChild(body);
      return Promise.all([docxLib(), prepareDocx(buf)]).then(function (r) {
        if (onProgress) onProgress(0.08);
        return r[0].renderAsync(r[1], body, styleBox, {
          inWrapper: false, ignoreLastRenderedPageBreak: false, breakPages: true, experimental: true,
          useBase64URL: true, renderComments: false, renderChanges: false, className: "docx"
        });
      }).then(function () {
        polishDocxStyles(host);
        fixFloatingDrawings(body);
        wordParagraphs(body, host);
        return settle(host);
      }).then(function () {
        if (onProgress) onProgress(0.15);
        var sections = Array.prototype.slice.call(body.querySelectorAll(":scope > section.docx"));
        if (!sections.length) throw K.UserError("El documento está vacío", "No se encontró contenido para convertir.");
        // Páginas vacías falsas: cuando un párrafo tiene «salto de página anterior» y además
        // la marca de salto que guardó Word, docx-preview corta dos veces y deja una página
        // en blanco. Se quitan las secciones sin contenido (si no es la única).
        sections = sections.filter(function (s, idx) {
          var art = s.querySelector(":scope > article");
          var empty = art && !/\S/.test(art.textContent || "") && !art.querySelector("img, svg, table, canvas, hr, image");
          if (empty && sections.length > 1) { s.remove(); return false; }
          return true;
        });
        var plan = [];
        // Como Word, no se deja el «espacio anterior» del primer párrafo al inicio de una página nueva
        sections.forEach(function (s, idx) {
          if (!idx) return;
          var first = s.querySelector(":scope > article p, :scope > article table");
          if (first && first.tagName === "P" && !first.previousElementSibling) first.style.marginTop = "0";
        });
        sections.forEach(function (s) {
          var cs = getComputedStyle(s);
          var W = s.offsetWidth;
          var H = pxFromCss(s.style.minHeight) || s.offsetHeight;
          var art = s.querySelector(":scope > article");
          var hd = s.querySelector(":scope > header"), ft = s.querySelector(":scope > footer");
          var padT = parseFloat(cs.paddingTop) || 0, padB = parseFloat(cs.paddingBottom) || 0;
          var normal = H - padT - padB;   // alto del cuerpo según los márgenes de la página
          // Seguro: un encabezado o pie no puede quitarle al texto más de lo razonable.
          // Si lo hace (objeto mal medido), se dibuja encima sin desplazar el contenido.
          if (art && hd && art.offsetTop > padT + normal * 0.35) {
            hd.style.position = "absolute"; hd.style.left = cs.paddingLeft; hd.style.right = cs.paddingRight;
            hd.style.top = Math.max(0, padT + (parseFloat(getComputedStyle(hd).marginTop) || 0)) + "px"; hd.style.marginTop = "0";
          }
          var footReach = 0;
          if (ft) {
            var fcs = getComputedStyle(ft), fd = padB + (parseFloat(fcs.marginBottom) || 0);
            footReach = fd + ft.offsetHeight;
            if (footReach - padB > normal * 0.35) {
              ft.style.position = "absolute"; ft.style.left = cs.paddingLeft; ft.style.right = cs.paddingRight;
              ft.style.bottom = Math.max(0, fd) + "px"; ft.style.marginBottom = "0";
              footReach = 0;
            }
          }
          // Como Word: el texto termina donde empieza el pie si este invade el margen inferior
          var bodyBottom = H - Math.max(padB, footReach);
          var limit = art ? bodyBottom - art.offsetTop : H;
          if (limit < normal * 0.5) limit = normal;
          var cuts = art ? computeCuts(art, Math.max(40, limit)) : [0, 0];
          var bg = cs.backgroundColor;
          if (!bg || bg === "rgba(0, 0, 0, 0)" || bg === "transparent") bg = "#ffffff";
          for (var k = 0; k < cuts.length - 1; k++) plan.push({ s: s, W: W, H: H, from: cuts[k], to: cuts[k + 1], bg: bg });
          if (cuts.length < 2) plan.push({ s: s, W: W, H: H, from: 0, to: 0, bg: bg });
        });
        // Las secciones originales ya no se necesitan en pantalla
        if (!dbg) sections.forEach(function (s) { s.remove(); });
        var total = plan.length;
        var sample = body.textContent || "";
        var pages = plan.map(function (pg, idx) {
          return {
            sample: idx === 0 ? sample : "",
            render: function (opt) {
              var c = pg.s.cloneNode(true);
              c.style.height = pg.H + "px"; c.style.minHeight = pg.H + "px"; c.style.maxHeight = pg.H + "px";
              c.style.overflow = "hidden"; c.style.background = pg.bg; c.style.boxShadow = "none"; c.style.margin = "0";
              var art = c.querySelector(":scope > article");
              if (art && (pg.from > 0 || pg.to < art.scrollHeight || true)) {
                var inner = document.createElement("div");
                inner.style.display = "flow-root"; inner.style.marginTop = (-pg.from) + "px";
                while (art.firstChild) inner.appendChild(art.firstChild);
                art.appendChild(inner);
                art.style.height = Math.max(1, pg.to - pg.from) + "px";
                art.style.flex = "none"; art.style.overflow = "hidden";
              }
              // Números de página reales
              var tw = document.createTreeWalker(c, NodeFilter.SHOW_TEXT, null), t;
              while ((t = tw.nextNode())) {
                if (t.nodeValue.indexOf("⟦") >= 0) t.nodeValue = t.nodeValue.split(TOK_PAGE).join(String(idx + 1)).split(TOK_TOTAL).join(String(total));
              }
              body.appendChild(c);
              return settle(c).then(function () {
                return rasterize(c, pg.bg);
              }).then(function (canvas) {
                var texts = opt && opt.noText ? null : textLayer(c);
                var w = c.offsetWidth, h = c.offsetHeight;
                c.remove();
                return { canvas: canvas, texts: texts, w: w, h: h };
              });
            }
          };
        });
        pages.cleanup = function () { host.remove(); };
        return pages;
      }).catch(function (e) { host.remove(); throw e; });
    }

    O.docxToPdf = function (buf, onProgress) {
      return docxPages(buf, onProgress).then(function (pages) {
        return pagesToPdf(pages, onProgress, 0.2, 1).then(function (blob) { pages.cleanup(); return { blob: blob, pages: pages.length }; },
          function (e) { pages.cleanup(); throw e; });
      });
    };
    O.docxToImages = function (buf, onProgress, type, quality) {
      return docxPages(buf, onProgress).then(function (pages) {
        return pagesToImages(pages, onProgress, 0.2, 1, type, quality).then(function (r) { pages.cleanup(); return r; },
          function (e) { pages.cleanup(); throw e; });
      });
    };

    /* ---------------------------------------------------------------
       5. Utilidades XML y de color (Excel y PowerPoint)
       --------------------------------------------------------------- */
    function parseX(str) { return new DOMParser().parseFromString(str, "application/xml"); }
    function kidsOf(el, name) {
      var out = [];
      if (!el) return out;
      for (var c = el.firstElementChild; c; c = c.nextElementSibling) if (!name || c.localName === name) out.push(c);
      return out;
    }
    function kidOf(el) {
      for (var i = 1; i < arguments.length && el; i++) el = kidsOf(el, arguments[i])[0] || null;
      return el || null;
    }
    function descOf(el, name) { return el ? Array.prototype.slice.call(el.getElementsByTagNameNS("*", name)) : []; }
    var NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    function relId(el, name) { return el ? (el.getAttributeNS(NS_R, name) || el.getAttribute("r:" + name)) : null; }
    function attrNum(el, name, def) { if (!el) return def; var v = el.getAttribute(name); if (v == null || v === "") return def; var n = parseFloat(v); return isNaN(n) ? def : n; }
    function attrBool(el, name, def) { if (!el) return def; var v = el.getAttribute(name); if (v == null) return def; return v === "1" || v === "true" || v === "on"; }
    function partDir(p) { return p.slice(0, p.lastIndexOf("/") + 1); }
    function resolveTarget(base, target) {
      if (!target) return null;
      if (target.charAt(0) === "/") return target.slice(1);
      var out = [];
      (partDir(base) + target).split("/").forEach(function (s) { if (s === "..") out.pop(); else if (s && s !== ".") out.push(s); });
      return out.join("/");
    }
    function relsOf(zip, partPath) {
      var f = zip.file(partDir(partPath) + "_rels/" + partPath.slice(partPath.lastIndexOf("/") + 1) + ".rels");
      if (!f) return Promise.resolve({});
      return f.async("string").then(function (s) {
        var map = {};
        kidsOf(parseX(s).documentElement, "Relationship").forEach(function (r) {
          map[r.getAttribute("Id")] = { type: r.getAttribute("Type") || "", target: resolveTarget(partPath, r.getAttribute("Target")), external: r.getAttribute("TargetMode") === "External" };
        });
        return map;
      });
    }
    function relByType(rels, suffix) { for (var k in rels) if (rels[k].type.slice(-suffix.length) === suffix) return rels[k]; return null; }
    function xmlPart(zip, path) { var f = path && zip.file(path); return f ? f.async("string").then(parseX) : Promise.resolve(null); }
    function hex2rgb(h) { h = String(h || "").replace("#", ""); if (h.length === 8) h = h.slice(2); if (h.length !== 6) return null; return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
    function rgb2hsl(c) {
      var r = c[0] / 255, g = c[1] / 255, b = c[2] / 255, mx = Math.max(r, g, b), mn = Math.min(r, g, b), h = 0, s = 0, l = (mx + mn) / 2;
      if (mx !== mn) { var d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn); h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
      return [h, s, l];
    }
    function hsl2rgb(x) {
      var h = x[0], s = x[1], l = x[2];
      if (!s) { var v = Math.round(l * 255); return [v, v, v]; }
      function f(p, q, t) { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; }
      var q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
      return [Math.round(f(p, q, h + 1 / 3) * 255), Math.round(f(p, q, h) * 255), Math.round(f(p, q, h - 1 / 3) * 255)];
    }
    function css(c, a) { return c ? (a != null && a < 1 ? "rgba(" + c.join(",") + "," + a + ")" : "rgb(" + c.join(",") + ")") : ""; }
    function escH(s) { return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;"); }
    function themeColors(themeDoc) {
      var t = {};
      var cs = themeDoc && descOf(themeDoc.documentElement, "clrScheme")[0];
      kidsOf(cs).forEach(function (c) {
        var v = kidOf(c, "srgbClr"), s = kidOf(c, "sysClr");
        t[c.localName] = v ? hex2rgb(v.getAttribute("val")) : s ? hex2rgb(s.getAttribute("lastClr") || (s.getAttribute("val") === "window" ? "FFFFFF" : "000000")) : null;
      });
      var fonts = {};
      var fs = themeDoc && descOf(themeDoc.documentElement, "fontScheme")[0];
      if (fs) {
        var mj = kidOf(fs, "majorFont", "latin"), mn = kidOf(fs, "minorFont", "latin");
        fonts.major = mj && mj.getAttribute("typeface"); fonts.minor = mn && mn.getAttribute("typeface");
      }
      t.$fonts = fonts;
      return t;
    }

    /* ---------------------------------------------------------------
       6. EXCEL (.xlsx) → páginas, como las imprime Excel
       --------------------------------------------------------------- */
    var XL_INDEXED = ["000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF", "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
      "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080", "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
      "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF", "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
      "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696", "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333"];
    var PAPER = { 1: [8.5, 11], 3: [11, 17], 5: [8.5, 14], 7: [7.25, 10.5], 8: [11.69, 16.54], 9: [8.27, 11.69], 11: [5.83, 8.27], 13: [7.17, 10.12], 14: [8.5, 13] };
    function xlColor(el, theme, palette) {
      if (!el) return null;
      if (el.getAttribute("auto") === "1") return null;
      var rgb = null;
      if (el.getAttribute("rgb")) rgb = hex2rgb(el.getAttribute("rgb"));
      else if (el.getAttribute("theme") != null) {
        var order = ["lt1", "dk1", "lt2", "dk2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];
        rgb = theme[order[parseInt(el.getAttribute("theme"), 10)]] || null;
      } else if (el.getAttribute("indexed") != null) {
        var ix = parseInt(el.getAttribute("indexed"), 10);
        rgb = ix === 64 ? [0, 0, 0] : ix === 65 ? [255, 255, 255] : hex2rgb((palette || XL_INDEXED)[ix]);
      }
      if (!rgb) return null;
      var tint = attrNum(el, "tint", 0);
      if (tint) {
        var h = rgb2hsl(rgb);
        h[2] = tint < 0 ? h[2] * (1 + tint) : h[2] * (1 - tint) + tint;
        rgb = hsl2rgb(h);
      }
      return rgb;
    }
    function colLetters(s) { var n = 0; for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64); return n - 1; }
    function cellRef(ref) { var m = /^\$?([A-Z]+)\$?(\d+)$/.exec(ref); return m ? { c: colLetters(m[1]), r: parseInt(m[2], 10) - 1 } : null; }
    function rangeRef(ref) {
      var p = String(ref || "").replace(/^'?[^!]*'?!/, "").split(":");
      var a = cellRef(p[0]), b = cellRef(p[1] || p[0]);
      return a && b ? { s: { r: Math.min(a.r, b.r), c: Math.min(a.c, b.c) }, e: { r: Math.max(a.r, b.r), c: Math.max(a.c, b.c) } } : null;
    }
    function xlBorderCss(side, theme, palette) {
      if (!side) return "";
      var st = side.getAttribute("style");
      if (!st || st === "none") return "";
      var col = css(xlColor(kidOf(side, "color"), theme, palette) || [0, 0, 0]);
      var w = { thin: "1px solid", hair: "1px dotted", dotted: "1px dotted", dashed: "1px dashed", dashDot: "1px dashed", dashDotDot: "1px dashed",
        medium: "2px solid", mediumDashed: "2px dashed", mediumDashDot: "2px dashed", mediumDashDotDot: "2px dashed", slantDashDot: "2px dashed",
        thick: "3px solid", double: "3px double" }[st] || "1px solid";
      return w + " " + col;
    }
    /* Ancho de columna de Excel (en caracteres) → píxeles */
    function colPx(w, mdw) { return Math.max(0, Math.floor(((256 * w + Math.floor(128 / mdw)) / 256) * mdw)); }

    function readWorkbook(buf) {
      var zip, X;
      return Promise.all([K.Engines.jszip(), K.Engines.xlsx()]).then(function (e) {
        X = e[1];
        return e[0].loadAsync(buf);
      }).then(function (z) {
        zip = z;
        var wbPath = "xl/workbook.xml";
        if (!zip.file(wbPath)) throw K.UserError("No es un libro de Excel válido", "El archivo no contiene hojas de cálculo .xlsx.");
        return Promise.all([xmlPart(zip, wbPath), relsOf(zip, wbPath)]).then(function (r) {
          var wb = r[0], rels = r[1];
          var stylesRel = relByType(rels, "/styles"), themeRel = relByType(rels, "/theme");
          return Promise.all([xmlPart(zip, stylesRel && stylesRel.target), xmlPart(zip, themeRel && themeRel.target)]).then(function (sr) {
            var sheets = descOf(wb.documentElement, "sheet").map(function (s, i) {
              var rel = rels[relId(s, "id")];
              return { name: s.getAttribute("name"), state: s.getAttribute("state") || "visible", path: rel && rel.target, index: i };
            });
            var printAreas = {};
            descOf(wb.documentElement, "definedName").forEach(function (d) {
              if (d.getAttribute("name") === "_xlnm.Print_Area" && d.getAttribute("localSheetId") != null) printAreas[d.getAttribute("localSheetId")] = d.textContent.split(",")[0];
            });
            var date1904 = attrBool(kidOf(wb.documentElement, "workbookPr"), "date1904", false);
            return { zip: zip, X: X, sheets: sheets, styles: sr[0], theme: themeColors(sr[1]), printAreas: printAreas, date1904: date1904 };
          });
        });
      }).then(function (book) {
        // Valores ya formateados como en Excel (fechas, moneda, %), calculados por SheetJS
        book.values = book.X.read(new Uint8Array(buf), { type: "array", cellText: true, cellNF: true, cellDates: false, cellStyles: false, sheetStubs: false });
        return book;
      });
    }

    function compileStyles(styles, theme) {
      var palette = null;
      var ic = styles && descOf(styles.documentElement, "indexedColors")[0];
      if (ic) palette = kidsOf(ic, "rgbColor").map(function (c) { return (c.getAttribute("rgb") || "000000").slice(-6); });
      var root = styles && styles.documentElement;
      var fonts = kidsOf(kidOf(root, "fonts"), "font").map(function (f) {
        var u = kidOf(f, "u");
        return {
          name: (kidOf(f, "name") || {}).getAttribute ? kidOf(f, "name").getAttribute("val") : null,
          scheme: kidOf(f, "scheme") ? kidOf(f, "scheme").getAttribute("val") : null,
          sz: attrNum(kidOf(f, "sz"), "val", 11),
          b: !!kidOf(f, "b") && attrBool(kidOf(f, "b"), "val", true), i: !!kidOf(f, "i") && attrBool(kidOf(f, "i"), "val", true),
          u: !!u && u.getAttribute("val") !== "none", strike: !!kidOf(f, "strike") && attrBool(kidOf(f, "strike"), "val", true),
          color: xlColor(kidOf(f, "color"), theme, palette),
          va: kidOf(f, "vertAlign") ? kidOf(f, "vertAlign").getAttribute("val") : null
        };
      });
      if (!fonts.length) fonts = [{ name: "Calibri", sz: 11 }];
      fonts.forEach(function (f) {
        if (f.scheme === "minor" && theme.$fonts.minor) f.name = theme.$fonts.minor;
        else if (f.scheme === "major" && theme.$fonts.major) f.name = theme.$fonts.major;
        if (!f.name) f.name = "Calibri";
      });
      var fills = kidsOf(kidOf(root, "fills"), "fill").map(function (f) {
        var pf = kidOf(f, "patternFill");
        if (pf) {
          var t = pf.getAttribute("patternType") || (kidOf(pf, "fgColor") ? "solid" : "none");
          if (t === "none") return null;
          var fg = xlColor(kidOf(pf, "fgColor"), theme, palette), bg = xlColor(kidOf(pf, "bgColor"), theme, palette);
          if (t === "solid") return fg || bg;
          if (!fg && !bg) return null;
          // Tramas: se aproximan con el color de primer plano suavizado
          var a = fg || [0, 0, 0], b = bg || [255, 255, 255], k = /gray125|gray0625/.test(t) ? 0.15 : /light/.test(t) ? 0.3 : 0.5;
          return [0, 1, 2].map(function (i) { return Math.round(a[i] * k + b[i] * (1 - k)); });
        }
        var gs = descOf(f, "stop")[0];
        return gs ? xlColor(kidOf(gs, "color"), theme, palette) : null;
      });
      var borders = kidsOf(kidOf(root, "borders"), "border").map(function (b) {
        return { l: xlBorderCss(kidOf(b, "left") || kidOf(b, "start"), theme, palette), r: xlBorderCss(kidOf(b, "right") || kidOf(b, "end"), theme, palette),
                 t: xlBorderCss(kidOf(b, "top"), theme, palette), b: xlBorderCss(kidOf(b, "bottom"), theme, palette) };
      });
      var xfs = kidsOf(kidOf(root, "cellXfs"), "xf").map(function (x) {
        var al = kidOf(x, "alignment");
        return {
          font: fonts[attrNum(x, "fontId", 0)] || fonts[0], fill: fills[attrNum(x, "fillId", 0)] || null,
          border: borders[attrNum(x, "borderId", 0)] || null,
          h: al ? al.getAttribute("horizontal") : null, v: al ? al.getAttribute("vertical") : null,
          wrap: al ? attrBool(al, "wrapText", false) : false, indent: al ? attrNum(al, "indent", 0) : 0,
          rot: al ? attrNum(al, "textRotation", 0) : 0, shrink: al ? attrBool(al, "shrinkToFit", false) : false
        };
      });
      if (!xfs.length) xfs = [{ font: fonts[0] }];
      return { xfs: xfs, fonts: fonts, palette: palette };
    }
    function visibleStyle(xf) { return !!(xf && (xf.fill || (xf.border && (xf.border.l || xf.border.r || xf.border.t || xf.border.b)))); }

    function readSheet(book, S, sheetInfo) {
      var zip = book.zip, path = sheetInfo.path;
      return Promise.all([xmlPart(zip, path), relsOf(zip, path)]).then(function (r) {
        var doc = r[0], rels = r[1];
        if (!doc) return null;
        var root = doc.documentElement;
        var defFont = S.fonts[0] || { name: "Calibri", sz: 11 };
        var mdw = /arial|helvetica|liberation sans/i.test(defFont.name) ? Math.round(defFont.sz * 0.7) : Math.max(5, Math.round(defFont.sz * 0.64));
        var fmt = kidOf(root, "sheetFormatPr");
        var defRowPt = attrNum(fmt, "defaultRowHeight", 15);
        var baseW = attrNum(fmt, "baseColWidth", 8);
        var defColW = attrNum(fmt, "defaultColWidth", 0);
        var defColPx = defColW ? colPx(defColW, mdw) : (Math.floor((baseW * mdw + 5) / 8) * 8 || 64);
        var cols = {}, rows = {}, cells = {}, maxR = -1, maxC = -1;
        kidsOf(kidOf(root, "cols"), "col").forEach(function (c) {
          var a = attrNum(c, "min", 1) - 1, b = attrNum(c, "max", 1) - 1;
          var info = { w: c.getAttribute("width") != null ? colPx(attrNum(c, "width", 0), mdw) : defColPx, hidden: attrBool(c, "hidden", false), s: attrNum(c, "style", 0) };
          for (var i = a; i <= Math.min(b, 16383); i++) cols[i] = info;
          if (b < 16383 && visibleStyle(S.xfs[info.s])) maxC = Math.max(maxC, b);
        });
        var vals = book.values.Sheets[sheetInfo.name] || {};
        kidsOf(kidOf(root, "sheetData"), "row").forEach(function (rw) {
          var ri = attrNum(rw, "r", 0) - 1;
          if (ri < 0) return;
          rows[ri] = { ht: rw.getAttribute("ht") != null ? attrNum(rw, "ht", defRowPt) : null, custom: attrBool(rw, "customHeight", false),
                       hidden: attrBool(rw, "hidden", false), s: attrBool(rw, "customFormat", false) ? attrNum(rw, "s", 0) : null };
          kidsOf(rw, "c").forEach(function (c) {
            var ref = c.getAttribute("r"), pos = ref && cellRef(ref);
            if (!pos) return;
            var v = vals[ref];
            var text = v ? (v.w != null ? v.w : v.v != null ? String(v.v) : "") : "";
            if (v && v.t === "b") text = v.v ? "VERDADERO" : "FALSO";
            if (v && v.t === "e") text = v.w || "#N/A";
            var xf = S.xfs[attrNum(c, "s", 0)] || S.xfs[0];
            cells[pos.r + ":" + pos.c] = { text: text, t: v ? v.t : null, xf: xf };
            if (text !== "" || visibleStyle(xf)) { maxR = Math.max(maxR, pos.r); maxC = Math.max(maxC, pos.c); }
          });
        });
        var merges = descOf(kidOf(root, "mergeCells"), "mergeCell").map(function (m) { return rangeRef(m.getAttribute("ref")); }).filter(Boolean);
        merges.forEach(function (m) { maxR = Math.max(maxR, m.e.r); maxC = Math.max(maxC, m.e.c); });
        var po = kidOf(root, "printOptions"), pm = kidOf(root, "pageMargins"), ps = kidOf(root, "pageSetup"), spr = kidOf(root, "sheetPr", "pageSetUpPr");
        var hf = kidOf(root, "headerFooter");
        var sheet = {
          name: sheetInfo.name, cols: cols, rows: rows, cells: cells, merges: merges, defColPx: defColPx, defRowPx: defRowPt * 96 / 72,
          defFont: defFont, maxR: maxR, maxC: maxC, images: [], charts: 0,
          grid: attrBool(po, "gridLines", false), hCenter: attrBool(po, "horizontalCentered", false), vCenter: attrBool(po, "verticalCentered", false),
          margins: { l: attrNum(pm, "left", 0.7), r: attrNum(pm, "right", 0.7), t: attrNum(pm, "top", 0.75), b: attrNum(pm, "bottom", 0.75), h: attrNum(pm, "header", 0.3), f: attrNum(pm, "footer", 0.3) },
          paper: PAPER[attrNum(ps, "paperSize", 1)] || PAPER[1], landscape: ps ? ps.getAttribute("orientation") === "landscape" : false,
          scale: attrNum(ps, "scale", 100) / 100, fit: attrBool(spr, "fitToPage", false), fitW: attrNum(ps, "fitToWidth", 1), fitH: attrNum(ps, "fitToHeight", 1),
          overThenDown: ps ? ps.getAttribute("pageOrder") === "overThenDown" : false,
          header: hf ? (kidOf(hf, "oddHeader") || {}).textContent || "" : "", footer: hf ? (kidOf(hf, "oddFooter") || {}).textContent || "" : "",
          rowBreaks: descOf(kidOf(root, "rowBreaks"), "brk").map(function (b) { return attrNum(b, "id", 0); }),
          colBreaks: descOf(kidOf(root, "colBreaks"), "brk").map(function (b) { return attrNum(b, "id", 0); })
        };
        var pa = book.printAreas[sheetInfo.index];
        sheet.area = pa ? rangeRef(pa) : (maxR >= 0 && maxC >= 0 ? { s: { r: 0, c: 0 }, e: { r: maxR, c: maxC } } : null);
        // Imágenes ancladas a celdas (los gráficos no se pueden dibujar aquí)
        var dr = relByType(rels, "/drawing");
        if (!dr) return sheet;
        return Promise.all([xmlPart(zip, dr.target), relsOf(zip, dr.target)]).then(function (d) {
          var ddoc = d[0], drels = d[1];
          if (!ddoc) return sheet;
          var jobs = [];
          kidsOf(ddoc.documentElement).forEach(function (anc) {
            var pic = descOf(anc, "pic")[0];
            function pt(el) { return el ? { c: parseInt((kidOf(el, "col") || {}).textContent, 10) || 0, co: (parseInt((kidOf(el, "colOff") || {}).textContent, 10) || 0) / 9525,
              r: parseInt((kidOf(el, "row") || {}).textContent, 10) || 0, ro: (parseInt((kidOf(el, "rowOff") || {}).textContent, 10) || 0) / 9525 } : null; }
            if (!pic) {
              var ch = descOf(anc, "chart")[0];
              if (!ch) return;
              var crel = drels[relId(ch, "id")];
              if (!crel) { sheet.charts++; return; }
              var cext = kidOf(anc, "ext");
              var chart = { chart: true, from: pt(kidOf(anc, "from")), to: pt(kidOf(anc, "to")), ext: cext ? { w: attrNum(cext, "cx", 0) / 9525, h: attrNum(cext, "cy", 0) / 9525 } : null, pos: null };
              jobs.push(xmlPart(zip, crel.target).then(function (cd) { if (cd) { chart.doc = cd; sheet.images.push(chart); } else sheet.charts++; }));
              return;
            }
            var blip = descOf(pic, "blip")[0], rel = blip && drels[relId(blip, "embed")];
            if (!rel || rel.external || /\.(emf|wmf)$/i.test(rel.target)) return;
            var img = { from: pt(kidOf(anc, "from")), to: pt(kidOf(anc, "to")), ext: null, pos: null };
            var ext = kidOf(anc, "ext") || descOf(pic, "ext")[0];
            if (ext) img.ext = { w: attrNum(ext, "cx", 0) / 9525, h: attrNum(ext, "cy", 0) / 9525 };
            var p = kidOf(anc, "pos");
            if (p) img.pos = { x: attrNum(p, "x", 0) / 9525, y: attrNum(p, "y", 0) / 9525 };
            var f = zip.file(rel.target);
            if (!f) return;
            jobs.push(f.async("base64").then(function (b64) {
              var ext2 = rel.target.split(".").pop().toLowerCase();
              img.src = "data:" + (ext2 === "svg" ? "image/svg+xml" : ext2 === "jpg" || ext2 === "jpeg" ? "image/jpeg" : "image/" + ext2) + ";base64," + b64;
              sheet.images.push(img);
            }));
          });
          return Promise.all(jobs).then(function () {
            // Como Excel, el área impresa incluye las imágenes y gráficos de la hoja
            if (!pa) sheet.images.forEach(function (im) {
              var end = im.to || (im.from && im.ext ? { c: im.from.c + Math.ceil(im.ext.w / Math.max(1, colWidth(sheet, im.from.c))), r: im.from.r + Math.ceil(im.ext.h / Math.max(1, rowHeight(sheet, im.from.r))) } : null);
              if (!end) return;
              if (!sheet.area) sheet.area = { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
              sheet.area.e.r = Math.max(sheet.area.e.r, end.r - (im.to && !im.to.ro ? 1 : 0));
              sheet.area.e.c = Math.max(sheet.area.e.c, end.c - (im.to && !im.to.co ? 1 : 0));
            });
            return sheet;
          });
        });
      });
    }

    /* Altura automática de las filas sin alto guardado (texto con ajuste de línea) */
    function autoRowHeights(sheet, host) {
      var need = [];
      Object.keys(sheet.cells).forEach(function (k) {
        var c = sheet.cells[k];
        if (!c.xf || !c.xf.wrap || !c.text) return;
        var rc = k.split(":"), r = +rc[0], col = +rc[1], row = sheet.rows[r];
        if (row && (row.custom || row.ht != null)) return;
        need.push({ r: r, c: col, cell: c });
      });
      if (!need.length) return;
      var probe = document.createElement("div");
      probe.style.cssText = "position:absolute;left:0;top:0;visibility:hidden;white-space:pre-wrap;overflow-wrap:break-word;padding:1px 3px;box-sizing:border-box;";
      host.appendChild(probe);
      need.forEach(function (n) {
        var f = n.cell.xf.font || sheet.defFont;
        probe.style.width = colWidth(sheet, n.c) + "px";
        probe.style.font = (f.i ? "italic " : "") + (f.b ? "700 " : "400 ") + f.sz + "pt " + fontStack('"' + f.name + '"');
        probe.style.lineHeight = "normal";
        probe.textContent = n.cell.text;
        var h = probe.offsetHeight + 2;
        var row = sheet.rows[n.r] || (sheet.rows[n.r] = { ht: null });
        row.autoPx = Math.max(row.autoPx || sheet.defRowPx, h);
      });
      probe.remove();
    }
    function colWidth(sheet, c) { var ci = sheet.cols[c]; return ci ? (ci.hidden ? 0 : ci.w) : sheet.defColPx; }
    function rowHeight(sheet, r) {
      var ri = sheet.rows[r];
      if (!ri) return sheet.defRowPx;
      if (ri.hidden) return 0;
      if (ri.ht != null) return ri.ht * 96 / 72;
      return ri.autoPx || sheet.defRowPx;
    }

    /* Encabezados y pies de página de Excel (&P, &N, &A, &D…) */
    function xlHeaderParts(code, ctx) {
      if (!code) return null;
      var parts = { L: "", C: "", R: "" }, cur = "C";
      var s = String(code).replace(/&"[^"]*"/g, "").replace(/&\d+/g, "").replace(/&K[0-9A-Fa-f]{6}/g, "").replace(/&[BIUSXYEOH]/g, "");
      for (var i = 0; i < s.length; i++) {
        var ch = s.charAt(i);
        if (ch === "&" && i + 1 < s.length) {
          var n = s.charAt(i + 1);
          i++;
          if (n === "L" || n === "C" || n === "R") { cur = n; continue; }
          if (n === "P") { parts[cur] += String(ctx.page); continue; }
          if (n === "N") { parts[cur] += String(ctx.total); continue; }
          if (n === "A") { parts[cur] += ctx.sheet; continue; }
          if (n === "F") { parts[cur] += ctx.file; continue; }
          if (n === "D") { parts[cur] += new Date().toLocaleDateString("es"); continue; }
          if (n === "T") { parts[cur] += new Date().toLocaleTimeString("es", { hour: "2-digit", minute: "2-digit" }); continue; }
          if (n === "Z") { continue; }
          if (n === "&") { parts[cur] += "&"; continue; }
          continue;
        }
        parts[cur] += ch;
      }
      return parts.L || parts.C || parts.R ? parts : null;
    }

    /* Reparte la hoja en páginas (filas y columnas) como lo hace Excel al imprimir */
    function planSheet(sheet) {
      if (!sheet.area) return [];
      var A = sheet.area, landscape = sheet.landscape;
      var pw = (landscape ? sheet.paper[1] : sheet.paper[0]) * 96, ph = (landscape ? sheet.paper[0] : sheet.paper[1]) * 96;
      var M = sheet.margins, printW = pw - (M.l + M.r) * 96, printH = ph - (M.t + M.b) * 96;
      var colsW = [], rowsH = [], totalW = 0, totalH = 0, c, r;
      for (c = A.s.c; c <= A.e.c; c++) { colsW.push(colWidth(sheet, c)); totalW += colsW[colsW.length - 1]; }
      for (r = A.s.r; r <= A.e.r; r++) { rowsH.push(rowHeight(sheet, r)); totalH += rowsH[rowsH.length - 1]; }
      var scale = sheet.scale || 1;
      if (sheet.fit) {
        var sw = sheet.fitW > 0 ? (sheet.fitW * printW) / Math.max(1, totalW) : Infinity;
        var sh = sheet.fitH > 0 ? (sheet.fitH * printH) / Math.max(1, totalH) : Infinity;
        scale = Math.min(1, sw, sh);
      }
      scale = Math.max(0.1, Math.min(4, scale));
      function groups(sizes, limit, start, breaks) {
        var out = [], from = 0, acc = 0;
        for (var i = 0; i < sizes.length; i++) {
          var forced = i > from && breaks.indexOf(start + i) >= 0;
          if ((acc + sizes[i] * scale > limit + 0.5 && i > from) || forced) { out.push([from, i - 1]); from = i; acc = 0; }
          acc += sizes[i] * scale;
        }
        out.push([from, sizes.length - 1]);
        return out;
      }
      // Los saltos manuales de Excel indican la última fila/columna antes del corte (base 0 = id)
      var rg = groups(rowsH, printH, A.s.r, sheet.rowBreaks.map(function (b) { return b; }));
      var cg = groups(colsW, printW, A.s.c, sheet.colBreaks.map(function (b) { return b; }));
      var pages = [];
      if (sheet.overThenDown) rg.forEach(function (ry) { cg.forEach(function (cx) { pages.push({ rows: ry, cols: cx }); }); });
      else cg.forEach(function (cx) { rg.forEach(function (ry) { pages.push({ rows: ry, cols: cx }); }); });
      return pages.map(function (p) {
        return { sheet: sheet, r0: A.s.r + p.rows[0], r1: A.s.r + p.rows[1], c0: A.s.c + p.cols[0], c1: A.s.c + p.cols[1], pw: pw, ph: ph, scale: scale };
      });
    }

    /* HTML de una página de la hoja */
    function sheetPageHtml(pg, ctx) {
      var S = pg.sheet, M = S.margins, s = pg.scale, r, c;
      var colW = [], rowH = [], gw = 0, gh = 0;
      for (c = pg.c0; c <= pg.c1; c++) { colW.push(colWidth(S, c)); gw += colW[colW.length - 1]; }
      for (r = pg.r0; r <= pg.r1; r++) { rowH.push(rowHeight(S, r)); gh += rowH[rowH.length - 1]; }
      // Celdas cubiertas por combinaciones y celdas de inicio de combinación
      var covered = {}, spans = {};
      S.merges.forEach(function (m) {
        if (m.e.r < pg.r0 || m.s.r > pg.r1 || m.e.c < pg.c0 || m.s.c > pg.c1) return;
        var r0 = Math.max(m.s.r, pg.r0), c0 = Math.max(m.s.c, pg.c0), r1 = Math.min(m.e.r, pg.r1), c1 = Math.min(m.e.c, pg.c1);
        for (var rr = r0; rr <= r1; rr++) for (var cc = c0; cc <= c1; cc++) covered[rr + ":" + cc] = 1;
        delete covered[r0 + ":" + c0];
        spans[r0 + ":" + c0] = { rs: r1 - r0 + 1, cs: c1 - c0 + 1, src: m.s.r + ":" + m.s.c };
      });
      function isEmpty(rr, cc) { var x = S.cells[rr + ":" + cc]; return !x || x.text === "" ; }
      var grid = S.grid ? "1px solid #9a9a9a" : "";
      var html = '<table style="border-collapse:collapse;table-layout:fixed;width:' + gw + "px;font-family:" + escH(fontStack('"' + S.defFont.name + '"')) + ";font-size:" + S.defFont.sz + 'pt">';
      html += "<colgroup>" + colW.map(function (w) { return '<col style="width:' + w + 'px">'; }).join("") + "</colgroup>";
      for (r = pg.r0; r <= pg.r1; r++) {
        var h = rowH[r - pg.r0];
        if (!h) continue;
        html += '<tr style="height:' + h + 'px">';
        for (c = pg.c0; c <= pg.c1; c++) {
          var key = r + ":" + c;
          if (covered[key]) continue;
          var w = colW[c - pg.c0];
          var sp = spans[key];
          var cell = S.cells[sp ? sp.src : key];
          var xf = cell ? cell.xf : (S.rows[r] && S.rows[r].s != null ? S.xfsRef[S.rows[r].s] : (S.cols[c] ? S.xfsRef[S.cols[c].s] : null));
          if (!w && !sp) { html += '<td style="width:0;padding:0;border:0"></td>'; continue; }
          var st = "padding:0;position:relative;overflow:visible;";
          if (xf && xf.fill) st += "background:" + css(xf.fill) + ";";
          var bd = xf && xf.border;
          if (bd && sp) {
            // En celdas combinadas, el borde derecho/inferior viene de la última celda
            var last = S.cells[(r + sp.rs - 1) + ":" + (c + sp.cs - 1)];
            if (last && last.xf && last.xf.border) bd = { l: bd.l, t: bd.t, r: last.xf.border.r || bd.r, b: last.xf.border.b || bd.b };
          }
          st += "border-left:" + ((bd && bd.l) || grid || "0") + ";border-right:" + ((bd && bd.r) || grid || "0") + ";border-top:" + ((bd && bd.t) || grid || "0") + ";border-bottom:" + ((bd && bd.b) || grid || "0") + ";";
          var attrs = sp ? (sp.rs > 1 ? ' rowspan="' + sp.rs + '"' : "") + (sp.cs > 1 ? ' colspan="' + sp.cs + '"' : "") : "";
          var inner = "";
          if (cell && cell.text !== "") {
            var f = (xf && xf.font) || S.defFont;
            var num = cell.t === "n" || cell.t === "d";
            var ha = xf && xf.h && xf.h !== "general" ? xf.h : (num ? "right" : cell.t === "b" || cell.t === "e" ? "center" : "left");
            if (ha === "centerContinuous" || ha === "distributed") ha = "center";
            if (ha === "fill") ha = "left";
            var va = xf && xf.v ? xf.v : "bottom";
            var wrap = xf && xf.wrap;
            // Ancho visible: la celda (o combinación) y, si el texto no se ajusta, las vecinas vacías
            var cw = 0, k2;
            for (k2 = 0; k2 < (sp ? sp.cs : 1); k2++) cw += colW[c - pg.c0 + k2] || 0;
            var left = 0, right = 0;
            if (!wrap && !num && !sp) {
              if (ha === "left" || ha === "center") for (k2 = c + 1; k2 <= pg.c1 && isEmpty(r, k2) && !covered[r + ":" + k2] && !spans[r + ":" + k2]; k2++) right += colW[k2 - pg.c0];
              if (ha === "right" || ha === "center") for (k2 = c - 1; k2 >= pg.c0 && isEmpty(r, k2) && !covered[r + ":" + k2] && !spans[r + ":" + k2]; k2--) left += colW[k2 - pg.c0];
              if (ha === "center") { var m2 = Math.min(left, right); left = right = m2; }
            }
            var fst = "font-family:" + fontStack('"' + f.name + '"') + ";font-size:" + f.sz + "pt;" + (f.b ? "font-weight:700;" : "") + (f.i ? "font-style:italic;" : "") +
              (f.color ? "color:" + css(f.color) + ";" : "color:#000;") +
              ((f.u || f.strike) ? "text-decoration:" + (f.u ? "underline " : "") + (f.strike ? "line-through" : "") + ";" : "");
            var ind = xf && xf.indent ? xf.indent * 9 : 0;
            var just = va === "top" ? "flex-start" : va === "center" ? "center" : "flex-end";
            inner = '<div style="position:absolute;top:0;bottom:0;left:' + (-left) + "px;right:" + (-right) + "px;overflow:hidden;display:flex;align-items:" + just + ";justify-content:" +
              (ha === "right" ? "flex-end" : ha === "center" ? "center" : "flex-start") + ";padding:0 " + (2 + (ha === "right" ? ind : 0)) + "px 0 " + (2 + (ha !== "right" ? ind : 0)) + 'px;box-sizing:border-box;">' +
              '<div style="' + escH(fst) + "line-height:1.2;" + (wrap ? "white-space:pre-wrap;overflow-wrap:break-word;width:100%;text-align:" + ha + ";" : "white-space:pre;") + '">' + escH(cell.text) + "</div></div>";
          }
          html += "<td" + attrs + ' style="' + st + '">' + inner + "</td>";
        }
        html += "</tr>";
      }
      html += "</table>";
      // Imágenes de la hoja que caen en esta página
      var imgs = "";
      function xAt(col, off) { var x = 0; for (var cc = pg.c0; cc < col; cc++) x += colWidth(S, cc); for (cc = col; cc < pg.c0; cc++) x -= colWidth(S, cc); return x + off; }
      function yAt(row, off) { var y = 0; for (var rr = pg.r0; rr < row; rr++) y += rowHeight(S, rr); for (rr = row; rr < pg.r0; rr++) y -= rowHeight(S, rr); return y + off; }
      S.images.forEach(function (im) {
        var x, y, w, h;
        if (im.from) { x = xAt(im.from.c, im.from.co); y = yAt(im.from.r, im.from.ro); }
        else if (im.pos) { x = xAt(0, im.pos.x); y = yAt(0, im.pos.y); }
        else return;
        if (im.to) { w = xAt(im.to.c, im.to.co) - x; h = yAt(im.to.r, im.to.ro) - y; }
        else if (im.ext) { w = im.ext.w; h = im.ext.h; }
        else return;
        if (x > gw || y > gh || x + w < 0 || y + h < 0) return;
        if (im.chart) { imgs += '<div style="position:absolute;left:0;top:0;width:0;height:0">' + chartSvg(im.doc, { x: x, y: y, w: w, h: h }, S.colorOf, S.theme) + "</div>"; return; }
        imgs += '<img src="' + im.src + '" style="position:absolute;left:' + x + "px;top:" + y + "px;width:" + w + "px;height:" + h + 'px" alt="">';
      });
      var ox = M.l * 96, oy = M.t * 96;
      var areaW = pg.pw - (M.l + M.r) * 96, areaH = pg.ph - (M.t + M.b) * 96;
      if (S.hCenter) ox += Math.max(0, (areaW - gw * s) / 2);
      if (S.vCenter) oy += Math.max(0, (areaH - gh * s) / 2);
      var out = '<div style="position:absolute;left:' + ox + "px;top:" + oy + "px;width:" + gw + "px;height:" + gh + "px;transform:scale(" + s + ");transform-origin:0 0;overflow:hidden;" + '">' + html + imgs + "</div>";
      function band(code, top) {
        var p = xlHeaderParts(code, ctx);
        if (!p) return "";
        var pos = top ? "top:" + M.h * 96 + "px;" : "bottom:" + M.f * 96 + "px;";
        var fs = "font-family:" + escH(fontStack('"' + S.defFont.name + '"')) + ";font-size:" + Math.min(S.defFont.sz, 11) + "pt;color:#000;white-space:pre;";
        return '<div style="position:absolute;left:' + M.l * 96 + "px;right:" + M.r * 96 + "px;" + pos + 'display:flex;justify-content:space-between;">' +
          '<div style="flex:1;text-align:left;' + fs + '">' + escH(p.L) + '</div><div style="flex:1;text-align:center;' + fs + '">' + escH(p.C) + '</div><div style="flex:1;text-align:right;' + fs + '">' + escH(p.R) + "</div></div>";
      }
      return out + band(S.header, true) + band(S.footer, false);
    }

    function xlsxPages(buf, onProgress, fileName) {
      ensureFontFaces();
      var host = makeHost();
      return readWorkbook(buf).then(function (book) {
        if (onProgress) onProgress(0.1);
        var S = compileStyles(book.styles, book.theme);
        var list = book.sheets.filter(function (s) { return s.state === "visible" && s.path; });
        var sheets = [], i = 0, notes = { charts: 0, truncated: 0 };
        function next() {
          if (i >= list.length) return Promise.resolve();
          return readSheet(book, S, list[i++]).then(function (sh) {
            if (sh) {
              sh.xfsRef = S.xfs; sh.theme = book.theme; sh.colorOf = makeColorOf(book.theme, { tx1: "dk1", bg1: "lt1", tx2: "dk2", bg2: "lt2" });
              if (sh.area && sh.area.e.r - sh.area.s.r + 1 > K.BRAND.limits.maxSheetRows) { sh.area.e.r = sh.area.s.r + K.BRAND.limits.maxSheetRows - 1; notes.truncated++; }
              notes.charts += sh.charts;
              sheets.push(sh);
            }
            if (onProgress) onProgress(0.1 + 0.1 * i / list.length);
          }).then(K.yieldUI).then(next);
        }
        return next().then(function () {
          sheets.forEach(function (sh) { autoRowHeights(sh, host); });
          var plan = [];
          sheets.forEach(function (sh) { plan = plan.concat(planSheet(sh)); });
          if (!plan.length) throw K.UserError("El libro está vacío", "No hay celdas con datos para convertir a PDF.");
          var total = plan.length, sample = "";
          sheets.forEach(function (sh) { Object.keys(sh.cells).forEach(function (k) { if (sample.length < 200000) sample += sh.cells[k].text + " "; }); });
          var pages = plan.map(function (pg, idx) {
            return {
              sample: idx === 0 ? sample : "",
              render: function (opt) {
                var div = document.createElement("div");
                div.style.cssText = "position:relative;overflow:hidden;background:#fff;width:" + pg.pw + "px;height:" + pg.ph + "px;";
                div.innerHTML = sheetPageHtml(pg, { page: idx + 1, total: total, sheet: pg.sheet.name, file: fileName || "" });
                host.appendChild(div);
                return settle(div).then(function () { return rasterize(div, "#ffffff"); }).then(function (canvas) {
                  var texts = opt && opt.noText ? null : textLayer(div);
                  div.remove();
                  return { canvas: canvas, texts: texts, w: pg.pw, h: pg.ph };
                });
              }
            };
          });
          pages.cleanup = function () { host.remove(); };
          pages.notes = notes; pages.sheets = sheets.length;
          return pages;
        });
      }).catch(function (e) { host.remove(); throw e; });
    }
    O.xlsxToPdf = function (buf, onProgress, fileName) {
      return xlsxPages(buf, onProgress, fileName).then(function (pages) {
        return pagesToPdf(pages, onProgress, 0.25, 1).then(function (blob) { pages.cleanup(); return { blob: blob, pages: pages.length, sheets: pages.sheets, notes: pages.notes }; },
          function (e) { pages.cleanup(); throw e; });
      });
    };
    O.xlsxToImages = function (buf, onProgress, type, quality, fileName) {
      return xlsxPages(buf, onProgress, fileName).then(function (pages) {
        return pagesToImages(pages, onProgress, 0.25, 1, type, quality).then(function (r) { pages.cleanup(); return r; }, function (e) { pages.cleanup(); throw e; });
      });
    };

    /* ---------------------------------------------------------------
       7. POWERPOINT (.pptx) → diapositivas, con formas, tablas y textos reales
       --------------------------------------------------------------- */
    var EMU = 9525;   // EMU por píxel CSS
    function px(v) { return (v || 0) / EMU; }
    function fmt(n) { return Math.round(n * 100) / 100; }

    /* Color DrawingML (srgb/tema/sistema) con sus modificadores; devuelve [r,g,b,alpha] */
    function makeColorOf(theme, clrMap) {
      return function colorOf(holder, phClr) {
        if (!holder) return null;
        var c = holder.localName && /^(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(holder.localName) ? holder : (kidOf(holder, "srgbClr") || kidOf(holder, "schemeClr") || kidOf(holder, "sysClr") || kidOf(holder, "prstClr") || kidOf(holder, "scrgbClr") || kidOf(holder, "hslClr"));
        if (!c) return null;
        var rgb = null, n = c.localName;
        if (n === "srgbClr") rgb = hex2rgb(c.getAttribute("val"));
        else if (n === "sysClr") rgb = hex2rgb(c.getAttribute("lastClr") || (c.getAttribute("val") === "window" ? "FFFFFF" : "000000"));
        else if (n === "scrgbClr") rgb = [attrNum(c, "r", 0), attrNum(c, "g", 0), attrNum(c, "b", 0)].map(function (v) { return Math.round(Math.min(1, v / 100000) * 255); });
        else if (n === "prstClr") rgb = ({ black: [0, 0, 0], white: [255, 255, 255], red: [255, 0, 0], blue: [0, 0, 255], green: [0, 128, 0], yellow: [255, 255, 0], gray: [128, 128, 128], grey: [128, 128, 128], orange: [255, 165, 0] })[c.getAttribute("val")] || [0, 0, 0];
        else if (n === "schemeClr") {
          var v = c.getAttribute("val");
          if (v === "phClr") rgb = phClr ? phClr.slice(0, 3) : null;
          else { var mapped = clrMap[v] || v; rgb = theme[mapped] || theme[v] || null; }
        }
        if (!rgb) return null;
        var a = 1;
        kidsOf(c).forEach(function (m) {
          var val = attrNum(m, "val", 0) / 100000, hsl;
          switch (m.localName) {
            case "lumMod": hsl = rgb2hsl(rgb); hsl[2] = Math.min(1, hsl[2] * val); rgb = hsl2rgb(hsl); break;
            case "lumOff": hsl = rgb2hsl(rgb); hsl[2] = Math.max(0, Math.min(1, hsl[2] + val)); rgb = hsl2rgb(hsl); break;
            case "satMod": hsl = rgb2hsl(rgb); hsl[1] = Math.max(0, Math.min(1, hsl[1] * val)); rgb = hsl2rgb(hsl); break;
            case "tint": rgb = rgb.map(function (x) { return Math.round(x * val + 255 * (1 - val)); }); break;
            case "shade": rgb = rgb.map(function (x) { return Math.round(x * val); }); break;
            case "alpha": a = val; break;
            case "alphaMod": a = a * val; break;
          }
        });
        return [rgb[0], rgb[1], rgb[2], a];
      };
    }
    function mix(a, b, k) { return [0, 1, 2].map(function (i) { return Math.round(a[i] * k + b[i] * (1 - k)); }).concat([1]); }
    function pattDensity(p) { var m = /pct(\d+)/.exec(p); if (m) return +m[1] / 100; return /^lt|^nar|^sm/.test(p) ? 0.25 : /^dk/.test(p) ? 0.6 : 0.4; }
    function rgba(c) { return c ? (c[3] != null && c[3] < 1 ? "rgba(" + c[0] + "," + c[1] + "," + c[2] + "," + fmt(c[3]) + ")" : "rgb(" + c[0] + "," + c[1] + "," + c[2] + ")") : "none"; }

    /* Formas predefinidas más usadas → trazado SVG (w,h en px; adj en 1/100000) */
    function presetPath(prst, w, h, av) {
      function adj(name, def) { return av[name] != null ? av[name] : def; }
      var ss = Math.min(w, h), P = [];
      function poly(pts) { return "M" + pts.map(function (p) { return fmt(p[0]) + "," + fmt(p[1]); }).join(" L") + " Z"; }
      switch (prst) {
        case "rect": case "flowChartProcess": case "flowChartPredefinedProcess": case "flowChartInternalStorage": return null;
        case "roundRect": case "flowChartAlternateProcess": {
          var r = ss * Math.min(0.5, adj("adj", 16667) / 100000);
          return "M" + r + ",0 H" + (w - r) + " A" + r + "," + r + " 0 0 1 " + w + "," + r + " V" + (h - r) + " A" + r + "," + r + " 0 0 1 " + (w - r) + "," + h + " H" + r + " A" + r + "," + r + " 0 0 1 0," + (h - r) + " V" + r + " A" + r + "," + r + " 0 0 1 " + r + ",0 Z";
        }
        case "ellipse": case "flowChartConnector": case "donut": case "pie": case "chord": case "blockArc":
          return "M0," + h / 2 + " A" + w / 2 + "," + h / 2 + " 0 1 1 " + w + "," + h / 2 + " A" + w / 2 + "," + h / 2 + " 0 1 1 0," + h / 2 + " Z";
        case "flowChartTerminator": { var rr = Math.min(h / 2, w / 2); return "M" + rr + ",0 H" + (w - rr) + " A" + rr + "," + (h / 2) + " 0 0 1 " + (w - rr) + "," + h + " H" + rr + " A" + rr + "," + (h / 2) + " 0 0 1 " + rr + ",0 Z"; }
        case "triangle": case "flowChartExtract": return poly([[w * adj("adj", 50000) / 100000, 0], [w, h], [0, h]]);
        case "flowChartMerge": return poly([[0, 0], [w, 0], [w / 2, h]]);
        case "rtTriangle": return poly([[0, 0], [w, h], [0, h]]);
        case "diamond": case "flowChartDecision": return poly([[w / 2, 0], [w, h / 2], [w / 2, h], [0, h / 2]]);
        case "parallelogram": case "flowChartInputOutput": { var o = ss * adj("adj", 25000) / 100000; return poly([[o, 0], [w, 0], [w - o, h], [0, h]]); }
        case "trapezoid": { var t = ss * adj("adj", 25000) / 100000; return poly([[t, 0], [w - t, 0], [w, h], [0, h]]); }
        case "pentagon": return poly([[w / 2, 0], [w, h * 0.382], [w * 0.809, h], [w * 0.191, h], [0, h * 0.382]]);
        case "hexagon": { var hx = ss * adj("adj", 25000) / 100000; return poly([[hx, 0], [w - hx, 0], [w, h / 2], [w - hx, h], [hx, h], [0, h / 2]]); }
        case "octagon": { var ox = ss * adj("adj", 29289) / 100000; return poly([[ox, 0], [w - ox, 0], [w, ox], [w, h - ox], [w - ox, h], [ox, h], [0, h - ox], [0, ox]]); }
        case "homePlate": { var hp = ss * adj("adj", 50000) / 100000; return poly([[0, 0], [w - hp, 0], [w, h / 2], [w - hp, h], [0, h]]); }
        case "chevron": { var cv = ss * adj("adj", 50000) / 100000; return poly([[0, 0], [w - cv, 0], [w, h / 2], [w - cv, h], [0, h], [cv, h / 2]]); }
        case "rightArrow": case "leftArrow": case "upArrow": case "downArrow": {
          var a1 = adj("adj1", 50000) / 100000, a2 = adj("adj2", 50000) / 100000;
          if (prst === "rightArrow" || prst === "leftArrow") {
            var th = h * a1, hl = Math.min(w, ss * a2), y0 = (h - th) / 2;
            P = [[0, y0], [w - hl, y0], [w - hl, 0], [w, h / 2], [w - hl, h], [w - hl, h - y0], [0, h - y0]];
            if (prst === "leftArrow") P = P.map(function (p) { return [w - p[0], p[1]]; });
          } else {
            var tw = w * a1, hh = Math.min(h, ss * a2), x0 = (w - tw) / 2;
            P = [[x0, h], [x0, hh], [0, hh], [w / 2, 0], [w, hh], [w - x0, hh], [w - x0, h]];
            if (prst === "downArrow") P = P.map(function (p) { return [p[0], h - p[1]]; });
          }
          return poly(P);
        }
        case "leftRightArrow": { var lt = h * adj("adj1", 50000) / 100000, lh = Math.min(w / 2, ss * adj("adj2", 50000) / 100000), ly = (h - lt) / 2;
          return poly([[0, h / 2], [lh, 0], [lh, ly], [w - lh, ly], [w - lh, 0], [w, h / 2], [w - lh, h], [w - lh, h - ly], [lh, h - ly], [lh, h]]); }
        case "plus": case "mathPlus": { var pl = ss * adj("adj", 25000) / 100000; return poly([[pl, 0], [w - pl, 0], [w - pl, pl], [w, pl], [w, h - pl], [w - pl, h - pl], [w - pl, h], [pl, h], [pl, h - pl], [0, h - pl], [0, pl], [pl, pl]]); }
        case "star5": {
          for (var i = 0; i < 10; i++) { var ang = -Math.PI / 2 + i * Math.PI / 5, rad = i % 2 ? 0.38 : 1; P.push([w / 2 + Math.cos(ang) * w / 2 * rad, h / 2 + 0.05 * h + Math.sin(ang) * h / 2 * rad]); }
          return poly(P);
        }
        case "snip1Rect": { var sn = ss * adj("adj", 16667) / 100000; return poly([[0, 0], [w - sn, 0], [w, sn], [w, h], [0, h]]); }
        case "snip2SameRect": { var s2 = ss * adj("adj1", 16667) / 100000; return poly([[s2, 0], [w - s2, 0], [w, s2], [w, h], [0, h], [0, s2]]); }
        case "round1Rect": { var r1 = ss * adj("adj", 16667) / 100000; return "M0,0 H" + (w - r1) + " A" + r1 + "," + r1 + " 0 0 1 " + w + "," + r1 + " V" + h + " H0 Z"; }
        case "round2SameRect": { var r2 = ss * adj("adj1", 16667) / 100000; return "M" + r2 + ",0 H" + (w - r2) + " A" + r2 + "," + r2 + " 0 0 1 " + w + "," + r2 + " V" + h + " H0 V" + r2 + " A" + r2 + "," + r2 + " 0 0 1 " + r2 + ",0 Z"; }
        case "wedgeRectCallout": case "wedgeRoundRectCallout": case "wedgeEllipseCallout": {
          var px1 = w * (0.5 + adj("adj1", -20833) / 100000), py1 = h * (0.5 + adj("adj2", 62500) / 100000);
          return "M0,0 H" + w + " V" + h + " H" + (w * 0.583) + " L" + fmt(px1) + "," + fmt(py1) + " L" + (w * 0.417) + "," + h + " H0 Z";
        }
        case "flowChartDocument": return "M0,0 H" + w + " V" + fmt(h * 0.802) + " C" + fmt(w * 0.5) + "," + fmt(h * 0.802) + " " + fmt(w * 0.5) + "," + fmt(h * 1.1075) + " 0," + fmt(h * 0.934) + " Z";
        case "wave": case "doubleWave": { var wa = h * adj("adj1", 12500) / 100000; return "M0," + fmt(wa) + " C" + fmt(w / 3) + "," + fmt(-wa) + " " + fmt(w * 2 / 3) + "," + fmt(wa * 3) + " " + w + "," + fmt(wa) + " V" + fmt(h - wa) + " C" + fmt(w * 2 / 3) + "," + fmt(h + wa) + " " + fmt(w / 3) + "," + fmt(h - wa * 3) + " 0," + fmt(h - wa) + " Z"; }
        case "flowChartPunchedTape": { var pa = h * 0.1; return "M0," + fmt(pa) + " C" + fmt(w / 4) + "," + fmt(-pa) + " " + fmt(w / 4) + "," + fmt(pa * 3) + " " + fmt(w / 2) + "," + fmt(pa) + " S" + fmt(w * 3 / 4) + "," + fmt(-pa) + " " + w + "," + fmt(pa) + " V" + fmt(h - pa) + " C" + fmt(w * 3 / 4) + "," + fmt(h + pa) + " " + fmt(w * 3 / 4) + "," + fmt(h - pa * 3) + " " + fmt(w / 2) + "," + fmt(h - pa) + " S" + fmt(w / 4) + "," + fmt(h + pa) + " 0," + fmt(h - pa) + " Z"; }
        case "line": case "straightConnector1": case "bentConnector2": case "bentConnector3": case "curvedConnector3": case "bentConnector4":
          return "LINE";
        case "can": case "flowChartMagneticDisk": { var cy = h * Math.min(0.25, adj("adj", 25000) / 100000 / 2) ; return "M0," + cy + " A" + w / 2 + "," + cy + " 0 0 1 " + w + "," + cy + " V" + (h - cy) + " A" + w / 2 + "," + cy + " 0 0 1 0," + (h - cy) + " Z M0," + cy + " A" + w / 2 + "," + cy + " 0 0 0 " + w + "," + cy; }
        case "frame": { var fr = ss * adj("adj1", 12500) / 100000; return "M0,0 H" + w + " V" + h + " H0 Z M" + fr + "," + fr + " V" + (h - fr) + " H" + (w - fr) + " V" + fr + " Z"; }
        case "heart": return "M" + w / 2 + "," + h * 0.25 + " C" + w / 2 + ",0 0,0 0," + h * 0.3 + " C0," + h * 0.6 + " " + w / 2 + "," + h * 0.8 + " " + w / 2 + "," + h + " C" + w / 2 + "," + h * 0.8 + " " + w + "," + h * 0.6 + " " + w + "," + h * 0.3 + " C" + w + ",0 " + w / 2 + ",0 " + w / 2 + "," + h * 0.25 + " Z";
        case "cloud": case "cloudCallout": case "flowChartOnlineStorage": case "teardrop": case "sun": case "moon": case "smileyFace": case "noSmoking":
          return "M0," + h / 2 + " A" + w / 2 + "," + h / 2 + " 0 1 1 " + w + "," + h / 2 + " A" + w / 2 + "," + h / 2 + " 0 1 1 0," + h / 2 + " Z";
        default: return null;
      }
    }
    /* Geometría libre (custGeom) → trazado SVG */
    function customPath(geom, w, h) {
      var out = "";
      descOf(geom, "path").forEach(function (p) {
        var pw = attrNum(p, "w", 0) || w * EMU, ph = attrNum(p, "h", 0) || h * EMU, sx = w / (pw || 1), sy = h / (ph || 1), cx = 0, cy = 0;
        function pt(el) { return [attrNum(el, "x", 0) * sx, attrNum(el, "y", 0) * sy]; }
        kidsOf(p).forEach(function (c) {
          var pts = kidsOf(c, "pt").map(pt);
          switch (c.localName) {
            case "moveTo": cx = pts[0][0]; cy = pts[0][1]; out += "M" + fmt(cx) + "," + fmt(cy) + " "; break;
            case "lnTo": cx = pts[0][0]; cy = pts[0][1]; out += "L" + fmt(cx) + "," + fmt(cy) + " "; break;
            case "cubicBezTo": out += "C" + pts.map(function (q) { return fmt(q[0]) + "," + fmt(q[1]); }).join(" ") + " "; cx = pts[2][0]; cy = pts[2][1]; break;
            case "quadBezTo": out += "Q" + pts.map(function (q) { return fmt(q[0]) + "," + fmt(q[1]); }).join(" ") + " "; cx = pts[1][0]; cy = pts[1][1]; break;
            case "arcTo": {
              var wr = attrNum(c, "wR", 0) * sx, hr = attrNum(c, "hR", 0) * sy, st = attrNum(c, "stAng", 0) / 60000 * Math.PI / 180, sw = attrNum(c, "swAng", 0) / 60000 * Math.PI / 180;
              var ox = cx - wr * Math.cos(st), oy = cy - hr * Math.sin(st), ex = ox + wr * Math.cos(st + sw), ey = oy + hr * Math.sin(st + sw);
              out += "A" + fmt(wr) + "," + fmt(hr) + " 0 " + (Math.abs(sw) > Math.PI ? 1 : 0) + " " + (sw > 0 ? 1 : 0) + " " + fmt(ex) + "," + fmt(ey) + " ";
              cx = ex; cy = ey; break;
            }
            case "close": out += "Z "; break;
          }
        });
      });
      return out || null;
    }

    function pptxPages(buf, onProgress) {
      ensureFontFaces();
      var host = makeHost(), zip, notes = { charts: 0, skippedImages: 0 };
      var mediaCache = {};
      function media(path) {
        if (!mediaCache[path]) {
          var f = zip.file(path);
          if (!f || /\.(emf|wmf|tiff?)$/i.test(path)) { notes.skippedImages++; mediaCache[path] = Promise.resolve(null); }
          else mediaCache[path] = f.async("base64").then(function (b64) {
            var ext = (path.split(".").pop() || "").toLowerCase();
            return "data:" + (ext === "svg" ? "image/svg+xml" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : "image/" + ext) + ";base64," + b64;
          });
        }
        return mediaCache[path];
      }
      return K.Engines.jszip().then(function (JSZip) { return JSZip.loadAsync(buf); }).then(function (z) {
        zip = z;
        var presPath = "ppt/presentation.xml";
        if (!zip.file(presPath)) throw K.UserError("No es una presentación válida", "El archivo no contiene diapositivas de PowerPoint.");
        return Promise.all([xmlPart(zip, presPath), relsOf(zip, presPath), xmlPart(zip, "ppt/tableStyles.xml")]);
      }).then(function (r) {
        var pres = r[0], presRels = r[1], tblStyles = r[2];
        var sz = kidOf(pres.documentElement, "sldSz");
        var W = px(attrNum(sz, "cx", 12192000)), H = px(attrNum(sz, "cy", 6858000));
        var defText = kidOf(pres.documentElement, "defaultTextStyle");
        var ids = descOf(pres.documentElement, "sldId").map(function (s) { var rel = presRels[relId(s, "id")]; return rel && rel.target; }).filter(Boolean);
        var layerCache = {}, themeCache = {};
        function layer(path) {
          if (!layerCache[path]) layerCache[path] = Promise.all([xmlPart(zip, path), relsOf(zip, path)]).then(function (x) { return { doc: x[0], rels: x[1], path: path }; });
          return layerCache[path];
        }
        function themeOf(path) {
          if (!themeCache[path]) themeCache[path] = xmlPart(zip, path).then(function (doc) {
            var t = themeColors(doc);
            var fmtS = doc && descOf(doc.documentElement, "fmtScheme")[0];
            t.$ln = kidsOf(kidOf(fmtS, "lnStyleLst"), "ln").map(function (l) { return attrNum(l, "w", 9525); });
            t.$eff = kidsOf(kidOf(fmtS, "effectStyleLst"), "effectStyle");
            t.$fill = kidsOf(kidOf(fmtS, "fillStyleLst"));
            t.$bgFill = kidsOf(kidOf(fmtS, "bgFillStyleLst"));
            return t;
          });
          return themeCache[path];
        }
        function phOf(sp) {
          var ph = descOf(kidOf(sp, "nvSpPr") || kidOf(sp, "nvPicPr") || kidOf(sp, "nvGraphicFramePr"), "ph")[0];
          return ph ? { type: ph.getAttribute("type") || "obj", idx: ph.getAttribute("idx") } : null;
        }
        function phClass(t) { return t === "title" || t === "ctrTitle" ? "title" : (t === "body" || t === "subTitle" || t === "obj" || t === "tbl" || t === "chart" || t === "pic" || t === "dgm" || t === "media" || t === "clipArt") ? "body" : "other"; }
        function findPh(doc, ph) {
          if (!doc || !ph) return null;
          var list = descOf(doc.documentElement, "sp").filter(function (s) { return phOf(s); });
          var hit = null;
          if (ph.idx != null) hit = list.filter(function (s) { return phOf(s).idx === ph.idx; })[0];
          if (!hit) hit = list.filter(function (s) { return phOf(s).type === ph.type; })[0];
          if (!hit && phClass(ph.type) !== "other") hit = list.filter(function (s) { return phClass(phOf(s).type) === phClass(ph.type); })[0];
          return hit || null;
        }
        function xfrm(el) {
          var x = el && (kidOf(el, "spPr", "xfrm") || kidOf(el, "xfrm") || kidOf(el, "grpSpPr", "xfrm"));
          if (!x) return null;
          var off = kidOf(x, "off"), ext = kidOf(x, "ext");
          if (!off || !ext) return null;
          return { x: px(attrNum(off, "x", 0)), y: px(attrNum(off, "y", 0)), w: px(attrNum(ext, "cx", 0)), h: px(attrNum(ext, "cy", 0)),
                   rot: attrNum(x, "rot", 0) / 60000, flipH: attrBool(x, "flipH", false), flipV: attrBool(x, "flipV", false) };
        }

        var slides = [], si = 0;
        function nextSlide() {
          if (si >= ids.length) return Promise.resolve();
          var path = ids[si++], num = si;
          var ctx = { slide: null, layout: null, master: null, theme: null, colorOf: null, num: num };
          return layer(path).then(function (s) {
            ctx.slide = s;
            if (!s.doc) return null;
            if (attrBool(s.doc.documentElement, "show", true) === false) return null;   // diapositiva oculta
            var lr = relByType(s.rels, "/slideLayout");
            return (lr ? layer(lr.target) : Promise.resolve(null)).then(function (l) {
              ctx.layout = l;
              var mr = l && relByType(l.rels, "/slideMaster");
              return mr ? layer(mr.target) : null;
            }).then(function (m) {
              ctx.master = m;
              var tr = m && relByType(m.rels, "/theme");
              return tr ? themeOf(tr.target) : { $fonts: {}, $ln: [], $eff: [], $fill: [], $bgFill: [] };
            }).then(function (theme) {
              ctx.theme = theme;
              var clrMap = {};
              var cm = ctx.master && ctx.master.doc && kidOf(ctx.master.doc.documentElement, "clrMap");
              if (cm) Array.prototype.forEach.call(cm.attributes, function (a) { clrMap[a.name] = a.value; });
              var ov = kidOf(s.doc.documentElement, "clrMapOvr", "overrideClrMapping");
              if (ov) Array.prototype.forEach.call(ov.attributes, function (a) { clrMap[a.name] = a.value; });
              ctx.colorOf = makeColorOf(theme, clrMap);
              return renderSlide(ctx);
            }).then(function (html) { if (html != null) slides.push({ html: html.html, bg: html.bg, text: html.text }); });
          }).then(function () { if (onProgress) onProgress(0.05 + 0.2 * si / ids.length); }).then(K.yieldUI).then(nextSlide);
        }

        /* ---- una diapositiva completa ---- */
        function renderSlide(ctx) {
          var colorOf = ctx.colorOf, theme = ctx.theme, text = [];
          var masterDoc = ctx.master && ctx.master.doc, layoutDoc = ctx.layout && ctx.layout.doc, slideDoc = ctx.slide.doc;
          var txStyles = masterDoc && kidOf(masterDoc.documentElement, "txStyles");
          function fontName(tf) {
            if (!tf) return null;
            if (tf === "+mj-lt" || tf === "+mj-ea") return theme.$fonts.major || "Calibri";
            if (tf === "+mn-lt" || tf === "+mn-ea") return theme.$fonts.minor || "Calibri";
            return tf;
          }
          /* Relleno → {css, svg} */
          function fillOf(spPr, styleEl, phClr) {
            if (spPr) {
              if (kidOf(spPr, "noFill")) return { none: true };
              var sf = kidOf(spPr, "solidFill");
              if (sf) return { color: colorOf(sf, phClr) };
              var gf = kidOf(spPr, "gradFill");
              if (gf) return gradOf(gf, phClr);
              var bf = kidOf(spPr, "blipFill");
              if (bf) return { blip: bf };
              var pf = kidOf(spPr, "pattFill");
              if (pf) {
                var fg = colorOf(kidOf(pf, "fgClr"), phClr) || [0, 0, 0, 1], bgc = colorOf(kidOf(pf, "bgClr"), phClr) || [255, 255, 255, 1];
                return { patt: { prst: pf.getAttribute("prst") || "pct50", fg: fg, bg: bgc }, color: mix(fg, bgc, pattDensity(pf.getAttribute("prst") || "")) };
              }
              if (kidOf(spPr, "grpFill")) return { group: true };
            }
            var fr = styleEl && kidOf(styleEl, "fillRef");
            if (fr) {
              var idx = attrNum(fr, "idx", 0);
              if (!idx) return { none: true };
              var base = colorOf(fr);
              var tf = idx >= 1000 ? theme.$bgFill[idx - 1001] : theme.$fill[idx - 1];
              if (tf && tf.localName === "gradFill") return gradOf(tf, base);
              if (tf && tf.localName === "noFill") return { none: true };
              return { color: base };
            }
            return null;
          }
          function gradOf(gf, phClr) {
            var stops = descOf(gf, "gs").map(function (g) { return { pos: attrNum(g, "pos", 0) / 1000, c: colorOf(g, phClr) }; }).filter(function (s) { return s.c; });
            stops.sort(function (a, b) { return a.pos - b.pos; });
            if (!stops.length) return null;
            var lin = kidOf(gf, "lin"), path = kidOf(gf, "path");
            return { grad: { stops: stops, ang: lin ? attrNum(lin, "ang", 0) / 60000 : 90, radial: !!path, rot: attrBool(gf, "rotWithShape", true) } };
          }
          function lineOf(spPr, styleEl) {
            var ln = spPr && kidOf(spPr, "ln");
            var ref = styleEl && kidOf(styleEl, "lnRef");
            var w = ln && ln.getAttribute("w") != null ? px(attrNum(ln, "w", 0)) : null, color = null, none = false, dash = null;
            if (ln) {
              if (kidOf(ln, "noFill")) none = true;
              var sf = kidOf(ln, "solidFill");
              if (sf) color = colorOf(sf);
              var gf = kidOf(ln, "gradFill");
              if (!color && gf) { var g1 = descOf(gf, "gs")[0]; color = g1 && colorOf(g1); }
              var pd = kidOf(ln, "prstDash");
              if (pd) dash = pd.getAttribute("val");
            }
            if (!none && !color && ref && attrNum(ref, "idx", 0) > 0) { color = colorOf(ref); if (w == null) w = px(theme.$ln[attrNum(ref, "idx", 1) - 1] || 9525); }
            if (none || !color) return null;
            if (w == null) w = px(9525);
            return { color: color, w: Math.max(0.5, w), dash: dash,
              head: ln && kidOf(ln, "headEnd") && kidOf(ln, "headEnd").getAttribute("type") !== "none" ? kidOf(ln, "headEnd").getAttribute("type") : null,
              tail: ln && kidOf(ln, "tailEnd") && kidOf(ln, "tailEnd").getAttribute("type") !== "none" ? kidOf(ln, "tailEnd").getAttribute("type") : null };
          }
          function shadowOf(spPr, styleEl) {
            var eff = spPr && kidOf(spPr, "effectLst");
            if (!eff && spPr && kidOf(spPr, "effectDag")) return "";
            if (!eff) {
              var er = styleEl && kidOf(styleEl, "effectRef"), idx = er ? attrNum(er, "idx", 0) : 0;
              var es = idx > 0 && theme.$eff[idx - 1];
              eff = es && kidOf(es, "effectLst");
            }
            var sh = eff && kidOf(eff, "outerShdw");
            if (!sh) return "";
            var dist = px(attrNum(sh, "dist", 0)), dir = attrNum(sh, "dir", 0) / 60000 * Math.PI / 180, blur = px(attrNum(sh, "blurRad", 0));
            var c = colorOf(sh) || [0, 0, 0, 0.35];
            return "filter:drop-shadow(" + fmt(Math.cos(dir) * dist) + "px " + fmt(Math.sin(dir) * dist) + "px " + fmt(blur / 2) + "px " + rgba(c) + ");";
          }
          var gradSeq = 0;
          function svgPaint(fill, defs, xf) {
            if (!fill || fill.none || fill.group) return "none";
            if (fill.patt) {
              var pid = "p" + ctx.num + "_" + (gradSeq++), P = fill.patt, pr = P.prst, sz = /^sm|nar/.test(pr) ? 4 : /^lg|wd/.test(pr) ? 10 : 6;
              var sw = /^dk/.test(pr) ? 2 : 1, ln = "";
              var stroke = ' stroke="' + rgba(P.fg) + '" stroke-width="' + sw + '"';
              if (/Grid|cross$/i.test(pr) && !/diag/i.test(pr)) ln = '<path d="M0,0.5 H' + sz + " M0.5,0 V" + sz + '"' + stroke + "/>";
              else if (/Horz/i.test(pr)) ln = '<path d="M0,0.5 H' + sz + '"' + stroke + "/>";
              else if (/Vert/i.test(pr)) ln = '<path d="M0.5,0 V' + sz + '"' + stroke + "/>";
              else if (/dnDiag/i.test(pr)) ln = '<path d="M0,0 L' + sz + "," + sz + '"' + stroke + "/>";
              else if (/upDiag/i.test(pr)) ln = '<path d="M0,' + sz + " L" + sz + ',0"' + stroke + "/>";
              else if (/diagCross|Check|Confetti|weave|trellis/i.test(pr)) ln = '<path d="M0,0 L' + sz + "," + sz + " M0," + sz + " L" + sz + ',0"' + stroke + "/>";
              else return rgba(fill.color);
              defs.push('<pattern id="' + pid + '" width="' + sz + '" height="' + sz + '" patternUnits="userSpaceOnUse"><rect width="' + sz + '" height="' + sz + '" fill="' + rgba(P.bg) + '"/>' + ln + "</pattern>");
              return "url(#" + pid + ")";
            }
            if (fill.color) return rgba(fill.color);
            if (fill.grad) {
              var id = "g" + ctx.num + "_" + (gradSeq++), g = fill.grad;
              var stops = g.stops.map(function (s) { return '<stop offset="' + s.pos + '%" stop-color="' + rgba([s.c[0], s.c[1], s.c[2]]) + '" stop-opacity="' + (s.c[3] != null ? s.c[3] : 1) + '"/>'; }).join("");
              if (g.radial) defs.push('<radialGradient id="' + id + '" cx="50%" cy="50%" r="70%">' + stops + "</radialGradient>");
              else {
                var ang = g.ang;
                // Sin «girar con la forma»: el degradado no se refleja aunque la forma sí
                if (!g.rot && xf) { if (xf.flipH) ang = 180 - ang; if (xf.flipV) ang = -ang; }
                var a = ang * Math.PI / 180, cx = Math.cos(a), cy = Math.sin(a);
                defs.push('<linearGradient id="' + id + '" x1="' + fmt(50 - cx * 50) + '%" y1="' + fmt(50 - cy * 50) + '%" x2="' + fmt(50 + cx * 50) + '%" y2="' + fmt(50 + cy * 50) + '%">' + stops + "</linearGradient>");
              }
              return "url(#" + id + ")";
            }
            return "none";
          }
          function cssPaint(fill) {
            if (!fill || fill.none) return "";
            if (fill.color) return rgba(fill.color);
            if (fill.grad) {
              var g = fill.grad, st = g.stops.map(function (s) { return rgba(s.c) + " " + s.pos + "%"; }).join(",");
              return g.radial ? "radial-gradient(circle," + st + ")" : "linear-gradient(" + fmt(g.ang + 90) + "deg," + st + ")";
            }
            return "";
          }

          /* Cadena de estilos de párrafo para un nivel */
          function lvlOf(container, lvl) { return container ? kidOf(container, "lvl" + (lvl + 1) + "pPr") : null; }
          function textChain(sp, ph, lsp, msp) {
            var cls = ph ? phClass(ph.type) : null;
            var masterStyle = txStyles && (cls === "title" ? kidOf(txStyles, "titleStyle") : cls === "body" ? kidOf(txStyles, "bodyStyle") : kidOf(txStyles, "otherStyle"));
            var chain = [kidOf(sp, "txBody", "lstStyle")];
            if (ph) chain.push(lsp && kidOf(lsp, "txBody", "lstStyle"), msp && kidOf(msp, "txBody", "lstStyle"), masterStyle);
            else chain.push(defText, txStyles && kidOf(txStyles, "otherStyle"));
            return chain.filter(Boolean);
          }
          /* Texto de una forma → HTML */
          function textHtml(txBody, chain, opts) {
            opts = opts || {};
            var paras = kidsOf(txBody, "p"), html = "", autoNum = {};
            var scale = opts.fontScale || 1, lnRed = opts.lnSpcReduction || 0;
            paras.forEach(function (p) {
              var pPr = kidOf(p, "pPr"), lvl = attrNum(pPr, "lvl", 0);
              var levels = [pPr].concat(chain.map(function (c) { return lvlOf(c, lvl); })).filter(Boolean);
              function pick(fn) { for (var k = 0; k < levels.length; k++) { var v = fn(levels[k]); if (v != null) return v; } return null; }
              function defR(fn) { return pick(function (l) { var d = kidOf(l, "defRPr"); return d ? fn(d) : null; }); }
              var algn = pick(function (l) { return l.getAttribute("algn"); }) || "l";
              var marL = px(pick(function (l) { return l.getAttribute("marL") != null ? attrNum(l, "marL", 0) : null; }) || 0);
              var indent = px(pick(function (l) { return l.getAttribute("indent") != null ? attrNum(l, "indent", 0) : null; }) || 0);
              var lnSpc = pick(function (l) { var x = kidOf(l, "lnSpc"); return x ? x : null; });
              var spcB = pick(function (l) { return kidOf(l, "spcBef"); }), spcA = pick(function (l) { return kidOf(l, "spcAft"); });
              // Tamaño de letra base del párrafo
              var defSz = defR(function (d) { return d.getAttribute("sz") != null ? attrNum(d, "sz", 1800) / 100 : null; }) || 18;
              var runs = kidsOf(p).filter(function (r) { return r.localName === "r" || r.localName === "br" || r.localName === "fld"; });
              var firstSz = defSz, firstFont = null, firstColor = null;
              var inner = runs.map(function (r, ri) {
                if (r.localName === "br") return "<br>";
                var rPr = kidOf(r, "rPr");
                function prop(name, def) {
                  if (rPr && rPr.getAttribute(name) != null) return rPr.getAttribute(name);
                  var v = defR(function (d) { return d.getAttribute(name); });
                  return v != null ? v : def;
                }
                var t = (kidOf(r, "t") || {}).textContent || "";
                if (r.localName === "fld" && /slidenum/i.test(r.getAttribute("type") || "")) t = String(ctx.num);
                var sz = attrNum({ getAttribute: function () { return prop("sz", null); } }, "x", defSz * 100) / 100 * scale;
                var b = prop("b", "0") === "1" || prop("b", "0") === "true", it = prop("i", "0") === "1" || prop("i", "0") === "true";
                var u = prop("u", "none"), strike = prop("strike", "noStrike"), cap = prop("cap", "none"), base = attrNum({ getAttribute: function () { return prop("baseline", null); } }, "x", 0);
                var spc = attrNum({ getAttribute: function () { return prop("spc", null); } }, "x", 0);
                var latin = rPr && kidOf(rPr, "latin") ? kidOf(rPr, "latin").getAttribute("typeface") : defR(function (d) { var l = kidOf(d, "latin"); return l ? l.getAttribute("typeface") : null; });
                var font = fontName(latin) || opts.font || theme.$fonts.minor || "Calibri";
                var col = null;
                var fillEl = rPr && (kidOf(rPr, "solidFill") || kidOf(rPr, "gradFill"));
                if (fillEl) col = fillEl.localName === "solidFill" ? colorOf(fillEl) : colorOf(descOf(fillEl, "gs")[0]);
                if (!col) {
                  // Prioridad de PowerPoint: párrafo/lista propia → color del estilo de la forma → estilos heredados
                  var local = [pPr, lvlOf(opts.ownLst, lvl)].filter(Boolean), lf = null;
                  for (var q = 0; q < local.length && !lf; q++) { var dd = kidOf(local[q], "defRPr"); lf = dd && kidOf(dd, "solidFill"); }
                  if (lf) col = colorOf(lf);
                  else if (opts.color) col = opts.color;
                  else { var dfill = defR(function (d) { return kidOf(d, "solidFill"); }); if (dfill) col = colorOf(dfill); }
                }
                if (rPr && kidOf(rPr, "hlinkClick")) { col = col || (theme.hlink ? theme.hlink.concat([1]) : null); if (u === "none") u = "sng"; }
                if (!col) col = opts.color || null;
                if (!col) col = (theme.dk1 || [0, 0, 0]).concat([1]);
                if (cap === "all") t = t.toUpperCase();
                var hl = rPr && kidOf(rPr, "highlight") ? colorOf(kidOf(rPr, "highlight")) : null;
                if (ri === 0 || firstFont == null) { firstSz = sz; firstFont = font; firstColor = col; }
                var st = "font-family:" + fontStack('"' + font + '"') + ";font-size:" + fmt(sz) + "pt;color:" + rgba(col) + ";" + (b ? "font-weight:700;" : "") + (it ? "font-style:italic;" : "") +
                  (u !== "none" || strike !== "noStrike" ? "text-decoration:" + (u !== "none" ? "underline " : "") + (strike !== "noStrike" ? "line-through" : "") + ";" : "") +
                  (cap === "small" ? "font-variant:small-caps;" : "") + (spc ? "letter-spacing:" + fmt(spc / 100) + "pt;" : "") +
                  (base ? "vertical-align:" + (base > 0 ? "super" : "sub") + ";font-size:" + fmt(sz * 0.7) + "pt;" : "") + (hl ? "background:" + rgba(hl) + ";" : "");
                if (t) text.push(t);
                return '<span style="' + escH(st) + '">' + escH(t) + "</span>";
              }).join("");
              if (!runs.length) {
                var end = kidOf(p, "endParaRPr");
                if (end && end.getAttribute("sz")) firstSz = attrNum(end, "sz", 1800) / 100 * scale; else firstSz = defSz * scale;
              }
              var fontForLh = firstFont || opts.font || theme.$fonts.minor || "Calibri";
              var single = LINE[firstFamily(fontForLh)] || 1.2;
              var lh = "line-height:" + fmt(single) + ";";
              if (lnSpc) {
                var pct = kidOf(lnSpc, "spcPct"), pts = kidOf(lnSpc, "spcPts");
                if (pct) lh = "line-height:" + fmt(single * attrNum(pct, "val", 100000) / 100000 * (1 - lnRed)) + ";";
                else if (pts) lh = "line-height:" + fmt(attrNum(pts, "val", 1200) / 100) + "pt;";
              } else if (lnRed) lh = "line-height:" + fmt(single * (1 - lnRed)) + ";";
              function spc(el) {
                if (!el) return 0;
                var a = kidOf(el, "spcPts"), b2 = kidOf(el, "spcPct");
                if (a) return attrNum(a, "val", 0) / 100;
                if (b2) return attrNum(b2, "val", 0) / 100000 * firstSz * single;
                return 0;
              }
              // Viñeta
              var bullet = "";
              var hasText = runs.some(function (r) { return r.localName !== "br" && ((kidOf(r, "t") || {}).textContent || "").length; });
              var buNone = pick(function (l) { return kidOf(l, "buNone") ? true : (kidOf(l, "buChar") || kidOf(l, "buAutoNum") || kidOf(l, "buBlip")) ? false : null; });
              if (hasText && buNone === false) {
                var bc = pick(function (l) { return kidOf(l, "buChar"); }), ba = pick(function (l) { return kidOf(l, "buAutoNum"); });
                var bfont = pick(function (l) { var f = kidOf(l, "buFont"); return f ? f.getAttribute("typeface") : null; });
                var bclr = pick(function (l) { var c = kidOf(l, "buClr"); return c ? colorOf(c) : null; }) || firstColor;
                var bsz = pick(function (l) { var s2 = kidOf(l, "buSzPct"); return s2 ? attrNum(s2, "val", 100000) / 100000 : null; }) || 1;
                var mark = "•";
                if (bc) {
                  mark = bc.getAttribute("char") || "•";
                  if (/wingdings|symbol/i.test(bfont || "")) mark = ({ "§": "▪", "Ø": "➢", "ü": "✓", "q": "❑", "v": "❖", "l": "●", "n": "■", "Ü": "➢", "à": "➔", "ð": "➢", "o": "○", "": "•" })[mark] || fixPua(mark);
                  bfont = /wingdings|symbol/i.test(bfont || "") ? null : bfont;
                } else if (ba) {
                  var typ = ba.getAttribute("type") || "arabicPeriod", start = attrNum(ba, "startAt", 1);
                  autoNum[lvl] = autoNum[lvl] != null ? autoNum[lvl] + 1 : start;
                  var n = autoNum[lvl];
                  var body = /alphaLc/.test(typ) ? String.fromCharCode(96 + ((n - 1) % 26) + 1) : /alphaUc/.test(typ) ? String.fromCharCode(64 + ((n - 1) % 26) + 1) :
                    /romanLc/.test(typ) ? roman(n).toLowerCase() : /romanUc/.test(typ) ? roman(n) : String(n);
                  mark = /ParenBoth/.test(typ) ? "(" + body + ")" : /ParenR/.test(typ) ? body + ")" : /Plain/.test(typ) ? body : body + ".";
                }
                var hang = Math.max(0, -indent);
                bullet = '<span style="' + escH("position:absolute;left:" + fmt(marL + indent) + "px;font-size:" + fmt(firstSz * bsz) + "pt;color:" + rgba(bclr) + ";" + (bfont ? "font-family:" + fontStack('"' + fontName(bfont) + '"') + ";" : "font-family:" + fontStack('"' + fontForLh + '"') + ";")) + '">' + escH(mark) + "</span>";
                if (!hang) indent = 0;
              }
              if (buNone !== false || !hasText) Object.keys(autoNum).forEach(function (k) { if (+k >= lvl) delete autoNum[k]; });
              var ta = { l: "left", ctr: "center", r: "right", just: "justify", dist: "justify" }[algn] || "left";
              var pst = "position:relative;margin:0;padding:" + fmt(spc(spcB)) + "pt 0 " + fmt(spc(spcA)) + "pt " + fmt(marL) + "px;text-indent:" + fmt(bullet ? 0 : indent) + "px;text-align:" + ta + ";" + lh +
                "font-size:" + fmt(firstSz) + "pt;font-family:" + fontStack('"' + fontForLh + '"') + ";" + (opts.nowrap ? "white-space:pre;" : "white-space:pre-wrap;overflow-wrap:break-word;");
              html += '<p style="' + escH(pst) + '">' + bullet + (inner || "&#8203;") + "</p>";
            });
            return html;
          }
          function roman(n) { var r = "", m = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]]; m.forEach(function (x) { while (n >= x[0]) { r += x[1]; n -= x[0]; } }); return r; }

          /* Caja con texto (bodyPr: márgenes internos, anclaje, ajuste) */
          function textBox(txBody, box, chain, bodies, opts) {
            if (!txBody || !kidsOf(txBody, "p").some(function (p) { return descOf(p, "t").some(function (t) { return t.textContent.length; }); })) return "";
            function bp(name, def) { for (var i = 0; i < bodies.length; i++) { var b = bodies[i] && kidOf(bodies[i], "bodyPr"); if (b && b.getAttribute(name) != null) return b.getAttribute(name); } return def; }
            var ownBp = kidOf(txBody, "bodyPr");
            var auto = ownBp && kidOf(ownBp, "normAutofit");
            var l = px(+bp("lIns", 91440)), r = px(+bp("rIns", 91440)), t = px(+bp("tIns", 45720)), b = px(+bp("bIns", 45720));
            var anchor = bp("anchor", "t"), wrap = bp("wrap", "square"), vert = bp("vert", "horz");
            var o = Object.assign({}, opts, { ownLst: kidOf(txBody, "lstStyle"), fontScale: auto ? attrNum(auto, "fontScale", 100000) / 100000 : 1, lnSpcReduction: auto ? attrNum(auto, "lnSpcReduction", 0) / 100000 : 0, nowrap: wrap === "none" });
            var inner = textHtml(txBody, chain, o);
            var jc = anchor === "ctr" ? "center" : anchor === "b" ? "flex-end" : "flex-start";
            var wm = vert === "vert" || vert === "eaVert" ? "writing-mode:vertical-rl;" : vert === "vert270" ? "writing-mode:vertical-rl;transform:rotate(180deg);" : "";
            return '<div style="position:absolute;left:0;top:0;width:' + fmt(box.w) + "px;height:" + fmt(box.h) + "px;box-sizing:border-box;padding:" + fmt(t) + "px " + fmt(r) + "px " + fmt(b) + "px " + fmt(l) + "px;display:flex;flex-direction:column;justify-content:" + jc + ";" + wm + '">' +
              '<div style="' + (wrap === "none" ? "" : "width:100%;") + '">' + inner + "</div></div>";
          }

          /* Una forma (sp) */
          function shapeHtml(sp, map, isLayer, rels, layerDocs) {
            var ph = phOf(sp);
            if (isLayer && ph) return "";   // los marcadores del patrón/diseño no se imprimen
            var cNvPr = descOf(kidOf(sp, "nvSpPr"), "cNvPr")[0];
            if (cNvPr && attrBool(cNvPr, "hidden", false)) return "";
            var lsp = ph && !isLayer ? findPh(layoutDoc, ph) : null, msp = ph && !isLayer ? findPh(masterDoc, ph) : null;
            if (lsp && !msp) { var lph = phOf(lsp); msp = findPh(masterDoc, lph); }
            var xf = xfrm(sp) || xfrm(lsp) || xfrm(msp);
            if (!xf) return "";
            var box = map(xf);
            var spPr = kidOf(sp, "spPr"), styleEl = kidOf(sp, "style");
            var fill = fillOf(spPr, styleEl) || (lsp && fillOf(kidOf(lsp, "spPr"), null)) || (msp && fillOf(kidOf(msp, "spPr"), null));
            var line = lineOf(spPr, styleEl) || (lsp && lineOf(kidOf(lsp, "spPr"), null)) || (msp && lineOf(kidOf(msp, "spPr"), null));
            var geomEl = (spPr && (kidOf(spPr, "prstGeom") || kidOf(spPr, "custGeom"))) || (lsp && kidOf(lsp, "spPr") && (kidOf(kidOf(lsp, "spPr"), "prstGeom"))) || null;
            var prst = geomEl && geomEl.localName === "prstGeom" ? geomEl.getAttribute("prst") : geomEl ? "cust" : "rect";
            var av = {};
            if (geomEl) descOf(geomEl, "gd").forEach(function (g) { var m = /val\s+(-?\d+)/.exec(g.getAttribute("fmla") || ""); if (m) av[g.getAttribute("name")] = +m[1]; });
            var d = prst === "cust" ? customPath(geomEl, box.w, box.h) : presetPath(prst, box.w, box.h, av);
            if (fill && fill.blip) fill = null;   // relleno de imagen en forma: se omite el fondo (poco común)
            var defs = [], shapeSvg = "";
            var shadow = shadowOf(spPr, styleEl);
            var dash = line && line.dash && line.dash !== "solid" ? ' stroke-dasharray="' + (/dot/i.test(line.dash) ? fmt(line.w) + " " + fmt(line.w * 2) : fmt(line.w * 4) + " " + fmt(line.w * 3)) + '"' : "";
            var strokeAttrs = line ? ' stroke="' + rgba([line.color[0], line.color[1], line.color[2]]) + '" stroke-opacity="' + (line.color[3] != null ? line.color[3] : 1) + '" stroke-width="' + fmt(line.w) + '"' + dash : ' stroke="none"';
            if (d === "LINE") {
              if (line) {
                var x1 = xf.flipH ? box.w : 0, y1 = xf.flipV ? box.h : 0, x2 = xf.flipH ? 0 : box.w, y2 = xf.flipV ? 0 : box.h;
                var mk = "";
                if (line.tail || line.head) {
                  var mid = "m" + ctx.num + "_" + (gradSeq++);
                  defs.push('<marker id="' + mid + '" viewBox="0 0 10 10" refX="8" refY="5" markerWidth="4" markerHeight="4" orient="auto-start-reverse"><path d="M0,0 L10,5 L0,10 z" fill="' + rgba(line.color) + '"/></marker>');
                  mk = (line.tail ? ' marker-end="url(#' + mid + ')"' : "") + (line.head ? ' marker-start="url(#' + mid + ')"' : "");
                }
                shapeSvg = '<line x1="' + fmt(x1) + '" y1="' + fmt(y1) + '" x2="' + fmt(x2) + '" y2="' + fmt(y2) + '"' + strokeAttrs + mk + ' stroke-linecap="round"/>';
              }
            } else if (fill || line) {
              var paint = svgPaint(fill, defs, xf);
              var flip = (xf.flipH || xf.flipV) ? ' transform="translate(' + (xf.flipH ? fmt(box.w) : 0) + "," + (xf.flipV ? fmt(box.h) : 0) + ") scale(" + (xf.flipH ? -1 : 1) + "," + (xf.flipV ? -1 : 1) + ')"' : "";
              var fillOpacity = fill && fill.color && fill.color[3] != null ? ' fill-opacity="' + fmt(fill.color[3]) + '"' : "";
              if (d) shapeSvg = '<path d="' + d + '" fill="' + paint + '"' + fillOpacity + strokeAttrs + flip + ' fill-rule="evenodd"/>';
              else shapeSvg = '<rect x="0" y="0" width="' + fmt(box.w) + '" height="' + fmt(box.h) + '" fill="' + paint + '"' + fillOpacity + strokeAttrs + flip + "/>";
            }
            var svg = shapeSvg ? '<svg width="' + fmt(box.w) + '" height="' + fmt(box.h) + '" viewBox="0 0 ' + fmt(box.w) + " " + fmt(box.h) + '" style="position:absolute;left:0;top:0;overflow:visible;' + shadow + '">' + (defs.length ? "<defs>" + defs.join("") + "</defs>" : "") + shapeSvg + "</svg>" : "";
            // Color de texto por defecto de la forma (fontRef del estilo)
            var fref = styleEl && kidOf(styleEl, "fontRef");
            var tOpts = { color: fref ? colorOf(fref) : null, font: fref ? fontName(fref.getAttribute("idx") === "major" ? "+mj-lt" : "+mn-lt") : null };
            if (ph && phClass(ph.type) === "title" && !tOpts.font) tOpts.font = theme.$fonts.major;
            var txt = textBox(kidOf(sp, "txBody"), box, textChain(sp, ph, lsp, msp), [kidOf(sp, "txBody"), lsp && kidOf(lsp, "txBody"), msp && kidOf(msp, "txBody")], tOpts);
            if (!svg && !txt) return "";
            return '<div style="position:absolute;left:' + fmt(box.x) + "px;top:" + fmt(box.y) + "px;width:" + fmt(box.w) + "px;height:" + fmt(box.h) + "px;" + (xf.rot ? "transform:rotate(" + fmt(xf.rot) + "deg);" : "") + '">' + svg + txt + "</div>";
          }

          /* Imagen (pic) */
          function picHtml(pic, map, rels, isLayer) {
            var ph = phOf(pic);
            var lsp = ph && !isLayer ? findPh(layoutDoc, ph) : null;
            var xf = xfrm(pic) || xfrm(lsp);
            var bf = kidOf(pic, "blipFill"), blip = bf && kidOf(bf, "blip");
            if (!xf || !blip) return Promise.resolve("");
            // Preferir la versión SVG si existe
            var svgB = descOf(blip, "svgBlip")[0];
            var rel = rels[relId(svgB, "embed") || relId(blip, "embed")] || rels[relId(blip, "embed")];
            if (!rel || rel.external) return Promise.resolve("");
            var box = map(xf), src = kidOf(bf, "srcRect");
            var cl = attrNum(src, "l", 0) / 100000, ct = attrNum(src, "t", 0) / 100000, cr = attrNum(src, "r", 0) / 100000, cb = attrNum(src, "b", 0) / 100000;
            var alpha = descOf(blip, "alphaModFix")[0], op = alpha ? attrNum(alpha, "amt", 100000) / 100000 : 1;
            var spPr = kidOf(pic, "spPr"), geom = spPr && kidOf(spPr, "prstGeom"), prst = geom ? geom.getAttribute("prst") : "rect";
            var clip = prst === "ellipse" ? "border-radius:50%;" : prst === "roundRect" ? "border-radius:" + fmt(Math.min(box.w, box.h) * 0.16667) + "px;" : "";
            var line = lineOf(spPr, kidOf(pic, "style"));
            var shadow = shadowOf(spPr, kidOf(pic, "style"));
            return media(rel.target).then(function (url) {
              if (!url) return "";
              var iw = box.w / Math.max(0.01, 1 - cl - cr), ih = box.h / Math.max(0.01, 1 - ct - cb);
              var flip = xf.flipH || xf.flipV ? "transform:scale(" + (xf.flipH ? -1 : 1) + "," + (xf.flipV ? -1 : 1) + ");" : "";
              return '<div style="position:absolute;left:' + fmt(box.x) + "px;top:" + fmt(box.y) + "px;width:" + fmt(box.w) + "px;height:" + fmt(box.h) + "px;" + (xf.rot ? "transform:rotate(" + fmt(xf.rot) + "deg);" : "") + shadow + '">' +
                '<div style="position:absolute;inset:0;overflow:hidden;' + clip + flip + "opacity:" + op + ';">' +
                '<img src="' + url + '" alt="" style="position:absolute;left:' + fmt(-cl * iw) + "px;top:" + fmt(-ct * ih) + "px;width:" + fmt(iw) + "px;height:" + fmt(ih) + 'px;max-width:none"></div>' +
                (line ? '<div style="position:absolute;inset:0;' + clip + "border:" + fmt(line.w) + "px solid " + rgba(line.color) + ';box-sizing:border-box"></div>' : "") + "</div>";
            });
          }

          /* Tabla */
          var BUILTIN_TBL = {
            // Estilo medio 2 - Énfasis 1 (el estilo por defecto de PowerPoint)
            "{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}": { accent: "accent1" }, "{21E4AEA4-8DFA-4A89-87EB-49C32662AFE8}": { accent: "accent2" },
            "{F5AB1C69-6EDB-4FF4-983F-18BD219EF322}": { accent: "accent3" }, "{00A15C55-8517-42AA-B614-E9B94910E393}": { accent: "accent4" },
            "{7DF18680-E054-41AD-8BC1-D1AEF772440D}": { accent: "accent5" }, "{93296810-A885-4BE3-A3E7-6D5BEEA58F35}": { accent: "accent6" },
            "{073A0DAA-6AF3-43AB-8588-CEC1D06C72B9}": { accent: "dk1" }
          };
          function tableHtml(frame, tbl, map) {
            var gxf = kidOf(frame, "xfrm");
            var off = kidOf(gxf, "off");
            var box = map({ x: px(attrNum(off, "x", 0)), y: px(attrNum(off, "y", 0)), w: 0, h: 0 });
            var tblPr = kidOf(tbl, "tblPr");
            var flags = { firstRow: attrBool(tblPr, "firstRow", false), lastRow: attrBool(tblPr, "lastRow", false), firstCol: attrBool(tblPr, "firstCol", false), lastCol: attrBool(tblPr, "lastCol", false), bandRow: attrBool(tblPr, "bandRow", false), bandCol: attrBool(tblPr, "bandCol", false) };
            var sid = (kidOf(tblPr, "tableStyleId") || {}).textContent || (tblStyles && tblStyles.documentElement.getAttribute("def")) || null;
            var styleDef = sid && tblStyles ? descOf(tblStyles.documentElement, "tblStyle").filter(function (s) { return s.getAttribute("styleId") === sid; })[0] : null;
            var tableFill = tblPr && fillOf(tblPr, null);
            function part(name) {
              if (styleDef) return kidOf(styleDef, name);
              return null;
            }
            var builtin = !styleDef && sid && BUILTIN_TBL[sid];
            function partStyle(name) {
              // Devuelve {fill, color, bold, borders{l,r,t,b,h,v}}
              var out = { borders: {} };
              if (builtin) {
                var acc = (theme[builtin.accent] || [68, 114, 196]);
                var light = function (k) { return acc.map(function (x) { return Math.round(x * k + 255 * (1 - k)); }).concat([1]); };
                var white = (theme.lt1 || [255, 255, 255]).concat([1]);
                if (name === "wholeTbl") { out.fill = light(0.2); out.color = (theme.dk1 || [0, 0, 0]).concat([1]); ["l", "r", "t", "b", "h", "v"].forEach(function (k) { out.borders[k] = { w: px(12700), color: white }; }); }
                if (name === "band1H" || name === "band1V") out.fill = light(0.4);
                if (name === "firstRow" || name === "lastRow" || name === "firstCol" || name === "lastCol") { out.fill = acc.concat([1]); out.color = white; out.bold = true; }
                if (name === "lastRow") { out.borders.t = { w: px(38100), color: white }; }
                if (name === "firstRow") { out.borders.b = { w: px(38100), color: white }; }
                return out;
              }
              var el = part(name);
              if (!el) return null;
              var tcTx = kidOf(el, "tcTxStyle"), tcSt = kidOf(el, "tcStyle");
              if (tcTx) {
                out.color = colorOf(tcTx) || (kidOf(tcTx, "fontRef") ? colorOf(kidOf(tcTx, "fontRef")) : null);
                if (tcTx.getAttribute("b") === "on") out.bold = true;
              }
              if (tcSt) {
                var f = kidOf(tcSt, "fill");
                if (f) { var ff = fillOf(f, null); if (ff && ff.color) out.fill = ff.color; }
                var fr = kidOf(tcSt, "fillRef");
                if (fr && !out.fill) out.fill = colorOf(fr);
                var bdr = kidOf(tcSt, "tcBdr");
                if (bdr) [["left", "l"], ["right", "r"], ["top", "t"], ["bottom", "b"], ["insideH", "h"], ["insideV", "v"]].forEach(function (pair) {
                  var e = kidOf(bdr, pair[0]);
                  if (!e) return;
                  var ln = kidOf(e, "ln");
                  if (ln) { if (kidOf(ln, "noFill")) out.borders[pair[1]] = null; else { var c = colorOf(kidOf(ln, "solidFill")); if (c) out.borders[pair[1]] = { w: px(attrNum(ln, "w", 12700)), color: c }; } }
                  var lr = kidOf(e, "lnRef");
                  if (!ln && lr) out.borders[pair[1]] = { w: px(theme.$ln[attrNum(lr, "idx", 1) - 1] || 12700), color: colorOf(lr) };
                });
              }
              return out;
            }
            var whole = partStyle("wholeTbl") || { borders: {} };
            var parts = {};
            ["band1H", "band2H", "band1V", "band2V", "firstRow", "lastRow", "firstCol", "lastCol"].forEach(function (n) { parts[n] = partStyle(n); });
            var gridCols = descOf(kidOf(tbl, "tblGrid"), "gridCol").map(function (g) { return px(attrNum(g, "w", 0)); });
            var rows = kidsOf(tbl, "tr"), nR = rows.length, nC = gridCols.length;
            var html = '<table style="position:absolute;left:' + fmt(box.x) + "px;top:" + fmt(box.y) + "px;border-collapse:collapse;table-layout:fixed;width:" + fmt(gridCols.reduce(function (s, v) { return s + v; }, 0)) + 'px">';
            html += "<colgroup>" + gridCols.map(function (w) { return '<col style="width:' + fmt(w) + 'px">'; }).join("") + "</colgroup>";
            rows.forEach(function (tr, ri) {
              html += '<tr style="height:' + fmt(px(attrNum(tr, "h", 370840))) + 'px">';
              var ci = 0;
              kidsOf(tr, "tc").forEach(function (tc) {
                var myCol = ci;
                ci += attrNum(tc, "gridSpan", 1);
                if (attrBool(tc, "hMerge", false) || attrBool(tc, "vMerge", false)) return;
                // Estilo de tabla por zonas (orden de prioridad de PowerPoint)
                var applied = [whole];
                if (flags.bandRow) { var bh = (ri - (flags.firstRow ? 1 : 0)) % 2 === 0 ? parts.band1H : parts.band2H; if (bh && !(flags.firstRow && ri === 0) && !(flags.lastRow && ri === nR - 1)) applied.push(bh); }
                if (flags.bandCol) { var bv = (myCol - (flags.firstCol ? 1 : 0)) % 2 === 0 ? parts.band1V : parts.band2V; if (bv) applied.push(bv); }
                if (flags.firstCol && myCol === 0 && parts.firstCol) applied.push(parts.firstCol);
                if (flags.lastCol && myCol === nC - 1 && parts.lastCol) applied.push(parts.lastCol);
                if (flags.firstRow && ri === 0 && parts.firstRow) applied.push(parts.firstRow);
                if (flags.lastRow && ri === nR - 1 && parts.lastRow) applied.push(parts.lastRow);
                var fill = null, color = null, bold = false, bd = { l: null, r: null, t: null, b: null };
                applied.forEach(function (a) {
                  if (!a) return;
                  if (a.fill) fill = a.fill;
                  if (a.color) color = a.color;
                  if (a.bold) bold = true;
                  var B = a.borders || {};
                  if ("l" in B) bd.l = myCol === 0 ? B.l : (B.v !== undefined ? B.v : B.l);
                  if ("r" in B) bd.r = B.r;
                  if ("t" in B) bd.t = ri === 0 ? B.t : (B.h !== undefined ? B.h : B.t);
                  if ("b" in B) bd.b = B.b;
                  if (B.h !== undefined) { if (ri > 0) bd.t = B.h; if (ri < nR - 1) bd.b = B.h; }
                  if (B.v !== undefined) { if (myCol > 0) bd.l = B.v; if (myCol < nC - 1) bd.r = B.v; }
                });
                if (!fill && tableFill && tableFill.color) fill = tableFill.color;
                var tcPr = kidOf(tc, "tcPr");
                if (tcPr) {
                  var own = fillOf(tcPr, null);
                  if (own && own.none) fill = null; else if (own && own.color) fill = own.color;
                  [["lnL", "l"], ["lnR", "r"], ["lnT", "t"], ["lnB", "b"]].forEach(function (pair) {
                    var ln = kidOf(tcPr, pair[0]);
                    if (!ln) return;
                    if (kidOf(ln, "noFill")) bd[pair[1]] = null;
                    else { var c = colorOf(kidOf(ln, "solidFill")); if (c) bd[pair[1]] = { w: px(attrNum(ln, "w", 12700)), color: c }; }
                  });
                }
                var mL = px(attrNum(tcPr, "marL", 91440)), mR = px(attrNum(tcPr, "marR", 91440)), mT = px(attrNum(tcPr, "marT", 45720)), mB = px(attrNum(tcPr, "marB", 45720));
                var anchor = tcPr ? tcPr.getAttribute("anchor") : null;
                function b(x) { return x ? fmt(x.w) + "px solid " + rgba(x.color) : "none"; }
                var st = "padding:" + fmt(mT) + "px " + fmt(mR) + "px " + fmt(mB) + "px " + fmt(mL) + "px;vertical-align:" + (anchor === "ctr" ? "middle" : anchor === "b" ? "bottom" : "top") + ";" +
                  (fill ? "background:" + rgba(fill) + ";" : "") + "border-left:" + b(bd.l) + ";border-right:" + b(bd.r) + ";border-top:" + b(bd.t) + ";border-bottom:" + b(bd.b) + ";overflow:hidden;";
                var chain = [kidOf(tc, "txBody", "lstStyle"), defText, txStyles && kidOf(txStyles, "otherStyle")].filter(Boolean);
                var inner = textHtml(kidOf(tc, "txBody") || tc, chain, { color: color, bold: bold, ownLst: kidOf(tc, "txBody", "lstStyle") });
                if (bold) inner = inner.replace(/<span style="/g, '<span style="font-weight:700;');
                var span = (attrNum(tc, "gridSpan", 1) > 1 ? ' colspan="' + attrNum(tc, "gridSpan", 1) + '"' : "") + (attrNum(tc, "rowSpan", 1) > 1 ? ' rowspan="' + attrNum(tc, "rowSpan", 1) + '"' : "");
                html += "<td" + span + ' style="' + escH(st) + '">' + inner + "</td>";
              });
              html += "</tr>";
            });
            return html + "</table>";
          }

          /* Árbol de formas (con grupos) */
          function treeHtml(tree, map, rels, isLayer) {
            var parts = [];
            kidsOf(tree).forEach(function (el) {
              var n = el.localName;
              if (n === "AlternateContent") { var ch = kidOf(el, "Choice") || kidOf(el, "Fallback"); if (ch) parts.push(treeHtml(ch, map, rels, isLayer)); return; }
              if (n === "grpSp") {
                var gx = kidOf(el, "grpSpPr", "xfrm"), inner = map;
                if (gx) {
                  var off = kidOf(gx, "off"), ext = kidOf(gx, "ext"), cho = kidOf(gx, "chOff"), che = kidOf(gx, "chExt");
                  if (off && ext && cho && che) {
                    var ox = px(attrNum(off, "x", 0)), oy = px(attrNum(off, "y", 0));
                    var sx = attrNum(ext, "cx", 1) / (attrNum(che, "cx", 1) || 1), sy = attrNum(ext, "cy", 1) / (attrNum(che, "cy", 1) || 1);
                    var cx0 = px(attrNum(cho, "x", 0)), cy0 = px(attrNum(cho, "y", 0));
                    inner = function (r) { return map({ x: ox + (r.x - cx0) * sx, y: oy + (r.y - cy0) * sy, w: r.w * sx, h: r.h * sy, rot: r.rot, flipH: r.flipH, flipV: r.flipV }); };
                  }
                }
                parts.push(treeHtml(el, inner, rels, isLayer));
              } else if (n === "sp" || n === "cxnSp") {
                parts.push(shapeHtml(el, map, isLayer, rels));
              } else if (n === "pic") {
                parts.push(picHtml(el, map, rels, isLayer));
              } else if (n === "graphicFrame") {
                var tbl = descOf(el, "tbl")[0];
                if (tbl) { parts.push(tableHtml(el, tbl, map)); return; }
                var rIds = descOf(el, "relIds")[0];
                if (rIds) {
                  // SmartArt: PowerPoint guarda un dibujo ya calculado de sus formas
                  var dm = rels[relId(rIds, "dm")];
                  var gxf = xfrm(el) || { x: 0, y: 0 };
                  var fb = gxf.x !== undefined ? map({ x: gxf.x, y: gxf.y, w: gxf.w || 0, h: gxf.h || 0 }) : { x: 0, y: 0 };
                  parts.push((dm ? xmlPart(zip, dm.target) : Promise.resolve(null)).then(function (dmDoc) {
                    var ext = dmDoc && descOf(dmDoc.documentElement, "dataModelExt")[0];
                    var drRel = ext && rels[ext.getAttribute("relId")];
                    if (!drRel) { var any = relByType(rels, "/diagramDrawing"); drRel = any; }
                    if (!drRel) { notes.charts++; return ""; }
                    return Promise.all([xmlPart(zip, drRel.target), relsOf(zip, drRel.target)]).then(function (dd) {
                      var dtree = dd[0] && descOf(dd[0].documentElement, "spTree")[0];
                      if (!dtree) return "";
                      var dmap = function (r) { return { x: fb.x + r.x, y: fb.y + r.y, w: r.w, h: r.h, rot: r.rot, flipH: r.flipH, flipV: r.flipV }; };
                      return treeHtml(dtree, dmap, dd[1], false);
                    });
                  }));
                  return;
                }
                if (descOf(el, "chart").length) {
                  var cr = rels[relId(descOf(el, "chart")[0], "id")];
                  var cxf = kidOf(el, "xfrm");
                  var cbox = cxf ? map({ x: px(attrNum(kidOf(cxf, "off"), "x", 0)), y: px(attrNum(kidOf(cxf, "off"), "y", 0)), w: px(attrNum(kidOf(cxf, "ext"), "cx", 0)), h: px(attrNum(kidOf(cxf, "ext"), "cy", 0)) }) : null;
                  if (cr && cbox) parts.push(xmlPart(zip, cr.target).then(function (cdoc) { var h = cdoc ? chartSvg(cdoc, cbox, colorOf, theme) : ""; if (!h) notes.charts++; return h; }));
                  else notes.charts++;
                  return;
                }
                var olePic = descOf(el, "pic")[0];
                if (olePic) parts.push(picHtml(olePic, map, rels, isLayer));
              }
            });
            return Promise.all(parts.map(function (p) { return Promise.resolve(p); })).then(function (list) { return list.join(""); });
          }

          // Fondo: diapositiva → diseño → patrón
          function bgOf(doc, rels) {
            var bg = doc && kidOf(doc.documentElement, "cSld", "bg");
            if (!bg) return null;
            var bgPr = kidOf(bg, "bgPr");
            if (bgPr) {
              var bf = kidOf(bgPr, "blipFill");
              if (bf) { var rel = rels[relId(kidOf(bf, "blip"), "embed")]; return rel ? media(rel.target).then(function (u) { return u ? "url(" + u + ") center/100% 100% no-repeat" : "#fff"; }) : "#fff"; }
              var f = fillOf(bgPr, null);
              return cssPaint(f) || "#fff";
            }
            var ref = kidOf(bg, "bgRef");
            if (ref) {
              var idx = attrNum(ref, "idx", 0), base = colorOf(ref);
              var tf = idx >= 1001 ? theme.$bgFill[idx - 1001] : null;
              if (tf && tf.localName === "gradFill") return cssPaint(gradOf(tf, base)) || rgba(base);
              return rgba(base);
            }
            return null;
          }
          var bgP = bgOf(slideDoc, ctx.slide.rels) || bgOf(layoutDoc, ctx.layout && ctx.layout.rels) || bgOf(masterDoc, ctx.master && ctx.master.rels) || "#fff";
          var identity = function (r) { return r; };
          var showMaster = attrBool(slideDoc.documentElement, "showMasterSp", true) && (!layoutDoc || attrBool(layoutDoc.documentElement, "showMasterSp", true));
          var layers = [];
          if (masterDoc && showMaster) layers.push(treeHtml(kidOf(masterDoc.documentElement, "cSld", "spTree"), identity, ctx.master.rels, true));
          if (layoutDoc && attrBool(slideDoc.documentElement, "showMasterSp", true)) layers.push(treeHtml(kidOf(layoutDoc.documentElement, "cSld", "spTree"), identity, ctx.layout.rels, true));
          layers.push(treeHtml(kidOf(slideDoc.documentElement, "cSld", "spTree"), identity, ctx.slide.rels, false));
          return Promise.all([Promise.resolve(bgP)].concat(layers)).then(function (all) {
            return { bg: all[0], html: all.slice(1).join(""), text: text.join(" ") };
          });
        }

        return nextSlide().then(function () {
          if (!slides.length) throw K.UserError("La presentación no tiene diapositivas", "No hay nada que convertir en este archivo.");
          var pages = slides.map(function (s, idx) {
            return {
              sample: idx === 0 ? slides.map(function (x) { return x.text; }).join(" ") : "",
              render: function (opt) {
                var div = document.createElement("div");
                div.style.cssText = "position:relative;overflow:hidden;width:" + fmt(W) + "px;height:" + fmt(H) + "px;background:" + (/url\(/.test(s.bg) ? "#fff" : s.bg) + ";";
                if (/url\(/.test(s.bg)) div.style.background = s.bg;
                div.innerHTML = s.html;
                host.appendChild(div);
                return settle(div).then(function () { return rasterize(div, "#ffffff"); }).then(function (canvas) {
                  var texts = opt && opt.noText ? null : textLayer(div);
                  div.remove();
                  return { canvas: canvas, texts: texts, w: W, h: H };
                });
              }
            };
          });
          pages.cleanup = function () { host.remove(); };
          pages.notes = notes; pages.debug = slides;
          return pages;
        });
      }).catch(function (e) { host.remove(); throw e; });
    }

    /* Gráficos sencillos (barras, columnas, líneas, áreas, circulares) con los valores guardados */
    function chartSvg(cdoc, box, colorOf, theme) {
      var plot = descOf(cdoc.documentElement, "plotArea")[0];
      if (!plot) return "";
      var types = ["barChart", "bar3DChart", "lineChart", "line3DChart", "pieChart", "pie3DChart", "doughnutChart", "areaChart", "area3DChart"];
      var chart = null;
      for (var i = 0; i < types.length && !chart; i++) chart = kidOf(plot, types[i]);
      if (!chart) return "";
      var kind = /bar/.test(chart.localName) ? "bar" : /line/.test(chart.localName) ? "line" : /area/.test(chart.localName) ? "area" : "pie";
      var dir = kind === "bar" && kidOf(chart, "barDir") ? kidOf(chart, "barDir").getAttribute("val") : "col";
      var grouping = kidOf(chart, "grouping") ? kidOf(chart, "grouping").getAttribute("val") : "clustered";
      var acc = ["accent1", "accent2", "accent3", "accent4", "accent5", "accent6"].map(function (k) { return theme[k] || [90, 90, 90]; });
      function cache(el) {
        var c = el && (descOf(el, "numCache")[0] || descOf(el, "strCache")[0]);
        if (!c) return [];
        var out = [];
        descOf(c, "pt").forEach(function (p) { out[attrNum(p, "idx", 0)] = (kidOf(p, "v") || {}).textContent; });
        return out;
      }
      var series = kidsOf(chart, "ser").map(function (s, k) {
        var sp = kidOf(s, "spPr"), sf = sp && kidOf(sp, "solidFill");
        var c = sf ? colorOf(sf) : acc[k % 6].concat([1]);
        var pts = kidsOf(s, "dPt").map(function (d) { var f = kidOf(d, "spPr") && kidOf(kidOf(d, "spPr"), "solidFill"); return { i: attrNum(kidOf(d, "idx"), "val", 0), c: f ? colorOf(f) : null }; });
        return { name: (cache(kidOf(s, "tx"))[0] || ""), cats: cache(kidOf(s, "cat")), vals: cache(kidOf(s, "val")).map(function (v) { return parseFloat(v) || 0; }), color: c, pts: pts };
      });
      if (!series.length) return "";
      var W = box.w, H = box.h, out = "";
      var titleEl = kidOf(cdoc.documentElement, "chart", "title");
      var title = titleEl ? descOf(titleEl, "t").map(function (t) { return t.textContent; }).join("") : (series.length === 1 && kidOf(cdoc.documentElement, "chart", "autoTitleDeleted") && kidOf(cdoc.documentElement, "chart", "autoTitleDeleted").getAttribute("val") === "1" ? "" : (series.length === 1 ? series[0].name : ""));
      var font = fontStack('"' + (theme.$fonts.minor || "Calibri") + '"');
      var top = title ? 34 : 10, legendH = 26, pad = 40;
      if (title) out += '<text x="' + W / 2 + '" y="24" text-anchor="middle" font-size="18" fill="#404040" font-family="' + escH(font) + '">' + escH(title) + "</text>";
      var cats = series[0].cats.length ? series[0].cats : series[0].vals.map(function (_, i) { return String(i + 1); });
      if (kind === "pie") {
        var vals = series[0].vals, tot = vals.reduce(function (s, v) { return s + Math.max(0, v); }, 0) || 1;
        var r = Math.min(W, H - top - legendH) / 2 - 10, cx = W / 2, cy = top + (H - top - legendH) / 2, a0 = -Math.PI / 2;
        var hole = chart.localName === "doughnutChart" ? 0.5 : 0;
        vals.forEach(function (v, i) {
          var a1 = a0 + Math.max(0, v) / tot * Math.PI * 2, pc = series[0].pts.filter(function (p) { return p.i === i; })[0];
          var col = pc && pc.c ? pc.c : acc[i % 6].concat([1]);
          var large = a1 - a0 > Math.PI ? 1 : 0;
          var p = "M" + fmt(cx + r * Math.cos(a0)) + "," + fmt(cy + r * Math.sin(a0)) + " A" + r + "," + r + " 0 " + large + " 1 " + fmt(cx + r * Math.cos(a1)) + "," + fmt(cy + r * Math.sin(a1));
          p += hole ? " L" + fmt(cx + r * hole * Math.cos(a1)) + "," + fmt(cy + r * hole * Math.sin(a1)) + " A" + r * hole + "," + r * hole + " 0 " + large + " 0 " + fmt(cx + r * hole * Math.cos(a0)) + "," + fmt(cy + r * hole * Math.sin(a0)) + " Z" : " L" + cx + "," + cy + " Z";
          out += '<path d="' + p + '" fill="' + rgba(col) + '" stroke="#fff" stroke-width="1.5"/>';
          a0 = a1;
        });
        var lx = 10;
        cats.forEach(function (c, i) {
          var pc = series[0].pts.filter(function (p) { return p.i === i; })[0], col = pc && pc.c ? pc.c : acc[i % 6].concat([1]);
          out += '<rect x="' + lx + '" y="' + (H - 18) + '" width="10" height="10" fill="' + rgba(col) + '"/><text x="' + (lx + 14) + '" y="' + (H - 9) + '" font-size="12" fill="#404040" font-family="' + escH(font) + '">' + escH(c) + "</text>";
          lx += 24 + String(c).length * 7;
        });
      } else {
        var stacked = grouping === "stacked" || grouping === "percentStacked";
        var max = 0, min = 0;
        cats.forEach(function (_, i) {
          if (stacked) { var s = 0; series.forEach(function (se) { s += se.vals[i] || 0; }); max = Math.max(max, s); }
          else series.forEach(function (se) { max = Math.max(max, se.vals[i] || 0); min = Math.min(min, se.vals[i] || 0); });
        });
        if (max === min) max = min + 1;
        var raw = (max - min) / 5, mag = Math.pow(10, Math.floor(Math.log10(raw))), nrm = raw / mag;
        var step = (nrm < 1.5 ? 1 : nrm < 3 ? 2 : nrm < 7 ? 5 : 10) * mag;
        if (max % step === 0) max += step * 0.5;   // como Excel, deja aire sobre la barra más alta
        var top2 = Math.ceil(max / step) * step, bot = Math.floor(min / step) * step;
        var x0 = pad + 10, x1 = W - 14, y0 = top, y1 = H - legendH - 22;
        var horiz = kind === "bar" && dir === "bar";
        function V(v) { return horiz ? x0 + (v - bot) / (top2 - bot) * (x1 - x0) : y1 - (v - bot) / (top2 - bot) * (y1 - y0); }
        for (var g = bot; g <= top2 + 1e-9; g += step) {
          var label = Math.abs(step) < 1 ? g.toFixed(1) : String(Math.round(g));
          if (horiz) out += '<line x1="' + fmt(V(g)) + '" y1="' + y0 + '" x2="' + fmt(V(g)) + '" y2="' + y1 + '" stroke="#d9d9d9"/><text x="' + fmt(V(g)) + '" y="' + (y1 + 14) + '" font-size="11" text-anchor="middle" fill="#595959" font-family="' + escH(font) + '">' + label + "</text>";
          else out += '<line x1="' + x0 + '" y1="' + fmt(V(g)) + '" x2="' + x1 + '" y2="' + fmt(V(g)) + '" stroke="#d9d9d9"/><text x="' + (x0 - 6) + '" y="' + fmt(V(g) + 4) + '" font-size="11" text-anchor="end" fill="#595959" font-family="' + escH(font) + '">' + label + "</text>";
        }
        var n = cats.length, band = ((horiz ? y1 - y0 : x1 - x0)) / Math.max(1, n);
        cats.forEach(function (c, i) {
          var mid = (horiz ? y0 : x0) + band * (i + 0.5);
          if (horiz) out += '<text x="' + (x0 - 6) + '" y="' + fmt(mid + 4) + '" font-size="11" text-anchor="end" fill="#595959" font-family="' + escH(font) + '">' + escH(c) + "</text>";
          else out += '<text x="' + fmt(mid) + '" y="' + (y1 + 14) + '" font-size="11" text-anchor="middle" fill="#595959" font-family="' + escH(font) + '">' + escH(c) + "</text>";
        });
        if (kind === "bar") {
          var bw = band * 0.7 / (stacked ? 1 : series.length);
          cats.forEach(function (_, i) {
            var acc0 = 0;
            series.forEach(function (se, k) {
              var v = se.vals[i] || 0, start = stacked ? acc0 : 0, end = start + v;
              acc0 = end;
              var off = (horiz ? y0 : x0) + band * i + band * 0.15 + (stacked ? 0 : bw * k);
              var a = V(start), b = V(end);
              if (horiz) out += '<rect x="' + fmt(Math.min(a, b)) + '" y="' + fmt(off) + '" width="' + fmt(Math.abs(b - a)) + '" height="' + fmt(bw) + '" fill="' + rgba(se.color) + '"/>';
              else out += '<rect x="' + fmt(off) + '" y="' + fmt(Math.min(a, b)) + '" width="' + fmt(bw) + '" height="' + fmt(Math.abs(b - a)) + '" fill="' + rgba(se.color) + '"/>';
            });
          });
        } else {
          series.forEach(function (se) {
            var pts = cats.map(function (_, i) { return fmt(x0 + band * (i + 0.5)) + "," + fmt(V(se.vals[i] || 0)); });
            if (kind === "area") out += '<polygon points="' + fmt(x0 + band * 0.5) + "," + fmt(V(bot)) + " " + pts.join(" ") + " " + fmt(x0 + band * (n - 0.5)) + "," + fmt(V(bot)) + '" fill="' + rgba(se.color) + '" fill-opacity="0.85"/>';
            else out += '<polyline points="' + pts.join(" ") + '" fill="none" stroke="' + rgba(se.color) + '" stroke-width="2.5"/>';
          });
        }
        var lx2 = 10;
        if (series.length > 1 || kind !== "bar") series.forEach(function (se) {
          out += '<rect x="' + lx2 + '" y="' + (H - 18) + '" width="10" height="10" fill="' + rgba(se.color) + '"/><text x="' + (lx2 + 14) + '" y="' + (H - 9) + '" font-size="12" fill="#404040" font-family="' + escH(font) + '">' + escH(se.name) + "</text>";
          lx2 += 24 + String(se.name).length * 7;
        });
      }
      return '<svg style="position:absolute;left:' + fmt(box.x) + "px;top:" + fmt(box.y) + 'px" width="' + fmt(W) + '" height="' + fmt(H) + '" viewBox="0 0 ' + fmt(W) + " " + fmt(H) + '">' + out + "</svg>";
    }

    O.pptxToPdf = function (buf, onProgress) {
      return pptxPages(buf, onProgress).then(function (pages) {
        return pagesToPdf(pages, onProgress, 0.3, 1).then(function (blob) { pages.cleanup(); return { blob: blob, pages: pages.length, notes: pages.notes }; },
          function (e) { pages.cleanup(); throw e; });
      });
    };
    O.pptxToImages = function (buf, onProgress, type, quality) {
      return pptxPages(buf, onProgress).then(function (pages) {
        return pagesToImages(pages, onProgress, 0.3, 1, type, quality).then(function (r) { pages.cleanup(); r.notes = pages.notes; return r; }, function (e) { pages.cleanup(); throw e; });
      });
    };

    O._docxPages = docxPages; O._pptxPages = pptxPages; O._xlsxPages = xlsxPages;
    O._internals = { encodePage: encodePage, fontStack: fontStack, textLayer: textLayer, rasterize: rasterize, makeHost: makeHost, settle: settle, ensureFontFaces: ensureFontFaces, pagesToPdf: pagesToPdf, pagesToImages: pagesToImages, PX2PT: PX2PT };
    return O;
  };
})();
