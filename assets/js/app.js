(function () {
  "use strict";

  /* ===========================================================
     0. DATOS DE MARCA Y AJUSTES (el dueño puede editarlos aquí)
     =========================================================== */
  window.__BRAND__ = {
    name: "ConvertirPDF",
    slug: "convertirpdf",
    ads: {
      modalOnDownload: true,      // Pop-up publicitario tras pulsar «Descargar»
      modalDelayMs: 350,          // Se abre DESPUÉS de iniciar la descarga (nunca la bloquea)
      modalCooldownSec: 0,        // 0 = en cada descarga; p. ej. 60 = como máximo 1 vez por minuto
      cornerDelayMs: 7000         // La notificación de esquina aparece pasados 7 s
    },
    limits: {                     // Límites honestos (memoria del navegador)
      pdf: 150 * 1024 * 1024,
      office: 60 * 1024 * 1024,
      image: 60 * 1024 * 1024,
      maxSheetRows: 20000         // Filas por hoja al pasar Excel a PDF/Word
    },
    concurrency: 2                // Archivos que se convierten a la vez
  };
  var BRAND = window.__BRAND__;
  /* Carpeta de motores: se calcula desde la ubicación de este script
     (assets/js/app.js → ../../lib/vendor/), así funciona en cualquier subcarpeta. */
  var VENDOR = (function () {
    try {
      var s = document.currentScript && document.currentScript.src;
      if (s) return new URL("../../lib/vendor/", s).href;
    } catch (e) { /* sin importancia */ }
    return new URL("lib/vendor/", document.baseURI).href;
  })();
  var BUILD = "20261017";          // Cambia este número en cada publicación (evita versiones viejas en caché)

  /* ===========================================================
     1. UTILIDADES GENERALES
     =========================================================== */
  var $ = function (sel, scope) { return (scope || document).querySelector(sel); };
  var $$ = function (sel, scope) { return Array.prototype.slice.call((scope || document).querySelectorAll(sel)); };
  var escHTML = function (s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  };
  function safe(fn, name) { try { fn(); } catch (e) { console.warn("[" + name + "]", e); } }

  function fmtBytes(n) {
    if (!isFinite(n)) return "—";
    if (n < 1024) return n + " B";
    var u = ["KB", "MB", "GB"], i = -1;
    do { n /= 1024; i++; } while (n >= 1024 && i < u.length - 1);
    return n.toLocaleString("es", { maximumFractionDigits: n < 10 ? 1 : 0 }) + " " + u[i];
  }
  function baseName(name) { return String(name).replace(/\.[^.]+$/, "") || "archivo"; }
  function absURL(path) { return new URL(path, document.baseURI).href; }

  /* --- Control del hilo de la interfaz -----------------------------------
     Las conversiones pesadas se trocean: cada cierto tiempo cedemos el turno
     al navegador con un setTimeout(0) para que pinte la barra de progreso y
     responda a clics. Así nunca aparece «La página no responde». */
  function yieldUI() { return new Promise(function (r) { setTimeout(r, 0); }); }
  function makeSlicer(budgetMs) {
    var last = performance.now();
    return function tick() {
      if (performance.now() - last < (budgetMs || 14)) return Promise.resolve();
      return yieldUI().then(function () { last = performance.now(); });
    };
  }

  /* Error pensado para el usuario: el mensaje se muestra tal cual en el aviso. */
  function UserError(title, detail, code) {
    var e = new Error(title);
    e.isUser = true; e.title = title; e.detail = detail || ""; e.code = code || "";
    return e;
  }

  /* Traduce cualquier error técnico a un mensaje humano, sin jerga. */
  function humanize(err, kind) {
    if (err && err.isUser) return err;
    var msg = String((err && (err.message || err.name)) || err || "");
    var name = err && err.name;
    if (name === "PasswordException") {
      return err.code === 2
        ? UserError("Contraseña incorrecta", "Revisa la clave del PDF y vuelve a intentarlo.", "password")
        : UserError("PDF protegido con contraseña", "Escribe la contraseña en la fila para desbloquearlo (no sale de tu equipo).", "password");
    }
    if (name === "InvalidPDFException" || /Invalid PDF|PDF header/i.test(msg)) {
      return UserError("El PDF está dañado", "No se pudo leer su estructura. Prueba a abrirlo y guardarlo de nuevo.");
    }
    if (/end of central directory|zip|Corrupted|Can't find|Could not find file/i.test(msg)) {
      return UserError("El archivo está dañado o no es un " + (kind ? kind.toUpperCase() : "documento") + " válido",
        "No se pudo abrir su contenido. Ábrelo en su programa original y guárdalo de nuevo.");
    }
    if (/memory|allocation|RangeError|too large|Invalid array length/i.test(msg) || name === "RangeError") {
      return UserError("El archivo es demasiado grande para tu navegador", "Prueba con un archivo más pequeño o cierra otras pestañas.");
    }
    if (/Failed to fetch|NetworkError|Importing a module|dynamically imported|load/i.test(msg)) {
      return UserError("No se pudo preparar el motor de conversión", "Revisa tu conexión y recarga la página (solo se descarga la primera vez).");
    }
    if (/decode|image/i.test(msg)) {
      return UserError("No se pudo leer la imagen", "Puede estar dañada o usar una variante poco común del formato.");
    }
    return UserError("No se pudo convertir este archivo", "Puede estar dañado o tener un contenido poco habitual.");
  }

  /* Descarga un Blob con un enlace temporal (el archivo se guarda de fondo).
     Si la página se abre dentro de Claude (vista previa), se usa su ventana
     de guardado, porque allí los enlaces de descarga están bloqueados. */
  function saveBlob(blob, name) {
    var host = window.claude && typeof window.claude.use === "function" ? window.claude : null;
    if (host) {
      return host.use("downloads").then(function (dl) {
        if (!dl) return anchorSave(blob, name);
        return dl.save({ filename: name, data: blob }).catch(function (e) {
          if (e && e.code !== "declined") toast("error", "No se pudo guardar «" + name + "»", "Inténtalo de nuevo en unos segundos.");
        });
      }).catch(function () { anchorSave(blob, name); });
    }
    anchorSave(blob, name);
  }
  function anchorSave(blob, name) {
    var a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = name;
    a.rel = "noopener";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(function () { URL.revokeObjectURL(a.href); }, 8000);
  }

  function readAsArrayBuffer(blob) {
    if (blob.arrayBuffer) return blob.arrayBuffer();
    return new Promise(function (ok, ko) {
      var r = new FileReader();
      r.onload = function () { ok(r.result); };
      r.onerror = function () { ko(r.error); };
      r.readAsArrayBuffer(blob);
    });
  }
  function bytesToBase64(u8) {
    var s = "", CH = 0x8000;
    for (var i = 0; i < u8.length; i += CH) s += String.fromCharCode.apply(null, u8.subarray(i, i + CH));
    return btoa(s);
  }
  function dataURLToBytes(url) {
    var b64 = url.split(",")[1] || "";
    var bin = atob(b64), u8 = new Uint8Array(bin.length);
    for (var i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
    return u8;
  }

  /* XML: escapar texto y quitar caracteres de control que rompen Word/PowerPoint. */
  function xmlEsc(s) {
    return String(s == null ? "" : s)
      .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F￾￿]/g, "")
      .replace(/[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(^|[^\uD800-\uDBFF])[\uDC00-\uDFFF]/g, "$1")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }

  /* ===========================================================
     2. CARGADOR DE MOTORES (bajo demanda, una sola vez)
     -----------------------------------------------------------
     Nada pesado se descarga al abrir la página. Cada motor se pide la
     primera vez que una conversión lo necesita y queda en memoria.
     Versiones fijadas (lib/vendor/):
       pdf.js 6.1.200 · jsPDF 4.2.1 · jspdf-autotable 5.0.8 ·
       JSZip 3.10.1 · mammoth 1.12.3 · SheetJS 0.20.3
     =========================================================== */
  var enginePromises = {};
  function loadScript(src) {
    return new Promise(function (ok, ko) {
      var full = absURL(src) + "?v=" + BUILD;
      var prev = document.querySelector('script[data-engine="' + src + '"]');
      if (prev && prev.dataset.loaded) return ok();
      var s = document.createElement("script");
      s.src = full; s.async = true; s.dataset.engine = src;
      s.onload = function () { s.dataset.loaded = "1"; ok(); };
      s.onerror = function () { s.remove(); ko(new Error("No se pudo cargar " + src)); };
      document.head.appendChild(s);
    });
  }
  function once(key, factory) {
    if (!enginePromises[key]) {
      enginePromises[key] = factory().catch(function (e) { delete enginePromises[key]; throw e; });
    }
    return enginePromises[key];
  }

  var Engines = {
    jszip: function () {
      return once("jszip", function () {
        return loadScript(VENDOR + "jszip.min.js").then(function () {
          if (!window.JSZip) throw new Error("load JSZip");
          return window.JSZip;
        });
      });
    },
    jspdf: function () {
      return once("jspdf", function () {
        return loadScript(VENDOR + "jspdf.umd.min.js")
          .then(function () { return loadScript(VENDOR + "jspdf.plugin.autotable.min.js"); })
          .then(function () {
            if (!window.jspdf || !window.jspdf.jsPDF) throw new Error("load jsPDF");
            return window.jspdf.jsPDF;
          });
      });
    },
    mammoth: function () {
      return once("mammoth", function () {
        return loadScript(VENDOR + "mammoth.browser.min.js").then(function () {
          if (!window.mammoth) throw new Error("load mammoth");
          return window.mammoth;
        });
      });
    },
    pdflib: function () {
      return once("pdflib", function () {
        return loadScript(VENDOR + "pdf-lib.min.js").then(function () {
          if (!window.PDFLib || !window.PDFLib.PDFDocument) throw new Error("load PDFLib");
          return window.PDFLib;
        });
      });
    },
    xlsx: function () {
      return once("xlsx", function () {
        return loadScript(VENDOR + "xlsx.core.min.js").then(function () {
          if (!window.XLSX) throw new Error("load XLSX");
          return window.XLSX;
        });
      });
    },
    /* pdf.js es un módulo moderno: se importa dinámicamente desde este script
       clásico (el «puente» de módulos). Necesita que la web se sirva por http(s). */
    pdfjs: function () {
      return once("pdfjs", function () {
        return import(absURL(VENDOR + "pdfjs/pdf.min.mjs") + "?v=" + BUILD).then(function (lib) {
          lib.GlobalWorkerOptions.workerSrc = absURL(VENDOR + "pdfjs/pdf.worker.min.mjs") + "?v=" + BUILD;
          return lib;
        });
      });
    },
    /* Fuente Unicode (Liberation Sans, incluida con pdf.js) para textos con
       caracteres que la fuente estándar del PDF no tiene (griego, cirílico…). */
    unicodeFont: function () {
      return once("unicodeFont", function () {
        var files = { normal: "LiberationSans-Regular.ttf", bold: "LiberationSans-Bold.ttf",
          italic: "LiberationSans-Italic.ttf", bolditalic: "LiberationSans-BoldItalic.ttf" };
        var keys = Object.keys(files);
        return Promise.all(keys.map(function (k) {
          return fetch(absURL(VENDOR + "pdfjs/standard_fonts/" + files[k]))
            .then(function (r) { if (!r.ok) throw new Error("font"); return r.arrayBuffer(); })
            .then(function (b) { return [k, files[k], bytesToBase64(new Uint8Array(b))]; });
        }));
      });
    }
  };

  function openPdf(bytes, password) {
    return Engines.pdfjs().then(function (pdfjsLib) {
      var task = pdfjsLib.getDocument({
        data: bytes,
        password: password || undefined,
        cMapUrl: absURL(VENDOR + "pdfjs/cmaps/"),
        cMapPacked: true,
        standardFontDataUrl: absURL(VENDOR + "pdfjs/standard_fonts/"),
        wasmUrl: absURL(VENDOR + "pdfjs/wasm/"),
        isEvalSupported: false
      });
      return task.promise.then(function (pdf) { pdfTasks.set(pdf, task); return pdf; });
    });
  }
  /* Libera la memoria del PDF abierto (el motor guarda la tarea de carga). */
  var pdfTasks = new WeakMap();
  function closePdf(pdf) {
    try {
      if (pdf && typeof pdf.destroy === "function") pdf.destroy();
      else if (pdfTasks.has(pdf)) pdfTasks.get(pdf).destroy();
    } catch (e) { /* sin importancia */ }
  }

  /* ===========================================================
     3. DETECCIÓN DE TIPO (por contenido real, no solo por extensión)
     =========================================================== */
  var KIND_LABEL = { pdf: "PDF", docx: "WORD", xlsx: "EXCEL", pptx: "PPT", png: "PNG", jpg: "JPG", webp: "WEBP" };
  var EXT_TO_KIND = { pdf: "pdf", docx: "docx", xlsx: "xlsx", pptx: "pptx", png: "png", jpg: "jpg", jpeg: "jpg", webp: "webp" };
  var IMAGE_KINDS = { png: 1, jpg: 1, webp: 1 };

  function detectKind(file) {
    var ext = (String(file.name).split(".").pop() || "").toLowerCase();
    var byExt = EXT_TO_KIND[ext] || null;
    return readAsArrayBuffer(file.slice(0, 16)).then(function (buf) {
      var b = new Uint8Array(buf);
      var ascii = String.fromCharCode.apply(null, b);
      var magic = null;
      if (ascii.indexOf("%PDF") === 0) magic = "pdf";
      else if (b[0] === 0x89 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) magic = "png";
      else if (b[0] === 0xFF && b[1] === 0xD8 && b[2] === 0xFF) magic = "jpg";
      else if (ascii.indexOf("RIFF") === 0 && ascii.slice(8, 12) === "WEBP") magic = "webp";
      else if (b[0] === 0x50 && b[1] === 0x4B) magic = "zip";
      else if (b[0] === 0xD0 && b[1] === 0xCF && b[2] === 0x11 && b[3] === 0xE0) magic = "cfb";
      else if (/^\s*%PDF/.test(ascii)) magic = "pdf";

      if (magic === "zip") {
        if (byExt === "docx" || byExt === "xlsx" || byExt === "pptx") return { kind: byExt };
        return { kind: null, reason: "Este ZIP no es un documento de Office compatible." };
      }
      if (magic === "cfb") {
        // Un .docx/.xlsx/.pptx con contraseña se guarda en un contenedor cifrado (o es un .doc/.xls antiguo).
        return { kind: null, encrypted: true, reason: "Está protegido con contraseña o es un formato antiguo de Office (.doc/.xls/.ppt). Ábrelo en su programa y guárdalo sin contraseña como .docx/.xlsx/.pptx." };
      }
      if (magic) return { kind: magic };
      if (byExt === "pdf") return { kind: null, reason: "No parece un PDF válido (puede estar dañado)." };
      if (byExt) return { kind: null, reason: "El contenido no coincide con su extensión ." + ext + " (puede estar dañado)." };
      return { kind: null, reason: "Formato no compatible." };
    });
  }

  /* ===========================================================
     4. MATRIZ DE COMPATIBILIDAD (solo salidas técnicamente posibles)
     =========================================================== */
  var TARGET_LABEL = {
    pdf: "PDF", docx: "Word (.docx)", xlsx: "Excel (.xlsx)", pptx: "PowerPoint (.pptx)",
    png: "PNG", jpg: "JPG", webp: "WebP"
  };
  var TARGETS = {
    pdf: ["docx", "xlsx", "pptx", "png", "jpg", "webp"],
    docx: ["pdf", "xlsx", "pptx", "png", "jpg", "webp"],
    xlsx: ["pdf", "docx", "pptx", "png", "jpg", "webp"],
    pptx: ["pdf", "docx", "png", "jpg", "webp"],
    png: ["jpg", "webp", "pdf", "docx", "pptx"],
    jpg: ["png", "webp", "pdf", "docx", "pptx"],
    webp: ["png", "jpg", "pdf", "docx", "pptx"]
  };
  var DEFAULT_TARGET = { pdf: "docx", docx: "pdf", xlsx: "pdf", pptx: "pdf", png: "jpg", jpg: "png", webp: "jpg" };
  var MIME = {
    pdf: "application/pdf",
    docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    png: "image/png", jpg: "image/jpeg", webp: "image/webp", zip: "application/zip"
  };

  /* Safari no sabe EXPORTAR WebP: si el navegador no puede, quitamos la opción. */
  var canExportWebP = (function () {
    try {
      var c = document.createElement("canvas"); c.width = c.height = 2;
      return c.toDataURL("image/webp").indexOf("data:image/webp") === 0;
    } catch (e) { return false; }
  })();
  function targetsFor(kind) {
    return (TARGETS[kind] || []).filter(function (t) { return t !== "webp" || canExportWebP; });
  }

  /* ===========================================================
     5. IMÁGENES (Canvas nativo de HTML5)
     =========================================================== */
  var MAX_IMAGE_AREA = 40e6;     // ~40 megapíxeles: por encima se reduce (límite de memoria)
  var MAX_CANVAS_SIDE = 16384;

  /* Decodifica un Blob de imagen respetando la orientación EXIF. */
  function decodeImage(blob) {
    if (window.createImageBitmap) {
      return createImageBitmap(blob, { imageOrientation: "from-image" }).catch(function () { return decodeWithImg(blob); });
    }
    return decodeWithImg(blob);
  }
  function decodeWithImg(blob) {
    return new Promise(function (ok, ko) {
      var url = URL.createObjectURL(blob), img = new Image();
      img.onload = function () { ok(img); setTimeout(function () { URL.revokeObjectURL(url); }, 1000); };
      img.onerror = function () { URL.revokeObjectURL(url); ko(new Error("image decode")); };
      img.src = url;
    });
  }
  function fitSize(w, h, maxSide, maxArea) {
    var s = 1;
    if (maxSide && Math.max(w, h) > maxSide) s = Math.min(s, maxSide / Math.max(w, h));
    if (maxArea && w * h * s * s > maxArea) s = Math.min(s, Math.sqrt(maxArea / (w * h)));
    return { w: Math.max(1, Math.round(w * s)), h: Math.max(1, Math.round(h * s)), scaled: s < 1 };
  }
  function canvasToBlob(canvas, mime, quality) {
    return new Promise(function (ok, ko) {
      if (canvas.toBlob) {
        canvas.toBlob(function (b) { b ? ok(b) : ko(new Error("image encode")); }, mime, quality);
      } else {
        try { var u = canvas.toDataURL(mime, quality); ok(new Blob([dataURLToBytes(u)], { type: u.slice(5, u.indexOf(";")) })); }
        catch (e) { ko(e); }
      }
    });
  }
  /* Dibuja la imagen en un lienzo; con fondo blanco si el destino no admite transparencia. */
  function imageToCanvas(img, opts) {
    opts = opts || {};
    var iw = img.width || img.naturalWidth, ih = img.height || img.naturalHeight;
    var size = fitSize(iw, ih, opts.maxSide || MAX_CANVAS_SIDE, opts.maxArea || MAX_IMAGE_AREA);
    var c = document.createElement("canvas");
    c.width = size.w; c.height = size.h;
    var ctx = c.getContext("2d");
    if (opts.whiteBg) { ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0, 0, c.width, c.height); }
    ctx.imageSmoothingQuality = "high";
    ctx.drawImage(img, 0, 0, c.width, c.height);
    c._scaled = size.scaled;
    return c;
  }
  /* Normaliza cualquier imagen a PNG o JPEG (lo que entienden PDF, Word y PowerPoint). */
  function normalizeImage(blob, opts) {
    opts = opts || {};
    return decodeImage(blob).then(function (img) {
      var keepAlpha = opts.prefer === "png" || (opts.prefer !== "jpeg" && /png|webp|gif|svg/i.test(blob.type || ""));
      var c = imageToCanvas(img, { maxSide: opts.maxSide || 3000, whiteBg: !keepAlpha });
      if (img.close) img.close();
      var mime = keepAlpha ? "image/png" : "image/jpeg";
      return canvasToBlob(c, mime, 0.9).then(readAsArrayBuffer).then(function (buf) {
        var r = { bytes: new Uint8Array(buf), mime: mime, w: c.width, h: c.height };
        c.width = c.height = 0; // libera memoria del lienzo
        return r;
      });
    });
  }

  /* ===========================================================
     6. ESCRITORES DE OFFICE (Word y PowerPoint generados a mano con JSZip)
     =========================================================== */
  var NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
  var NS_OFFREL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  var XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  function corePropsXml(title) {
    var now = new Date().toISOString().replace(/\.\d+Z$/, "Z");
    return XML_HEAD + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
      "<dc:title>" + xmlEsc(title || "") + "</dc:title><dc:creator>" + xmlEsc(BRAND.name) + "</dc:creator>" +
      '<dcterms:created xsi:type="dcterms:W3CDTF">' + now + '</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">' + now + "</dcterms:modified></cp:coreProperties>";
  }
  function appPropsXml() {
    return XML_HEAD + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>' + xmlEsc(BRAND.name) + "</Application></Properties>";
  }
  function zipToBlob(zip, mime, onProgress) {
    return zip.generateAsync({ type: "blob", mimeType: mime, compression: "DEFLATE", compressionOptions: { level: 6 } },
      function (meta) { if (onProgress) onProgress(meta.percent / 100); });
  }

  /* ---------- 6.1 Word (.docx) a partir del modelo de documento ----------
     Modelo: { title, landscape, blocks: [
       {type:"heading", level:1-3, runs:[{text,bold,italic,underline}]},
       {type:"para", runs:[...], align?}, {type:"list", ordered, items:[{runs, level}]},
       {type:"table", rows:[["celda", ...]], header:true}, {type:"image", blob|img}, {type:"pagebreak"} ] } */
  var EMU_PER_INCH = 914400, TWIP_PER_INCH = 1440;

  function runsXml(runs, extraRPr) {
    return (runs || []).map(function (r) {
      var rPr = (r.bold ? "<w:b/>" : "") + (r.italic ? "<w:i/>" : "") + (extraRPr || "") + (r.underline ? '<w:u w:val="single"/>' : "");
      var parts = String(r.text == null ? "" : r.text).split(/(\n|\t)/);
      var inner = parts.map(function (p) {
        if (p === "\n") return "<w:br/>";
        if (p === "\t") return "<w:tab/>";
        return p ? '<w:t xml:space="preserve">' + xmlEsc(p) + "</w:t>" : "";
      }).join("");
      return inner ? "<w:r>" + (rPr ? "<w:rPr>" + rPr + "</w:rPr>" : "") + inner + "</w:r>" : "";
    }).join("");
  }

  function buildDocx(model, onProgress) {
    return Engines.jszip().then(function (JSZip) {
      var zip = new JSZip();
      var rels = [], media = [], body = [], imgId = 0;
      var tick = makeSlicer();
      var landscape = !!model.landscape;
      var pageW = landscape ? 16838 : 11906, pageH = landscape ? 11906 : 16838, margin = 1134;
      var contentW = pageW - 2 * margin, contentH = pageH - 2 * margin;
      var blocks = model.blocks || [];

      function addImage(norm, full) {
        imgId++;
        var ext = norm.mime === "image/png" ? "png" : "jpeg";
        var file = "image" + imgId + "." + ext, rid = "rIdImg" + imgId;
        media.push({ path: "word/media/" + file, bytes: norm.bytes });
        rels.push('<Relationship Id="' + rid + '" Type="' + NS_OFFREL + '/image" Target="media/' + file + '"/>');
        // Ajuste al ancho/alto útil de la página manteniendo proporción
        var maxWIn = contentW / TWIP_PER_INCH * (full ? 1 : 0.75), maxHIn = contentH / TWIP_PER_INCH * (full ? 0.96 : 0.5);
        var wIn = norm.w / 96, hIn = norm.h / 96;
        var s = Math.min(1, maxWIn / wIn, maxHIn / hIn);
        var cx = Math.round(wIn * s * EMU_PER_INCH), cy = Math.round(hIn * s * EMU_PER_INCH);
        return '<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
          '<wp:extent cx="' + cx + '" cy="' + cy + '"/><wp:docPr id="' + imgId + '" name="Imagen ' + imgId + '"/>' +
          '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>' +
          '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic>' +
          '<pic:nvPicPr><pic:cNvPr id="' + imgId + '" name="' + file + '"/><pic:cNvPicPr/></pic:nvPicPr>' +
          '<pic:blipFill><a:blip r:embed="' + rid + '"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>' +
          '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>' +
          "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>";
      }

      function tableXml(b) {
        var rows = b.rows || [];
        var cols = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0) || 1;
        var colW = Math.floor(contentW / cols);
        var small = cols > 6 ? '<w:sz w:val="16"/><w:szCs w:val="16"/>' : "";
        var border = 'w:val="single" w:sz="4" w:space="0" w:color="BFBAB0"';
        var grid = "";
        for (var c = 0; c < cols; c++) grid += '<w:gridCol w:w="' + colW + '"/>';
        var trs = rows.map(function (r, ri) {
          var head = b.header && ri === 0;
          var tds = "";
          for (var ci = 0; ci < cols; ci++) {
            var txt = r[ci] == null ? "" : String(r[ci]);
            tds += '<w:tc><w:tcPr><w:tcW w:w="' + colW + '" w:type="dxa"/>' + (head ? '<w:shd w:val="clear" w:color="auto" w:fill="EDEBE6"/>' : "") +
              '</w:tcPr><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr>' +
              runsXml([{ text: txt, bold: head }], small) + "</w:p></w:tc>";
          }
          return "<w:tr>" + (head ? "<w:trPr><w:tblHeader/></w:trPr>" : "") + tds + "</w:tr>";
        }).join("");
        return '<w:tbl><w:tblPr><w:tblW w:w="' + (colW * cols) + '" w:type="dxa"/><w:tblBorders>' +
          "<w:top " + border + "/><w:left " + border + "/><w:bottom " + border + "/><w:right " + border + "/>" +
          "<w:insideH " + border + "/><w:insideV " + border + '/></w:tblBorders><w:tblLayout w:type="fixed"/></w:tblPr>' +
          "<w:tblGrid>" + grid + "</w:tblGrid>" + trs + "</w:tbl><w:p/>";
      }

      var i = 0;
      function step() {
        if (i >= blocks.length) return Promise.resolve();
        var b = blocks[i++];
        if (onProgress) onProgress(0.85 * i / blocks.length);
        var p;
        if (b.type === "heading") {
          body.push('<w:p><w:pPr><w:pStyle w:val="Heading' + Math.min(3, Math.max(1, b.level || 1)) + '"/></w:pPr>' + runsXml(b.runs) + "</w:p>");
        } else if (b.type === "para") {
          var jc = b.align === "center" ? '<w:jc w:val="center"/>' : b.align === "right" ? '<w:jc w:val="right"/>' : "";
          body.push("<w:p>" + (jc ? "<w:pPr>" + jc + "</w:pPr>" : "") + runsXml(b.runs) + "</w:p>");
        } else if (b.type === "list") {
          (b.items || []).forEach(function (it, n) {
            var lvl = it.level || 0, left = 720 + lvl * 360;
            var marker = b.ordered ? (n + 1) + ".\t" : "•\t";
            body.push('<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="' + left + '"/></w:tabs><w:spacing w:after="60"/><w:ind w:left="' + left + '" w:hanging="360"/></w:pPr>' +
              runsXml([{ text: marker }].concat(it.runs || [])) + "</w:p>");
          });
        } else if (b.type === "table") {
          body.push(tableXml(b));
        } else if (b.type === "pagebreak") {
          body.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
        } else if (b.type === "image") {
          p = (b.norm ? Promise.resolve(b.norm) : normalizeImage(b.blob, { maxSide: 2400 })).then(function (norm) {
            body.push(addImage(norm, b.full));
          }).catch(function () {
            body.push("<w:p>" + runsXml([{ text: "[Imagen no compatible omitida]", italic: true }]) + "</w:p>");
          });
        }
        return (p || Promise.resolve()).then(tick).then(step);
      }

      return step().then(function () {
        if (!body.length) body.push("<w:p/>");
        var sect = '<w:sectPr><w:pgSz w:w="' + pageW + '" w:h="' + pageH + '"' + (landscape ? ' w:orient="landscape"' : "") + "/>" +
          '<w:pgMar w:top="' + margin + '" w:right="' + margin + '" w:bottom="' + margin + '" w:left="' + margin + '" w:header="708" w:footer="708" w:gutter="0"/></w:sectPr>';
        var doc = XML_HEAD + '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="' + NS_OFFREL + '" ' +
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>' + body.join("") + sect + "</w:body></w:document>";

        var hStyle = function (n, sz, before) {
          return '<w:style w:type="paragraph" w:styleId="Heading' + n + '"><w:name w:val="heading ' + n + '"/><w:basedOn w:val="Normal"/><w:next w:val="Normal"/><w:uiPriority w:val="9"/><w:qFormat/>' +
            '<w:pPr><w:keepNext/><w:spacing w:before="' + before + '" w:after="120"/><w:outlineLvl w:val="' + (n - 1) + '"/></w:pPr><w:rPr><w:b/><w:color w:val="1F2430"/><w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/></w:rPr></w:style>';
        };
        var styles = XML_HEAD + '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
          '<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Calibri" w:cs="Calibri"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:val="es-ES"/></w:rPr></w:rPrDefault>' +
          '<w:pPrDefault><w:pPr><w:spacing w:after="160" w:line="276" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>' +
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style>' +
          hStyle(1, 36, 360) + hStyle(2, 30, 280) + hStyle(3, 26, 240) +
          '<w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:uiPriority w:val="99"/><w:semiHidden/><w:tblPr><w:tblInd w:w="0" w:type="dxa"/>' +
          '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style></w:styles>';

        zip.file("[Content_Types].xml", XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
          '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
          '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
          '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>');
        zip.file("_rels/.rels", XML_HEAD + '<Relationships xmlns="' + NS_REL + '">' +
          '<Relationship Id="rId1" Type="' + NS_OFFREL + '/officeDocument" Target="word/document.xml"/>' +
          '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
          '<Relationship Id="rId3" Type="' + NS_OFFREL + '/extended-properties" Target="docProps/app.xml"/></Relationships>');
        zip.file("word/_rels/document.xml.rels", XML_HEAD + '<Relationships xmlns="' + NS_REL + '">' +
          '<Relationship Id="rIdStyles" Type="' + NS_OFFREL + '/styles" Target="styles.xml"/>' + rels.join("") + "</Relationships>");
        zip.file("word/document.xml", doc);
        zip.file("word/styles.xml", styles);
        zip.file("docProps/core.xml", corePropsXml(model.title));
        zip.file("docProps/app.xml", appPropsXml());
        media.forEach(function (m) { zip.file(m.path, m.bytes, { compression: "STORE" }); });
        return zipToBlob(zip, MIME.docx, function (f) { if (onProgress) onProgress(0.85 + 0.15 * f); });
      });
    });
  }

  /* ---------- 6.2 PowerPoint (.pptx): una imagen por diapositiva ----------
     slides: [{ norm:{bytes,mime,w,h} }], tamaño de diapositiva según la 1.ª imagen. */
  function buildPptx(slides, title, onProgress) {
    return Engines.jszip().then(function (JSZip) {
      var zip = new JSZip();
      var P = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="' + NS_OFFREL + '" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';
      var first = slides[0] && slides[0].norm;
      var cx = 12192000, cy = 6858000; // 16:9 por defecto
      if (first && first.w && first.h) {
        cy = Math.round(cx * first.h / first.w);
        if (cy > 51206400) { cy = 51206400; cx = Math.round(cy * first.w / first.h); }
        if (cy < 914400) cy = 914400;
      }
      var emptyTree = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>';
      var ct = [], presRels = [], sldIds = [];

      slides.forEach(function (s, idx) {
        var n = idx + 1, norm = s.norm;
        var ext = norm.mime === "image/png" ? "png" : "jpeg";
        // Ajuste de la imagen centrada dentro de la diapositiva
        var sc = Math.min(cx / norm.w, cy / norm.h);
        var w = Math.round(norm.w * sc), h = Math.round(norm.h * sc);
        var x = Math.round((cx - w) / 2), y = Math.round((cy - h) / 2);
        zip.file("ppt/media/image" + n + "." + ext, norm.bytes, { compression: "STORE" });
        zip.file("ppt/slides/slide" + n + ".xml", XML_HEAD + "<p:sld " + P + "><p:cSld><p:spTree>" + emptyTree +
          '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Imagen ' + n + '"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rId2"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          '<p:spPr><a:xfrm><a:off x="' + x + '" y="' + y + '"/><a:ext cx="' + w + '" cy="' + h + '"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>' +
          "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>");
        zip.file("ppt/slides/_rels/slide" + n + ".xml.rels", XML_HEAD + '<Relationships xmlns="' + NS_REL + '">' +
          '<Relationship Id="rId1" Type="' + NS_OFFREL + '/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
          '<Relationship Id="rId2" Type="' + NS_OFFREL + '/image" Target="../media/image' + n + "." + ext + '"/></Relationships>');
        ct.push('<Override PartName="/ppt/slides/slide' + n + '.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>');
        presRels.push('<Relationship Id="rIdS' + n + '" Type="' + NS_OFFREL + '/slide" Target="slides/slide' + n + '.xml"/>');
        sldIds.push('<p:sldId id="' + (255 + n) + '" r:id="rIdS' + n + '"/>');
      });

      zip.file("[Content_Types].xml", XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>' +
        '<Default Extension="png" ContentType="image/png"/><Default Extension="jpeg" ContentType="image/jpeg"/>' +
        '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>' +
        '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>' +
        '<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>' +
        '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
        '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>' +
        '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>' +
        '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
        '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
        ct.join("") + "</Types>");
      zip.file("_rels/.rels", XML_HEAD + '<Relationships xmlns="' + NS_REL + '">' +
        '<Relationship Id="rId1" Type="' + NS_OFFREL + '/officeDocument" Target="ppt/presentation.xml"/>' +
        '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
        '<Relationship Id="rId3" Type="' + NS_OFFREL + '/extended-properties" Target="docProps/app.xml"/></Relationships>');
      zip.file("ppt/presentation.xml", XML_HEAD + "<p:presentation " + P + ' saveSubsetFonts="1">' +
        '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdM"/></p:sldMasterIdLst>' +
        "<p:sldIdLst>" + sldIds.join("") + "</p:sldIdLst>" +
        '<p:sldSz cx="' + cx + '" cy="' + cy + '"/><p:notesSz cx="6858000" cy="9144000"/></p:presentation>');
      zip.file("ppt/_rels/presentation.xml.rels", XML_HEAD + '<Relationships xmlns="' + NS_REL + '">' +
        '<Relationship Id="rIdM" Type="' + NS_OFFREL + '/slideMaster" Target="slideMasters/slideMaster1.xml"/>' +
        '<Relationship Id="rIdT" Type="' + NS_OFFREL + '/theme" Target="theme/theme1.xml"/>' +
        '<Relationship Id="rIdP" Type="' + NS_OFFREL + '/presProps" Target="presProps.xml"/>' +
        '<Relationship Id="rIdTS" Type="' + NS_OFFREL + '/tableStyles" Target="tableStyles.xml"/>' +
        presRels.join("") + "</Relationships>");
      zip.file("ppt/presProps.xml", XML_HEAD + "<p:presentationPr " + P + "/>");
      zip.file("ppt/tableStyles.xml", XML_HEAD + '<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>');
      zip.file("ppt/slideMasters/slideMaster1.xml", XML_HEAD + "<p:sldMaster " + P + ">" +
        '<p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>' + emptyTree + "</p:spTree></p:cSld>" +
        '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>' +
        '<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>' +
        "<p:txStyles><p:titleStyle/><p:bodyStyle/><p:otherStyle/></p:txStyles></p:sldMaster>");
      zip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", XML_HEAD + '<Relationships xmlns="' + NS_REL + '">' +
        '<Relationship Id="rId1" Type="' + NS_OFFREL + '/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>' +
        '<Relationship Id="rId2" Type="' + NS_OFFREL + '/theme" Target="../theme/theme1.xml"/></Relationships>');
      zip.file("ppt/slideLayouts/slideLayout1.xml", XML_HEAD + "<p:sldLayout " + P + ' type="blank" preserve="1">' +
        '<p:cSld name="En blanco"><p:spTree>' + emptyTree + "</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>");
      zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", XML_HEAD + '<Relationships xmlns="' + NS_REL + '">' +
        '<Relationship Id="rId1" Type="' + NS_OFFREL + '/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>');
      var clr = function (tag, hex) { return "<a:" + tag + '><a:srgbClr val="' + hex + '"/></a:' + tag + ">"; };
      var fill3 = '<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>';
      zip.file("ppt/theme/theme1.xml", XML_HEAD + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="ConvertirPDF"><a:themeElements>' +
        '<a:clrScheme name="ConvertirPDF"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>' +
        clr("dk2", "44546A") + clr("lt2", "E7E6E6") + clr("accent1", "4472C4") + clr("accent2", "ED7D31") + clr("accent3", "A5A5A5") +
        clr("accent4", "FFC000") + clr("accent5", "5B9BD5") + clr("accent6", "70AD47") + clr("hlink", "0563C1") + clr("folHlink", "954F72") + "</a:clrScheme>" +
        '<a:fontScheme name="ConvertirPDF"><a:majorFont><a:latin typeface="Calibri Light"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>' +
        '<a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme>' +
        '<a:fmtScheme name="ConvertirPDF"><a:fillStyleLst>' + fill3 + fill3 + fill3 + "</a:fillStyleLst>" +
        '<a:lnStyleLst><a:ln w="6350">' + fill3 + '</a:ln><a:ln w="12700">' + fill3 + '</a:ln><a:ln w="19050">' + fill3 + "</a:ln></a:lnStyleLst>" +
        "<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>" +
        "<a:bgFillStyleLst>" + fill3 + fill3 + fill3 + "</a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>");
      zip.file("docProps/core.xml", corePropsXml(title));
      zip.file("docProps/app.xml", appPropsXml());
      return zipToBlob(zip, MIME.pptx, onProgress);
    });
  }

  /* ---------- 6.3 Excel (.xlsx) con SheetJS ---------- */
  function smartValue(v) {
    if (typeof v !== "string") return v;
    var s = v.trim();
    if (/^-?\d{1,15}$/.test(s) && !/^0\d/.test(s)) return Number(s);
    if (/^-?\d+[.,]\d+$/.test(s) && !/^-?\d{1,3}[.,]\d{3}$/.test(s)) return Number(s.replace(",", "."));
    return v;
  }
  function safeSheetName(name, used) {
    var base = String(name || "Hoja").replace(/[\[\]:*?\/\\]/g, " ").trim().slice(0, 31) || "Hoja";
    var n = base, k = 2;
    while (used[n.toLowerCase()]) { n = base.slice(0, 28) + " " + k++; }
    used[n.toLowerCase()] = 1;
    return n;
  }
  function buildXlsx(sheets, onProgress) {
    return Engines.xlsx().then(function (XLSX) {
      var wb = XLSX.utils.book_new(), used = {}, tick = makeSlicer();
      var i = 0;
      function step() {
        if (i >= sheets.length) return Promise.resolve();
        var sh = sheets[i++];
        var rows = (sh.rows && sh.rows.length ? sh.rows : [[""]]).map(function (r) { return r.map(smartValue); });
        var ws = XLSX.utils.aoa_to_sheet(rows);
        var widths = [];
        rows.forEach(function (r) { r.forEach(function (v, c) { widths[c] = Math.max(widths[c] || 6, Math.min(60, String(v == null ? "" : v).length + 2)); }); });
        ws["!cols"] = widths.map(function (w) { return { wch: w }; });
        XLSX.utils.book_append_sheet(wb, ws, safeSheetName(sh.name, used));
        if (onProgress) onProgress(0.8 * i / sheets.length);
        return tick().then(step);
      }
      return step().then(yieldUI).then(function () {
        var out = XLSX.write(wb, { type: "array", bookType: "xlsx", compression: true });
        if (onProgress) onProgress(1);
        return new Blob([out], { type: MIME.xlsx });
      });
    });
  }

  /* ===========================================================
     7. GENERADOR DE PDF (jsPDF): texto real y seleccionable, no capturas
     =========================================================== */
  // Caracteres que la fuente estándar del PDF (WinAnsi) sabe dibujar.
  var WINANSI_RE = /[^\x09\x0A\x0D\x20-\x7E\xA0-\xFF€‚ƒ„…†‡ˆ‰Š‹ŒŽ‘’“”•–—˜™š›œžŸ]/;
  var INK = [28, 28, 32];

  function collectText(obj, out) {
    out = out || [];
    if (obj == null) return out;
    if (typeof obj === "string") { out.push(obj); return out; }
    if (Array.isArray(obj)) { obj.forEach(function (o) { collectText(o, out); }); return out; }
    if (typeof obj === "object" && !(obj instanceof Blob) && !(obj instanceof Uint8Array)) {
      Object.keys(obj).forEach(function (k) { if (k !== "blob" && k !== "norm") collectText(obj[k], out); });
    }
    return out;
  }
  function needsUnicodeFont(obj) {
    var all = collectText(obj);
    for (var i = 0; i < all.length; i++) if (WINANSI_RE.test(all[i])) return true;
    return false;
  }
  /* Crea el documento PDF y registra la fuente adecuada. */
  function newPdf(jsPDF, fonts, opts) {
    var doc = new jsPDF({ unit: "pt", format: opts.format || "a4", orientation: opts.orientation || "portrait", compress: true });
    var FONT = "helvetica";
    if (fonts) {
      fonts.forEach(function (f) { doc.addFileToVFS(f[1], f[2]); doc.addFont(f[1], "LiberationSans", f[0]); });
      FONT = "LiberationSans";
    }
    doc.setProperties({ creator: BRAND.name, title: opts.title || "" });
    return { doc: doc, FONT: FONT };
  }
  function pdfEngines(obj) {
    return Promise.all([Engines.jspdf(), needsUnicodeFont(obj) ? Engines.unicodeFont() : null]);
  }
  function styleOf(r) { return r.bold && r.italic ? "bolditalic" : r.bold ? "bold" : r.italic ? "italic" : "normal"; }

  /* Reparte texto enriquecido (negritas, cursivas, tamaños, colores) en líneas. */
  function layoutRuns(doc, FONT, runs, maxW, defSize) {
    var lines = [], cur = { segs: [], w: 0, h: 0 };
    function push() { lines.push(cur); cur = { segs: [], w: 0, h: 0 }; }
    function trimEnd() {
      var last = cur.segs[cur.segs.length - 1];
      if (last && / $/.test(last.text)) {
        doc.setFont(FONT, last.style); doc.setFontSize(last.size);
        var sw = doc.getTextWidth(" ");
        last.text = last.text.replace(/ $/, ""); last.w -= sw; cur.w -= sw;
      }
    }
    function addSeg(text, w, r, size, style) {
      var last = cur.segs[cur.segs.length - 1];
      var color = r.color || null;
      if (last && last.style === style && last.size === size && String(last.color) === String(color) && last.underline === !!r.underline) {
        last.text += text; last.w += w;
      } else {
        cur.segs.push({ text: text, w: w, style: style, size: size, color: color, underline: !!r.underline });
      }
      cur.w += w; cur.h = Math.max(cur.h, size);
    }
    (runs || []).forEach(function (r) {
      var size = r.size || defSize, style = styleOf(r);
      doc.setFont(FONT, style); doc.setFontSize(size);
      String(r.text == null ? "" : r.text).replace(/\r/g, "").replace(/\t/g, "    ").split(/(\n|\s+)/).forEach(function (tok) {
        if (!tok) return;
        if (tok === "\n") { cur.h = cur.h || size; trimEnd(); push(); return; }
        var isSpace = /^\s+$/.test(tok);
        if (isSpace) { if (!cur.segs.length) return; tok = " "; }
        var w = doc.getTextWidth(tok);
        if (!isSpace && cur.w + w > maxW && cur.segs.length) { trimEnd(); push(); }
        if (!isSpace && w > maxW) {           // palabra más larga que la línea: se corta por letras
          var chunk = "";
          Array.from(tok).forEach(function (ch) {
            if (doc.getTextWidth(chunk + ch) > maxW && chunk) { addSeg(chunk, doc.getTextWidth(chunk), r, size, style); push(); chunk = ch; }
            else chunk += ch;
          });
          tok = chunk; w = doc.getTextWidth(tok);
        }
        addSeg(tok, w, r, size, style);
      });
    });
    if (cur.segs.length || !lines.length) { cur.h = cur.h || defSize; trimEnd(); push(); }
    return lines;
  }
  /* Dibuja líneas ya repartidas. Devuelve la coordenada Y final. */
  function drawLines(doc, FONT, lines, x, y, maxW, align, lh, defColor) {
    lines.forEach(function (ln) {
      var h = ln.h * (lh || 1.3);
      var base = y + ln.h * 0.86;
      var cx = x;
      if (align === "center") cx = x + (maxW - ln.w) / 2;
      else if (align === "right") cx = x + (maxW - ln.w);
      ln.segs.forEach(function (s) {
        doc.setFont(FONT, s.style); doc.setFontSize(s.size);
        var c = s.color || defColor || INK;
        doc.setTextColor(c[0], c[1], c[2]);
        if (s.text) doc.text(s.text, cx, base);
        if (s.underline && s.text.trim()) { doc.setDrawColor(c[0], c[1], c[2]); doc.setLineWidth(Math.max(0.5, s.size / 18)); doc.line(cx, base + 1.5, cx + s.w, base + 1.5); }
        cx += s.w;
      });
      y += h;
    });
    return y;
  }
  function linesHeight(lines, lh) { return lines.reduce(function (s, l) { return s + l.h * (lh || 1.3); }, 0); }

  /* ---------- 7.1 Modelo de documento → PDF (Word, Excel, imágenes) ---------- */
  function modelToPdf(model, onProgress) {
    return pdfEngines(model).then(function (eng) {
      var made = newPdf(eng[0], eng[1], { orientation: model.landscape ? "landscape" : "portrait", title: model.title });
      var doc = made.doc, FONT = made.FONT;
      var W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
      var M = 50, CW = W - 2 * M, y = M, tick = makeSlicer();
      var blocks = model.blocks || [], i = 0, pageHasContent = false;

      function ensure(h) { if (y + h > H - M && pageHasContent) { doc.addPage(); y = M; pageHasContent = false; } }
      function writeLines(lines, x, maxW, align, lh) {
        lines.forEach(function (ln) { ensure(ln.h * (lh || 1.3)); y = drawLines(doc, FONT, [ln], x, y, maxW, align, lh); pageHasContent = true; });
      }

      function table(b) {
        var rows = (b.rows || []).map(function (r) { return r.map(function (v) { return v == null ? "" : String(v); }); });
        if (!rows.length) return Promise.resolve();
        var cols = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
        rows = rows.map(function (r) { while (r.length < cols) r.push(""); return r; });
        var fs = cols <= 6 ? 9 : cols <= 10 ? 8 : cols <= 16 ? 7 : 6;
        var head = b.header ? [rows[0]] : undefined;
        var body = b.header ? rows.slice(1) : rows;
        var CH = 250, k = 0;
        // La tabla se dibuja en tandas de 250 filas, cediendo el turno entre tandas.
        function chunk() {
          if (k >= Math.max(body.length, 1) && k > 0) return Promise.resolve();
          var part = body.slice(k, k + CH);
          window.autoTable(doc, {
            startY: y, head: k === 0 ? head : undefined, body: part, theme: "grid", showHead: "firstPage",
            margin: { left: M, right: M, top: M, bottom: M }, tableWidth: "auto",
            styles: { font: FONT, fontSize: fs, cellPadding: 3.5, overflow: "linebreak", textColor: INK, lineColor: [205, 200, 192], lineWidth: 0.5, valign: "middle" },
            headStyles: { fillColor: [237, 235, 230], textColor: [20, 20, 22], fontStyle: "bold" }
          });
          y = doc.lastAutoTable.finalY; pageHasContent = true;
          k += CH;
          if (k >= body.length) { y += 12; return Promise.resolve(); }
          return yieldUI().then(chunk);
        }
        return chunk();
      }

      function image(b) {
        var p = b.norm ? Promise.resolve(b.norm) : normalizeImage(b.blob, { maxSide: 2400 });
        return p.then(function (n) {
          var wPt = n.w * 0.75, hPt = n.h * 0.75; // 96 ppp → puntos
          // Imágenes dentro del texto: como máximo 75 % del ancho y media página de alto
          var s = b.full ? Math.min(1, CW / wPt, (H - 2 * M) / hPt) : Math.min(1, CW * 0.75 / wPt, (H - 2 * M) * 0.5 / hPt);
          wPt *= s; hPt *= s;
          ensure(hPt);
          doc.addImage(n.bytes, n.mime === "image/png" ? "PNG" : "JPEG", M + (CW - wPt) / 2, y, wPt, hPt, undefined, "FAST");
          y += hPt + 10; pageHasContent = true;
        }).catch(function () {
          writeLines(layoutRuns(doc, FONT, [{ text: "[Imagen no compatible omitida]", italic: true }], CW, 10), M, CW);
        });
      }

      function step() {
        if (i >= blocks.length) return Promise.resolve();
        var b = blocks[i++], p = null;
        if (onProgress) onProgress(0.92 * i / blocks.length);
        if (b.type === "heading") {
          var hs = b.level === 1 ? 20 : b.level === 2 ? 16 : 13.5;
          var hl = layoutRuns(doc, FONT, (b.runs || []).map(function (r) { return Object.assign({}, r, { bold: true }); }), CW, hs);
          if (pageHasContent) y += hs * 0.5;
          ensure(linesHeight(hl, 1.25) + 24);
          writeLines(hl, M, CW, "left", 1.25); y += 4;
        } else if (b.type === "para") {
          var pl = layoutRuns(doc, FONT, b.runs, CW, 11);
          writeLines(pl, M, CW, b.align, 1.38); y += 6;
        } else if (b.type === "list") {
          (b.items || []).forEach(function (it, n) {
            var ind = 16 + (it.level || 0) * 16;
            var ll = layoutRuns(doc, FONT, it.runs, CW - ind, 11);
            ensure(ll[0].h * 1.38);
            doc.setFont(FONT, "normal"); doc.setFontSize(11); doc.setTextColor(INK[0], INK[1], INK[2]);
            doc.text(b.ordered ? (n + 1) + "." : "•", M + ind - 12, y + 11 * 0.86);
            writeLines(ll, M + ind, CW - ind, "left", 1.38); y += 2;
          });
          y += 6;
        } else if (b.type === "table") {
          p = table(b);
        } else if (b.type === "image") {
          p = image(b);
        } else if (b.type === "pagebreak") {
          if (pageHasContent) { doc.addPage(); y = M; pageHasContent = false; }
        }
        return (p || Promise.resolve()).then(tick).then(step);
      }
      return step().then(yieldUI).then(function () {
        var out = doc.output("arraybuffer");
        if (onProgress) onProgress(1);
        return new Blob([out], { type: MIME.pdf });
      });
    });
  }

  /* ---------- 7.2 Imágenes → PDF (una por página A4, centradas) ---------- */
  function imagesToPdf(files, onProgress) {
    return Engines.jspdf().then(function (jsPDF) {
      var doc = null, i = 0, M = 28;
      function step() {
        if (i >= files.length) return Promise.resolve();
        var f = files[i++];
        return normalizeImage(f, { maxSide: 3500, prefer: "jpeg" }).then(function (n) {
          var orient = n.w > n.h ? "landscape" : "portrait";
          if (!doc) doc = new jsPDF({ unit: "pt", format: "a4", orientation: orient, compress: true });
          else doc.addPage("a4", orient);
          var W = doc.internal.pageSize.getWidth(), H = doc.internal.pageSize.getHeight();
          var s = Math.min((W - 2 * M) / n.w, (H - 2 * M) / n.h);
          var w = n.w * s, h = n.h * s;
          doc.addImage(n.bytes, "JPEG", (W - w) / 2, (H - h) / 2, w, h, undefined, "FAST");
          if (onProgress) onProgress(0.95 * i / files.length);
        }).then(yieldUI).then(step);
      }
      return step().then(function () {
        doc.setProperties({ creator: BRAND.name });
        return new Blob([doc.output("arraybuffer")], { type: MIME.pdf });
      });
    });
  }

  /* ===========================================================
     8. LECTURA DE PDF (pdf.js): texto, tablas y páginas como imagen
     =========================================================== */
  function pageLines(tc) {
    var items = [];
    tc.items.forEach(function (it) {
      // Los «espacios falsos» que añade el motor entre columnas se descartan: la separación se mide por posición.
      if (typeof it.str !== "string" || !it.str.trim()) return;
      var t = it.transform, h = Math.hypot(t[2], t[3]) || it.height || 10;
      items.push({ s: it.str, x: t[4], y: t[5], w: it.width || 0, h: h });
    });
    items.sort(function (a, b) { return (b.y - a.y) || (a.x - b.x); });
    var lines = [];
    items.forEach(function (it) {
      var ln = lines[lines.length - 1];
      if (ln && Math.abs(ln.y - it.y) <= Math.max(2, Math.min(ln.h, it.h) * 0.45)) { ln.items.push(it); ln.h = Math.max(ln.h, it.h); }
      else lines.push({ y: it.y, h: it.h, items: [it] });
    });
    lines.forEach(function (ln) {
      ln.items.sort(function (a, b) { return a.x - b.x; });
      var text = "", cells = [], cell = null, prev = null;
      ln.items.forEach(function (it) {
        var gap = prev ? it.x - (prev.x + prev.w) : 0;
        if (!prev || gap > ln.h * 1.1) {               // separación grande = nueva celda
          cell = { x: it.x, s: it.s }; cells.push(cell);
          text += (prev ? " " : "") + it.s;
        } else {
          var sp = gap > ln.h * 0.15 && !/\s$/.test(prev.s) && !/^\s/.test(it.s) ? " " : "";
          cell.s += sp + it.s; text += sp + it.s;
        }
        prev = it;
      });
      ln.text = text.replace(/\s+/g, " ").trim();
      ln.cells = cells.map(function (c) { return { x: c.x, s: c.s.replace(/\s+/g, " ").trim() }; }).filter(function (c) { return c.s; });
      ln.x = ln.items[0].x;
    });
    return lines.filter(function (l) { return l.text; });
  }

  function median(arr) {
    if (!arr.length) return 11;
    var a = arr.slice().sort(function (x, y) { return x - y; });
    return a[Math.floor(a.length / 2)];
  }

  /* Calcula una escala de render segura (memoria y límites del lienzo). */
  function renderScale(page, dpi, maxSide) {
    var vp = page.getViewport({ scale: 1 });
    var s = dpi / 72;
    var fit = fitSize(vp.width * s, vp.height * s, maxSide || 4096, 16e6);
    return s * fit.w / (vp.width * s);
  }
  function renderPage(page, dpi, mime, quality, maxSide) {
    var scale = renderScale(page, dpi, maxSide);
    var vp = page.getViewport({ scale: scale });
    var c = document.createElement("canvas");
    c.width = Math.ceil(vp.width); c.height = Math.ceil(vp.height);
    var ctx = c.getContext("2d");
    ctx.fillStyle = "#FFFFFF"; ctx.fillRect(0, 0, c.width, c.height);
    return page.render({ canvasContext: ctx, viewport: vp, background: "#FFFFFF" }).promise.then(function () {
      return canvasToBlob(c, mime, quality);
    }).then(function (blob) {
      var r = { blob: blob, w: c.width, h: c.height };
      c.width = c.height = 0;
      page.cleanup();
      return r;
    });
  }
  /* Recorre todas las páginas de un PDF, una a una, cediendo el turno entre páginas. */
  function eachPage(pdf, fn, onProgress) {
    var n = 1, out = [];
    function step() {
      if (n > pdf.numPages) return Promise.resolve(out);
      var k = n++;
      return pdf.getPage(k).then(function (page) { return fn(page, k); }).then(function (r) {
        out.push(r);
        if (onProgress) onProgress((k) / pdf.numPages, k);
      }).then(yieldUI).then(step);
    }
    return step();
  }

  function pdfToImages(pdf, mime, onProgress) {
    return eachPage(pdf, function (page) { return renderPage(page, 150, MIME[mime], 0.92); }, onProgress);
  }

  /* Líneas con varias celdas → filas alineadas en columnas (tolerancia de 8 pt) */
  function cellsToRows(lines) {
    var xs = [];
    lines.forEach(function (l) { l.cells.forEach(function (c) { xs.push(c.x); }); });
    xs.sort(function (a, b) { return a - b; });
    var anchors = [];
    xs.forEach(function (x) { if (!anchors.length || x - anchors[anchors.length - 1] > 8) anchors.push(x); });
    if (anchors.length > 60) anchors = anchors.slice(0, 60);
    var rows = lines.map(function (l) {
      var row = [];
      l.cells.forEach(function (c) {
        var col = 0;
        for (var a = 0; a < anchors.length; a++) if (anchors[a] <= c.x + 4) col = a;
        while (row[col] != null && row[col] !== "") col++;
        row[col] = c.s;
      });
      for (var j = 0; j < row.length; j++) if (row[j] == null) row[j] = "";
      return row;
    });
    var maxC = rows.reduce(function (m, r) { return Math.max(m, r.length); }, 0);
    var used = [];
    for (var c = 0; c < maxC; c++) used[c] = rows.some(function (r) { return r[c]; });
    return rows.map(function (r) { return r.filter(function (_, ci) { return used[ci]; }); });
  }

  /* PDF → modelo de documento editable (para Word) */
  function pdfToModel(pdf, onProgress) {
    var blocks = [], scanned = 0;
    return eachPage(pdf, function (page, k) {
      return page.getTextContent().then(function (tc) {
        var lines = pageLines(tc);
        var chars = lines.reduce(function (s, l) { return s + l.text.length; }, 0);
        if (k > 1) blocks.push({ type: "pagebreak" });
        if (chars < 25) {
          // Página escaneada (sin texto): se inserta como imagen para no perderla.
          scanned++;
          return renderPage(page, 150, "image/jpeg", 0.88, 2400).then(function (r) { blocks.push({ type: "image", blob: r.blob, full: true }); });
        }
        var body = median(lines.map(function (l) { return l.h; }));
        var para = null;
        function flush() { if (para) { blocks.push(para.block); para = null; } }
        // Marca las líneas que forman una tabla (2+ líneas seguidas con 2+ columnas)
        var inTable = [];
        for (var a = 0; a < lines.length;) {
          var e = a;
          while (e < lines.length && lines[e].cells.length >= 2) e++;
          if (e - a >= 2) { for (var t = a; t < e; t++) inTable[t] = a; a = e; } else a = Math.max(a + 1, e);
        }
        lines.forEach(function (ln, idx) {
          if (inTable[idx] != null) {
            if (inTable[idx] === idx) {
              flush();
              var run = lines.filter(function (_, j) { return inTable[j] === idx; });
              blocks.push({ type: "table", rows: cellsToRows(run), header: true });
            }
            return;
          }
          var prev = lines[idx - 1];
          var ratio = ln.h / body;
          var isHead = ratio >= 1.25 && ln.text.length < 160;
          var isBullet = /^([•●▪◦‣\-–]|\d{1,3}[.)])\s/.test(ln.text);
          var gap = prev ? prev.y - ln.y : 0;
          var breakPara = !para || isHead !== para.head || isBullet || Math.abs(ln.h - para.h) > body * 0.2 ||
            gap > Math.max(prev.h, ln.h) * 1.75 || gap < 0;
          if (breakPara) {
            flush();
            para = {
              head: isHead, h: ln.h,
              block: isHead ? { type: "heading", level: ratio >= 1.6 ? 1 : 2, runs: [{ text: ln.text }] } : { type: "para", runs: [{ text: ln.text }] }
            };
          } else {
            var r = para.block.runs[0];
            if (/[a-záéíóúñ]-$/i.test(r.text) && /^[a-záéíóúñ]/.test(ln.text)) r.text = r.text.slice(0, -1) + ln.text;
            else r.text += " " + ln.text;
          }
        });
        flush();
      });
    }, onProgress).then(function () {
      return { model: { blocks: blocks }, scanned: scanned, pages: pdf.numPages };
    });
  }

  /* PDF → hojas de cálculo: columnas alineadas por posición horizontal */
  function pdfToSheets(pdf, onProgress) {
    var sheets = [], empty = 0;
    return eachPage(pdf, function (page, k) {
      return page.getTextContent().then(function (tc) {
        var lines = pageLines(tc);
        if (!lines.length) { empty++; return; }
        var rows = cellsToRows(lines);
        sheets.push({ name: "Página " + k, rows: rows });
      });
    }, onProgress).then(function () {
      if (sheets.length > 40) {   // Muchas páginas: todo en una sola hoja
        var all = [];
        sheets.forEach(function (s, n) { if (n) all.push([]); all.push(["— " + s.name + " —"]); all = all.concat(s.rows); });
        sheets = [{ name: "Contenido", rows: all }];
      }
      return { sheets: sheets, empty: empty };
    });
  }

  /* ===========================================================
     9. LECTURA DE WORD (.docx) con mammoth → modelo de documento
     =========================================================== */
  function docxToModel(buf, onProgress) {
    return Engines.mammoth().then(function (mammoth) {
      if (onProgress) onProgress(0.3);
      return mammoth.convertToHtml({ arrayBuffer: buf });
    }).then(function (res) {
      if (onProgress) onProgress(0.8);
      var dom = new DOMParser().parseFromString("<body>" + res.value + "</body>", "text/html");
      var blocks = [];
      function inline(node, st, runs, imgs) {
        node.childNodes.forEach(function (n) {
          if (n.nodeType === 3) {
            var t = n.nodeValue.replace(/\s+/g, " ");
            if (t) runs.push({ text: t, bold: st.b, italic: st.i, underline: st.u });
          } else if (n.nodeType === 1) {
            var tag = n.tagName.toLowerCase();
            if (tag === "br") runs.push({ text: "\n" });
            else if (tag === "img") imgs.push(n.getAttribute("src"));
            else if (tag === "ul" || tag === "ol") return;
            else inline(n, { b: st.b || tag === "strong" || tag === "b", i: st.i || tag === "em" || tag === "i", u: st.u || tag === "u" }, runs, imgs);
          }
        });
        return runs;
      }
      function pushImages(imgs) {
        imgs.forEach(function (src) {
          if (src && src.indexOf("data:") === 0) {
            var mime = src.slice(5, src.indexOf(";"));
            blocks.push({ type: "image", blob: new Blob([dataURLToBytes(src)], { type: mime }) });
          }
        });
      }
      function cleanRuns(runs) {
        if (runs.length) { runs[0].text = runs[0].text.replace(/^ +/, ""); var l = runs[runs.length - 1]; l.text = l.text.replace(/ +$/, ""); }
        return runs.filter(function (r) { return r.text !== ""; });
      }
      function list(el, ordered, level, items) {
        Array.prototype.forEach.call(el.children, function (li) {
          if (li.tagName.toLowerCase() !== "li") return;
          var imgs = [];
          var runs = cleanRuns(inline(li, {}, [], imgs));
          if (runs.length) items.push({ runs: runs, level: level });
          Array.prototype.forEach.call(li.children, function (sub) {
            var t = sub.tagName.toLowerCase();
            if (t === "ul" || t === "ol") list(sub, t === "ol", level + 1, items);
          });
        });
        return items;
      }
      function walk(el) {
        Array.prototype.forEach.call(el.childNodes, function (n) {
          if (n.nodeType === 3) {
            var t = n.nodeValue.trim();
            if (t) blocks.push({ type: "para", runs: [{ text: t }] });
            return;
          }
          if (n.nodeType !== 1) return;
          var tag = n.tagName.toLowerCase(), imgs = [];
          if (/^h[1-6]$/.test(tag)) {
            var hr = cleanRuns(inline(n, {}, [], imgs));
            if (hr.length) blocks.push({ type: "heading", level: Math.min(3, +tag[1]), runs: hr });
            pushImages(imgs);
          } else if (tag === "p") {
            var pr = cleanRuns(inline(n, {}, [], imgs));
            if (pr.length) blocks.push({ type: "para", runs: pr });
            pushImages(imgs);
          } else if (tag === "ul" || tag === "ol") {
            var items = list(n, tag === "ol", 0, []);
            if (items.length) blocks.push({ type: "list", ordered: tag === "ol", items: items });
          } else if (tag === "table") {
            var rows = [], hasTh = false;
            Array.prototype.forEach.call(n.querySelectorAll("tr"), function (tr) {
              if (tr.closest("table") !== n) return;
              var row = [];
              Array.prototype.forEach.call(tr.children, function (td) {
                if (td.tagName === "TH") hasTh = true;
                row.push(td.textContent.replace(/\s+/g, " ").trim());
                var span = +td.getAttribute("colspan") || 1;
                for (var s = 1; s < span; s++) row.push("");
              });
              rows.push(row);
            });
            if (rows.length) blocks.push({ type: "table", rows: rows, header: hasTh || rows.length > 1 });
          } else if (tag === "img") {
            pushImages([n.getAttribute("src")]);
          } else {
            walk(n);
          }
        });
      }
      walk(dom.body);
      var notes = (res.messages || []).filter(function (m) { return m.type === "warning"; }).length;
      return { model: { blocks: blocks }, warnings: notes };
    });
  }

  /* Modelo → hojas de Excel (las tablas a su hoja; el texto en otra) */
  function modelToSheets(model) {
    var sheets = [], text = [], t = 0;
    (model.blocks || []).forEach(function (b) {
      var line = function (runs) { return (runs || []).map(function (r) { return r.text; }).join("").trim(); };
      if (b.type === "table") sheets.push({ name: "Tabla " + (++t), rows: b.rows });
      else if (b.type === "heading" || b.type === "para") { var s = line(b.runs); if (s) text.push([s]); }
      else if (b.type === "list") b.items.forEach(function (it) { text.push(["• " + line(it.runs)]); });
    });
    if (text.length) sheets.unshift({ name: sheets.length ? "Texto" : "Documento", rows: text });
    if (!sheets.length) sheets.push({ name: "Documento", rows: [["(El documento no contiene texto ni tablas)"]] });
    return sheets;
  }

  /* ===========================================================
     10. LECTURA DE EXCEL (.xlsx) con SheetJS → modelo de documento
     =========================================================== */
  function xlsxToModel(buf, onProgress) {
    return Engines.xlsx().then(function (XLSX) {
      if (onProgress) onProgress(0.3);
      return yieldUI().then(function () {
        var wb;
        try { wb = XLSX.read(new Uint8Array(buf), { type: "array", dense: true }); }
        catch (e) {
          if (/password|encrypt/i.test(e.message)) throw UserError("Excel protegido con contraseña", "Ábrelo en Excel, quita la contraseña y vuelve a intentarlo.");
          throw e;
        }
        var blocks = [], truncated = 0, maxCols = 0, names = wb.SheetNames || [];
        names.forEach(function (name, si) {
          var rows = XLSX.utils.sheet_to_json(wb.Sheets[name], { header: 1, raw: false, defval: "", blankrows: false });
          rows = rows.filter(function (r) { return r.some(function (v) { return String(v).trim() !== ""; }); });
          var used = 0;
          rows.forEach(function (r) { for (var c = r.length - 1; c >= 0; c--) if (String(r[c]).trim() !== "") { used = Math.max(used, c + 1); break; } });
          rows = rows.map(function (r) { return r.slice(0, used); });
          if (rows.length > BRAND.limits.maxSheetRows) { truncated++; rows = rows.slice(0, BRAND.limits.maxSheetRows); }
          maxCols = Math.max(maxCols, used);
          if (si > 0) blocks.push({ type: "pagebreak" });
          if (names.length > 1) blocks.push({ type: "heading", level: 2, runs: [{ text: name }] });
          if (rows.length) blocks.push({ type: "table", rows: rows, header: true });
          else blocks.push({ type: "para", runs: [{ text: "(Hoja vacía)", italic: true }] });
        });
        return { model: { blocks: blocks, landscape: maxCols > 6 }, truncated: truncated, sheets: names.length };
      });
    });
  }

  /* ===========================================================
     11. LECTURA DE POWERPOINT (.pptx) → diapositivas con posiciones reales
     -----------------------------------------------------------
     Se leen los XML de cada diapositiva: cuadros de texto (con la herencia
     de tamaños y viñetas del patrón), imágenes, tablas, fondos y colores
     del tema. Animaciones, gráficos y SmartArt no se reproducen.
     =========================================================== */
  function parseXml(str) {
    var d = new DOMParser().parseFromString(str, "application/xml");
    if (d.getElementsByTagName("parsererror").length) throw new Error("Corrupted XML");
    return d;
  }
  function kids(el, name) {
    var out = [];
    if (!el) return out;
    for (var c = el.firstElementChild; c; c = c.nextElementSibling) if (!name || c.localName === name) out.push(c);
    return out;
  }
  function kid(el) {
    for (var i = 1; i < arguments.length && el; i++) el = kids(el, arguments[i])[0] || null;
    return el || null;
  }
  function desc(el, name) { return el ? Array.prototype.slice.call(el.getElementsByTagNameNS("*", name)) : []; }
  function rAttr(el, name) { return el ? (el.getAttributeNS(NS_OFFREL, name) || el.getAttribute("r:" + name)) : null; }
  function num(v, d) { var n = parseInt(v, 10); return isNaN(n) ? d : n; }
  function dirOf(p) { return p.slice(0, p.lastIndexOf("/") + 1); }
  function resolvePart(base, target) {
    if (!target) return null;
    if (target.charAt(0) === "/") return target.slice(1);
    var parts = (dirOf(base) + target).split("/"), out = [];
    parts.forEach(function (s) { if (s === "..") out.pop(); else if (s && s !== ".") out.push(s); });
    return out.join("/");
  }
  function readRels(zip, partPath) {
    var relPath = dirOf(partPath) + "_rels/" + partPath.slice(partPath.lastIndexOf("/") + 1) + ".rels";
    var f = zip.file(relPath);
    if (!f) return Promise.resolve({});
    return f.async("string").then(function (s) {
      var map = {};
      kids(parseXml(s).documentElement, "Relationship").forEach(function (r) {
        map[r.getAttribute("Id")] = { type: r.getAttribute("Type") || "", target: resolvePart(partPath, r.getAttribute("Target")), external: r.getAttribute("TargetMode") === "External" };
      });
      return map;
    });
  }
  function relOfType(rels, suffix) {
    for (var k in rels) if (rels[k].type.slice(-suffix.length) === suffix) return rels[k];
    return null;
  }
  function readXmlPart(zip, path) {
    var f = path && zip.file(path);
    return f ? f.async("string").then(parseXml) : Promise.resolve(null);
  }

  /* --- Colores (srgb, del tema, con luminosidad) --- */
  function hexToRgb(h) { h = String(h || "").replace("#", ""); if (h.length !== 6) return null; return [parseInt(h.slice(0, 2), 16), parseInt(h.slice(2, 4), 16), parseInt(h.slice(4, 6), 16)]; }
  function rgbToHsl(c) {
    var r = c[0] / 255, g = c[1] / 255, b = c[2] / 255, mx = Math.max(r, g, b), mn = Math.min(r, g, b), h = 0, s = 0, l = (mx + mn) / 2;
    if (mx !== mn) {
      var d = mx - mn; s = l > 0.5 ? d / (2 - mx - mn) : d / (mx + mn);
      h = mx === r ? (g - b) / d + (g < b ? 6 : 0) : mx === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6;
    }
    return [h, s, l];
  }
  function hslToRgb(hsl) {
    var h = hsl[0], s = hsl[1], l = hsl[2];
    if (!s) { var v = Math.round(l * 255); return [v, v, v]; }
    function f(p, q, t) { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; }
    var q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    return [Math.round(f(p, q, h + 1 / 3) * 255), Math.round(f(p, q, h) * 255), Math.round(f(p, q, h - 1 / 3) * 255)];
  }
  function makeColorResolver(theme, clrMap) {
    return function colorOf(holder) {
      if (!holder) return null;
      var c = kid(holder, "srgbClr") || kid(holder, "schemeClr") || kid(holder, "sysClr") || kid(holder, "prstClr");
      if (!c) return null;
      var rgb = null;
      if (c.localName === "srgbClr") rgb = hexToRgb(c.getAttribute("val"));
      else if (c.localName === "sysClr") rgb = hexToRgb(c.getAttribute("lastClr") || (c.getAttribute("val") === "window" ? "FFFFFF" : "000000"));
      else if (c.localName === "prstClr") rgb = ({ black: [0, 0, 0], white: [255, 255, 255], red: [255, 0, 0], blue: [0, 0, 255], green: [0, 128, 0] })[c.getAttribute("val")] || null;
      else {
        var v = c.getAttribute("val");
        var mapped = clrMap[v] || v;
        rgb = theme[mapped] || null;
      }
      if (!rgb) return null;
      var lumMod = kid(c, "lumMod"), lumOff = kid(c, "lumOff"), tint = kid(c, "tint"), shade = kid(c, "shade");
      if (lumMod || lumOff) {
        var hsl = rgbToHsl(rgb);
        hsl[2] = Math.min(1, Math.max(0, hsl[2] * (lumMod ? num(lumMod.getAttribute("val"), 100000) / 100000 : 1) + (lumOff ? num(lumOff.getAttribute("val"), 0) / 100000 : 0)));
        rgb = hslToRgb(hsl);
      }
      if (tint) { var t = num(tint.getAttribute("val"), 100000) / 100000; rgb = rgb.map(function (x) { return Math.round(x * t + 255 * (1 - t)); }); }
      if (shade) { var sh = num(shade.getAttribute("val"), 100000) / 100000; rgb = rgb.map(function (x) { return Math.round(x * sh); }); }
      return rgb;
    };
  }

  function xfrmOf(el) {
    var x = el && (kid(el, "spPr", "xfrm") || kid(el, "xfrm") || kid(el, "grpSpPr", "xfrm"));
    if (!x) return null;
    var off = kid(x, "off"), ext = kid(x, "ext");
    if (!off || !ext) return null;
    return { x: num(off.getAttribute("x"), 0), y: num(off.getAttribute("y"), 0), w: num(ext.getAttribute("cx"), 0), h: num(ext.getAttribute("cy"), 0) };
  }
  function phOf(sp) {
    var ph = kid(sp, "nvSpPr", "nvPr", "ph") || kid(sp, "nvPicPr", "nvPr", "ph");
    if (!ph) return null;
    return { type: ph.getAttribute("type") || "obj", idx: ph.getAttribute("idx") };
  }
  function phClass(t) { return t === "title" || t === "ctrTitle" ? "title" : (t === "body" || t === "subTitle" || t === "obj") ? "body" : "other"; }
  /* Índice de marcadores (placeholders) de un diseño o patrón */
  function phIndex(doc) {
    var list = [];
    if (doc) desc(doc.documentElement, "sp").forEach(function (sp) { var ph = phOf(sp); if (ph) list.push({ ph: ph, sp: sp }); });
    return {
      find: function (ph) {
        var hit = null;
        if (ph.idx != null) hit = list.filter(function (e) { return e.ph.idx === ph.idx; })[0];
        if (!hit) hit = list.filter(function (e) { return e.ph.type === ph.type; })[0];
        if (!hit) hit = list.filter(function (e) { return phClass(e.ph.type) === phClass(ph.type) && phClass(ph.type) !== "other"; })[0];
        return hit ? hit.sp : null;
      }
    };
  }
  function lvlPPr(container, lvl) { return container ? kid(container, "lvl" + (lvl + 1) + "pPr") : null; }
  function bgOf(cSldDoc, colorOf) {
    var bg = cSldDoc && kid(cSldDoc.documentElement, "cSld", "bg");
    if (!bg) return null;
    var bgPr = kid(bg, "bgPr");
    if (bgPr) {
      var sf = kid(bgPr, "solidFill");
      if (sf) return { color: colorOf(sf) };
      var gf = kid(bgPr, "gradFill");
      if (gf) { var gs = desc(gf, "gs")[0]; if (gs) return { color: colorOf(gs) }; }
      var blip = kid(bgPr, "blipFill", "blip");
      if (blip) return { embed: rAttr(blip, "embed") };
    }
    var ref = kid(bg, "bgRef");
    if (ref) return { color: colorOf(ref) };
    return null;
  }

  function pptxToSlides(buf, onProgress) {
    var zip, presPath = "ppt/presentation.xml";
    return Engines.jszip().then(function (JSZip) { return JSZip.loadAsync(buf); }).then(function (z) {
      zip = z;
      if (!zip.file(presPath)) throw UserError("No es una presentación válida", "El archivo no contiene diapositivas de PowerPoint.");
      return Promise.all([readXmlPart(zip, presPath), readRels(zip, presPath)]);
    }).then(function (r) {
      var pres = r[0], rels = r[1];
      var sz = kid(pres.documentElement, "sldSz");
      var W = num(sz && sz.getAttribute("cx"), 12192000), H = num(sz && sz.getAttribute("cy"), 6858000);
      var ids = desc(pres.documentElement, "sldId").map(function (s) { var rel = rels[rAttr(s, "id")]; return rel && rel.target; }).filter(Boolean);
      var cache = {};    // diseños y patrones ya leídos
      function loadLayer(path) {
        if (!cache[path]) cache[path] = Promise.all([readXmlPart(zip, path), readRels(zip, path)]).then(function (x) { return { doc: x[0], rels: x[1], path: path, ph: phIndex(x[0]) }; });
        return cache[path];
      }
      var themeCache = {};
      function loadTheme(path) {
        if (!themeCache[path]) themeCache[path] = readXmlPart(zip, path).then(function (doc) {
          var t = {};
          var cs = doc && desc(doc.documentElement, "clrScheme")[0];
          kids(cs).forEach(function (c) {
            var v = kid(c, "srgbClr"), s = kid(c, "sysClr");
            t[c.localName] = v ? hexToRgb(v.getAttribute("val")) : s ? hexToRgb(s.getAttribute("lastClr") || "000000") : null;
          });
          return t;
        });
        return themeCache[path];
      }
      var slides = [], i = 0, notes = { charts: 0, skippedImages: 0 };
      function nextSlide() {
        if (i >= ids.length) return Promise.resolve();
        var path = ids[i++];
        var slide = { w: W, h: H, items: [], bg: null };
        var sDoc, sRels, layout, master, colorOf;
        return Promise.all([readXmlPart(zip, path), readRels(zip, path)]).then(function (x) {
          sDoc = x[0]; sRels = x[1];
          if (!sDoc) return;
          var lr = relOfType(sRels, "/slideLayout");
          return (lr ? loadLayer(lr.target) : Promise.resolve(null)).then(function (l) {
            layout = l;
            var mr = layout && relOfType(layout.rels, "/slideMaster");
            return mr ? loadLayer(mr.target) : null;
          }).then(function (m) {
            master = m;
            var tr = master && relOfType(master.rels, "/theme");
            return tr ? loadTheme(tr.target) : {};
          }).then(function (theme) {
            var clrMap = {};
            var cm = master && master.doc && kid(master.doc.documentElement, "clrMap");
            if (cm) Array.prototype.forEach.call(cm.attributes, function (a) { clrMap[a.name] = a.value; });
            var ov = kid(sDoc.documentElement, "clrMapOvr", "overrideClrMapping");
            if (ov) Array.prototype.forEach.call(ov.attributes, function (a) { clrMap[a.name] = a.value; });
            colorOf = makeColorResolver(theme, clrMap);
            var txStyles = master && master.doc && kid(master.doc.documentElement, "txStyles");

            // Fondo: diapositiva → diseño → patrón
            var bgSrc = [[sDoc, sRels], [layout && layout.doc, layout && layout.rels], [master && master.doc, master && master.rels]];
            for (var b = 0; b < bgSrc.length && !slide.bg; b++) {
              var bg = bgOf(bgSrc[b][0], colorOf);
              if (bg) { if (bg.embed && bgSrc[b][1] && bgSrc[b][1][bg.embed]) bg.path = bgSrc[b][1][bg.embed].target; slide.bg = bg; }
            }

            // Recorre el árbol de formas (con grupos y sus transformaciones)
            var media = [];
            function walkTree(tree, map) {
              kids(tree).forEach(function (el) {
                var n = el.localName;
                if (n === "grpSp") {
                  var gx = kid(el, "grpSpPr", "xfrm"), inner = map;
                  if (gx) {
                    var off = kid(gx, "off"), ext = kid(gx, "ext"), cho = kid(gx, "chOff"), che = kid(gx, "chExt");
                    if (off && ext && cho && che) {
                      var ox = num(off.getAttribute("x"), 0), oy = num(off.getAttribute("y"), 0);
                      var sx = num(ext.getAttribute("cx"), 1) / (num(che.getAttribute("cx"), 1) || 1), sy = num(ext.getAttribute("cy"), 1) / (num(che.getAttribute("cy"), 1) || 1);
                      var cx0 = num(cho.getAttribute("x"), 0), cy0 = num(cho.getAttribute("y"), 0);
                      inner = function (r) { return map({ x: ox + (r.x - cx0) * sx, y: oy + (r.y - cy0) * sy, w: r.w * sx, h: r.h * sy }); };
                    }
                  }
                  walkTree(el, inner);
                } else if (n === "sp") {
                  shape(el, map);
                } else if (n === "pic") {
                  var blip = kid(el, "blipFill", "blip"), xf = xfrmOf(el), rel = blip && sRels[rAttr(blip, "embed")];
                  if (!xf) { var ph = phOf(el); var lsp = ph && layout && layout.ph.find(ph); xf = lsp && xfrmOf(lsp); }
                  if (rel && xf && !rel.external) { var it = { kind: "image", box: map(xf), path: rel.target }; slide.items.push(it); media.push(it); }
                } else if (n === "graphicFrame") {
                  var tbl = desc(el, "tbl")[0], gxf = kid(el, "xfrm");
                  if (tbl && gxf) table(tbl, map(boxFromXfrm(gxf)));
                  else if (desc(el, "chart").length || desc(el, "relIds").length) notes.charts++;
                }
              });
            }
            function boxFromXfrm(x) {
              var off = kid(x, "off"), ext = kid(x, "ext");
              return { x: num(off && off.getAttribute("x"), 0), y: num(off && off.getAttribute("y"), 0), w: num(ext && ext.getAttribute("cx"), 0), h: num(ext && ext.getAttribute("cy"), 0) };
            }
            function parasOf(txBody, chain, cls) {
              var out = [];
              kids(txBody, "p").forEach(function (p) {
                var pPr = kid(p, "pPr"), lvl = num(pPr && pPr.getAttribute("lvl"), 0);
                var levels = [pPr].concat(chain.map(function (c) { return lvlPPr(c, lvl); })).filter(Boolean);
                function pick(fn) { for (var k = 0; k < levels.length; k++) { var v = fn(levels[k]); if (v != null) return v; } return null; }
                var defSz = pick(function (l) { var d = kid(l, "defRPr"); return d && d.getAttribute("sz") ? num(d.getAttribute("sz")) / 100 : null; });
                var defB = pick(function (l) { var d = kid(l, "defRPr"); return d && d.getAttribute("b") != null ? d.getAttribute("b") === "1" : null; });
                var defC = pick(function (l) { var d = kid(l, "defRPr"); return d && kid(d, "solidFill") ? colorOf(kid(d, "solidFill")) : null; });
                var algn = pick(function (l) { return l.getAttribute("algn"); });
                var marL = pick(function (l) { return l.getAttribute("marL"); });
                var indent = pick(function (l) { return l.getAttribute("indent"); });
                var bullet = pick(function (l) {
                  if (kid(l, "buNone")) return "";
                  var bc = kid(l, "buChar"); if (bc) return bc.getAttribute("char") || "•";
                  if (kid(l, "buAutoNum")) return "#";
                  return null;
                });
                var runs = [];
                kids(p).forEach(function (r) {
                  if (r.localName === "br") { runs.push({ text: "\n" }); return; }
                  if (r.localName !== "r" && r.localName !== "fld") return;
                  var t = kid(r, "t"), rPr = kid(r, "rPr");
                  var txt = t ? t.textContent : "";
                  if (!txt) return;
                  var cap = rPr && rPr.getAttribute("cap");
                  if (cap === "all") txt = txt.toUpperCase();
                  runs.push({
                    text: txt,
                    size: rPr && rPr.getAttribute("sz") ? num(rPr.getAttribute("sz")) / 100 : null,
                    bold: rPr && rPr.getAttribute("b") != null ? rPr.getAttribute("b") === "1" : !!defB,
                    italic: rPr ? rPr.getAttribute("i") === "1" : false,
                    underline: rPr ? (rPr.getAttribute("u") && rPr.getAttribute("u") !== "none") : false,
                    color: (rPr && kid(rPr, "solidFill") ? colorOf(kid(rPr, "solidFill")) : null) || defC
                  });
                });
                var endPr = kid(p, "endParaRPr");
                var size = defSz || (endPr && endPr.getAttribute("sz") ? num(endPr.getAttribute("sz")) / 100 : null) || (cls === "title" ? 40 : 18);
                out.push({ runs: runs, size: size, align: algn === "ctr" ? "center" : algn === "r" ? "right" : "left", bullet: bullet || "", level: lvl,
                  marL: marL != null ? num(marL, 0) : null, indent: indent != null ? num(indent, 0) : null });
              });
              return out;
            }
            function shape(sp, map) {
              var ph = phOf(sp), cls = ph ? phClass(ph.type) : "other";
              var lsp = ph && layout ? layout.ph.find(ph) : null;
              var msp = ph && master ? master.ph.find(ph) : null;
              var xf = xfrmOf(sp) || xfrmOf(lsp) || xfrmOf(msp);
              if (!xf) return;
              var spPr = kid(sp, "spPr");
              var fill = null;
              if (spPr && kid(spPr, "solidFill")) fill = colorOf(kid(spPr, "solidFill"));
              else if (spPr && !kid(spPr, "noFill")) { var fr = kid(sp, "style", "fillRef"); if (fr && num(fr.getAttribute("idx"), 0) > 0) fill = colorOf(fr); }
              var txBody = kid(sp, "txBody");
              var bodies = [txBody, lsp && kid(lsp, "txBody"), msp && kid(msp, "txBody")];
              var anchor = null;
              bodies.forEach(function (bd) { var bp = bd && kid(bd, "bodyPr"); if (anchor == null && bp && bp.getAttribute("anchor")) anchor = bp.getAttribute("anchor"); });
              var ownBp = txBody && kid(txBody, "bodyPr");
              var autofit = ownBp && kid(ownBp, "normAutofit");
              var scale = autofit && autofit.getAttribute("fontScale") ? num(autofit.getAttribute("fontScale"), 100000) / 100000 : 1;
              var ins = function (a, d) { var v = ownBp && ownBp.getAttribute(a); return v != null ? num(v, d) : d; };
              var chain = [txBody && kid(txBody, "lstStyle"), lsp && kid(lsp, "txBody", "lstStyle"), msp && kid(msp, "txBody", "lstStyle"),
                txStyles && kid(txStyles, cls === "title" ? "titleStyle" : cls === "body" ? "bodyStyle" : "otherStyle")];
              var paras = txBody ? parasOf(txBody, chain, cls) : [];
              var hasText = paras.some(function (p) { return p.runs.some(function (r) { return r.text.trim(); }); });
              if (!hasText && !fill) return;
              slide.items.push({
                kind: "text", box: map(xf), fill: fill, paras: hasText ? paras : [], cls: cls, scale: scale,
                anchor: anchor || (cls === "title" ? "ctr" : "t"),
                ins: { l: ins("lIns", 91440), r: ins("rIns", 91440), t: ins("tIns", 45720), b: ins("bIns", 45720) }
              });
            }
            function table(tbl, box) {
              var widths = desc(tbl, "gridCol").map(function (g) { return num(g.getAttribute("w"), 0); });
              var rows = kids(tbl, "tr").map(function (tr) {
                return { h: num(tr.getAttribute("h"), 370840), cells: kids(tr, "tc").map(function (tc) {
                  var txt = kids(kid(tc, "txBody"), "p").map(function (p) { return desc(p, "t").map(function (t) { return t.textContent; }).join(""); }).join("\n");
                  var tcPr = kid(tc, "tcPr");
                  var sz = desc(tc, "rPr").map(function (r) { return r.getAttribute("sz"); }).filter(Boolean)[0];
                  return { text: txt, merged: tc.getAttribute("hMerge") === "1" || tc.getAttribute("vMerge") === "1", span: num(tc.getAttribute("gridSpan"), 1),
                    fill: tcPr && kid(tcPr, "solidFill") ? colorOf(kid(tcPr, "solidFill")) : null, size: sz ? num(sz) / 100 : 14 };
                }) };
              });
              slide.items.push({ kind: "table", box: box, widths: widths, rows: rows });
            }
            var tree = kid(sDoc.documentElement, "cSld", "spTree");
            walkTree(tree, function (r) { return r; });

            // Carga las imágenes de esta diapositiva (y del fondo)
            if (slide.bg && slide.bg.path) media.push(slide.bg);
            return Promise.all(media.map(function (m) {
              var f = zip.file(m.path);
              if (!f || /\.(emf|wmf|tiff?)$/i.test(m.path)) { m.skip = true; notes.skippedImages++; return null; }
              var ext = (m.path.split(".").pop() || "").toLowerCase();
              var type = ext === "svg" ? "image/svg+xml" : ext === "jpg" || ext === "jpeg" ? "image/jpeg" : "image/" + ext;
              return f.async("uint8array").then(function (u8) {
                return normalizeImage(new Blob([u8], { type: type }), { maxSide: 2000 }).then(function (n) { m.norm = n; m.blob = new Blob([u8], { type: type }); });
              }).catch(function () { m.skip = true; notes.skippedImages++; });
            }));
          });
        }).then(function () {
          slides.push(slide);
          if (onProgress) onProgress(i / Math.max(1, ids.length));
        }).then(yieldUI).then(nextSlide);
      }
      return nextSlide().then(function () {
        if (!slides.length) throw UserError("La presentación no tiene diapositivas", "No hay nada que convertir en este archivo.");
        return { slides: slides, notes: notes };
      });
    });
  }

  /* Diapositivas → PDF respetando posiciones (1 página por diapositiva) */
  function slidesToPdf(slides, onProgress) {
    return pdfEngines(slides.map(function (s) { return s.items.map(function (it) { return [it.paras, it.rows]; }); })).then(function (eng) {
      var first = slides[0];
      var E = 12700; // EMU por punto
      var pw = first.w / E, ph = first.h / E;
      var made = newPdf(eng[0], eng[1], { format: [pw, ph], orientation: pw >= ph ? "landscape" : "portrait" });
      var doc = made.doc, FONT = made.FONT, i = 0;
      function rect(b) { return { x: b.x / E, y: b.y / E, w: b.w / E, h: b.h / E }; }
      function drawText(it) {
        var b = rect(it.box);
        if (it.fill) { doc.setFillColor(it.fill[0], it.fill[1], it.fill[2]); doc.rect(b.x, b.y, b.w, b.h, "F"); }
        if (!it.paras.length) return;
        var x = b.x + it.ins.l / E, w = Math.max(10, b.w - (it.ins.l + it.ins.r) / E);
        var laid = it.paras.map(function (p) {
          var size = p.size * it.scale;
          // Sangría del texto (marL) y posición de la viñeta (marL + indent), como en PowerPoint
          var ind = p.marL != null ? p.marL / E : p.level * size * 0.9 + (p.bullet ? size * 0.9 : 0);
          var bx = p.marL != null ? Math.max(0, (p.marL + (p.indent || 0)) / E) : p.level * size * 0.9;
          if (p.bullet && ind - bx < size * 0.6) ind = bx + size * 0.75;
          var runs = p.runs.length ? p.runs.map(function (r) { return Object.assign({}, r, { size: (r.size || p.size) * it.scale }); }) : [{ text: " ", size: size }];
          return { p: p, ind: ind, bx: bx, size: size, lines: layoutRuns(doc, FONT, runs, w - ind, size) };
        });
        var total = laid.reduce(function (s, l) { return s + linesHeight(l.lines, 1.18) + l.size * 0.25; }, 0);
        var top = b.y + it.ins.t / E, avail = b.h - (it.ins.t + it.ins.b) / E;
        if (it.anchor === "ctr") top += (avail - total) / 2;
        else if (it.anchor === "b") top += avail - total;
        var y = top, autoN = 0;
        laid.forEach(function (l) {
          var col = (l.p.runs[0] && l.p.runs[0].color) || INK;
          if (l.p.bullet && l.p.runs.length) {
            autoN = l.p.bullet === "#" ? autoN + 1 : 0;
            var mark = l.p.bullet === "#" ? autoN + "." : (/[-]/.test(l.p.bullet) || WINANSI_RE.test(l.p.bullet) && FONT === "helvetica" ? "•" : l.p.bullet);
            doc.setFont(FONT, "normal"); doc.setFontSize(l.size); doc.setTextColor(col[0], col[1], col[2]);
            doc.text(mark, x + l.bx, y + l.lines[0].h * 0.86);
          }
          y = drawLines(doc, FONT, l.lines, x + l.ind, y, w - l.ind, l.p.align, 1.18) + l.size * 0.25;
        });
      }
      function drawTable(it) {
        var b = rect(it.box), sumW = it.widths.reduce(function (s, v) { return s + v; }, 0) || 1;
        var y = b.y;
        it.rows.forEach(function (row) {
          var x = b.x, rh = row.h / E, col = 0;
          // Altura real: la mayor de las celdas según su texto
          row.cells.forEach(function (c, ci) {
            var cw = it.widths.slice(ci, ci + c.span).reduce(function (s, v) { return s + v; }, 0) / sumW * b.w;
            var lines = layoutRuns(doc, FONT, [{ text: c.text }], Math.max(10, cw - 8), c.size);
            rh = Math.max(rh, linesHeight(lines, 1.15) + 8);
          });
          row.cells.forEach(function (c) {
            var cw = it.widths.slice(col, col + c.span).reduce(function (s, v) { return s + v; }, 0) / sumW * b.w;
            if (!c.merged) {
              if (c.fill) { doc.setFillColor(c.fill[0], c.fill[1], c.fill[2]); doc.rect(x, y, cw, rh, "F"); }
              doc.setDrawColor(180, 176, 168); doc.setLineWidth(0.6); doc.rect(x, y, cw, rh, "S");
              var lines = layoutRuns(doc, FONT, [{ text: c.text }], Math.max(10, cw - 8), c.size);
              drawLines(doc, FONT, lines, x + 4, y + 4, cw - 8, "left", 1.15);
            }
            x += cw; col += c.span;
          });
          y += rh;
        });
      }
      function step() {
        if (i >= slides.length) return Promise.resolve();
        var s = slides[i++];
        if (i > 1) doc.addPage([pw, ph], pw >= ph ? "landscape" : "portrait");
        var bg = s.bg;
        if (bg && bg.norm && !bg.skip) doc.addImage(bg.norm.bytes, bg.norm.mime === "image/png" ? "PNG" : "JPEG", 0, 0, pw, ph, undefined, "FAST");
        else if (bg && bg.color) { doc.setFillColor(bg.color[0], bg.color[1], bg.color[2]); doc.rect(0, 0, pw, ph, "F"); }
        s.items.forEach(function (it) {
          try {
            if (it.kind === "image" && it.norm && !it.skip) {
              var r = rect(it.box);
              doc.addImage(it.norm.bytes, it.norm.mime === "image/png" ? "PNG" : "JPEG", r.x, r.y, r.w, r.h, undefined, "FAST");
            } else if (it.kind === "text") drawText(it);
            else if (it.kind === "table") drawTable(it);
          } catch (e) { console.warn("[diapositiva " + i + "]", e); }
        });
        if (onProgress) onProgress(0.95 * i / slides.length);
        return yieldUI().then(step);
      }
      return step().then(function () {
        doc.setProperties({ creator: BRAND.name });
        return new Blob([doc.output("arraybuffer")], { type: MIME.pdf });
      });
    });
  }

  /* Diapositivas → modelo de documento (para Word): títulos, viñetas, imágenes y tablas */
  function slidesToModel(slides) {
    var blocks = [];
    slides.forEach(function (s, n) {
      if (n) blocks.push({ type: "pagebreak" });
      var items = s.items.slice().sort(function (a, b) { return (a.box.y - b.box.y) || (a.box.x - b.box.x); });
      var title = items.filter(function (it) { return it.kind === "text" && it.cls === "title" && it.paras.length; })[0];
      var titleText = title ? title.paras.map(function (p) { return p.runs.map(function (r) { return r.text; }).join(""); }).join(" ").trim() : "";
      blocks.push({ type: "heading", level: 1, runs: [{ text: titleText ? titleText : "Diapositiva " + (n + 1) }] });
      items.forEach(function (it) {
        if (it === title) return;
        if (it.kind === "text") {
          var list = null;
          it.paras.forEach(function (p) {
            var runs = p.runs.map(function (r) { return { text: r.text, bold: r.bold, italic: r.italic, underline: r.underline }; });
            if (!runs.some(function (r) { return r.text.trim(); })) return;
            if (p.bullet) {
              if (!list) { list = { type: "list", ordered: p.bullet === "#", items: [] }; blocks.push(list); }
              list.items.push({ runs: runs, level: p.level });
            } else {
              list = null;
              blocks.push({ type: "para", runs: runs, align: p.align === "center" ? "center" : undefined });
            }
          });
        } else if (it.kind === "image" && it.blob && !it.skip) {
          blocks.push({ type: "image", blob: it.blob, norm: it.norm });
        } else if (it.kind === "table") {
          blocks.push({ type: "table", rows: it.rows.map(function (r) { return r.cells.map(function (c) { return c.text; }); }), header: true });
        }
      });
    });
    return { blocks: blocks };
  }

  /* ===========================================================
     12. ENRUTADOR DE CONVERSIONES (origen → destino)
     =========================================================== */
  /* Sub-progreso: reparte la barra entre fases (p. ej. leer 0–30 %, escribir 30–100 %). */
  function phase(report, from, to, label) {
    return function (f, extra) { report(from + (to - from) * Math.max(0, Math.min(1, f || 0)), label + (extra ? " " + extra : "")); };
  }
  function blobsToZip(entries, onProgress) {
    return Engines.jszip().then(function (JSZip) {
      var zip = new JSZip(), used = {};
      entries.forEach(function (e) {
        var n = e.name, k = 2;
        while (used[n.toLowerCase()]) { n = e.name.replace(/(\.[^.]+)?$/, " (" + (k++) + ")$1"); }
        used[n.toLowerCase()] = 1;
        zip.file(n, e.blob, { compression: "STORE" });
      });
      return zip.generateAsync({ type: "blob", mimeType: MIME.zip, streamFiles: true }, function (m) { if (onProgress) onProgress(m.percent / 100); });
    });
  }
  function pagesResult(pages, base, target, report) {
    var ext = target === "jpg" ? "jpg" : target;
    if (pages.length === 1) return Promise.resolve({ blob: pages[0].blob, name: base + "." + ext, note: "1 imagen" });
    var pad = String(pages.length).length;
    var entries = pages.map(function (p, n) { return { blob: p.blob, name: base + "-" + String(n + 1).padStart(Math.max(2, pad), "0") + "." + ext }; });
    return blobsToZip(entries, report).then(function (zip) {
      return { blob: zip, name: base + "-" + ext + ".zip", note: pages.length + " imágenes en un ZIP" };
    });
  }
  /* PDF ya abierto → imágenes o PowerPoint */
  function fromOpenPdf(pdf, base, target, report) {
    if (target === "pptx") {
      return eachPage(pdf, function (page) { return renderPage(page, 150, "image/jpeg", 0.88, 2400); }, phase(report, 0.05, 0.85, "Renderizando página")).then(function (pages) {
        return Promise.all(pages.map(function (p) { return readAsArrayBuffer(p.blob).then(function (b) { return { norm: { bytes: new Uint8Array(b), mime: "image/jpeg", w: p.w, h: p.h } }; }); }));
      }).then(function (slides) {
        return buildPptx(slides, base, phase(report, 0.85, 1, "Creando PowerPoint")).then(function (blob) {
          return { blob: blob, name: base + ".pptx", note: slides.length + " diapositivas (como imagen)" };
        });
      });
    }
    return pdfToImages(pdf, target, phase(report, 0.05, 0.9, "Renderizando página")).then(function (pages) {
      return pagesResult(pages, base, target, phase(report, 0.9, 1, "Empaquetando"));
    });
  }
  /* Un PDF en memoria (Blob) → imágenes/PPT, reutilizando pdf.js */
  function viaPdf(pdfBlob, base, target, report) {
    return readAsArrayBuffer(pdfBlob).then(function (buf) { return openPdf(new Uint8Array(buf)); }).then(function (pdf) {
      return fromOpenPdf(pdf, base, target, report).then(function (r) { closePdf(pdf); return r; });
    });
  }

  function convertJob(job, report) {
    var kind = job.kind, target = job.target, base = baseName(job.file.name);
    var file = job.file;

    // ---------- IMÁGENES ----------
    if (IMAGE_KINDS[kind]) {
      if (IMAGE_KINDS[target]) {
        report(0.1, "Leyendo imagen");
        return decodeImage(file).then(function (img) {
          report(0.45, "Convirtiendo");
          var c = imageToCanvas(img, { whiteBg: target === "jpg" });
          if (img.close) img.close();
          return canvasToBlob(c, MIME[target], 0.92).then(function (blob) {
            var scaled = c._scaled; c.width = c.height = 0;
            if (blob.type && blob.type !== MIME[target]) throw UserError("Tu navegador no puede crear " + target.toUpperCase(), "Prueba con otro formato o con Chrome, Edge o Firefox.");
            return { blob: blob, name: base + "." + target, note: scaled ? "Reducida para caber en la memoria del navegador" : "" };
          });
        });
      }
      if (target === "pdf") return imagesToPdf([file], phase(report, 0.05, 1, "Creando PDF")).then(function (blob) { return { blob: blob, name: base + ".pdf" }; });
      if (target === "docx") return buildDocx({ title: base, blocks: [{ type: "image", blob: file, full: true }] }, phase(report, 0.05, 1, "Creando Word")).then(function (blob) { return { blob: blob, name: base + ".docx" }; });
      if (target === "pptx") {
        return normalizeImage(file, { maxSide: 3000 }).then(function (norm) {
          return buildPptx([{ norm: norm }], base, phase(report, 0.3, 1, "Creando PowerPoint"));
        }).then(function (blob) { return { blob: blob, name: base + ".pptx" }; });
      }
    }

    return readAsArrayBuffer(file).then(function (buf) {
      // ---------- PDF ----------
      if (kind === "pdf") {
        report(0.02, "Abriendo PDF");
        return openPdf(new Uint8Array(buf), job.password).then(function (pdf) {
          var done = function (r) { closePdf(pdf); return r; };
          if (target === "docx") {
            return pdfToModel(pdf, phase(report, 0.05, 0.8, "Leyendo página")).then(function (res) {
              res.model.title = base;
              return buildDocx(res.model, phase(report, 0.8, 1, "Creando Word")).then(function (blob) {
                var note = res.pages + (res.pages === 1 ? " página" : " páginas") + (res.scanned ? " · " + res.scanned + " escaneada(s) insertada(s) como imagen" : "");
                return { blob: blob, name: base + ".docx", note: note };
              });
            }).then(done);
          }
          if (target === "xlsx") {
            return pdfToSheets(pdf, phase(report, 0.05, 0.85, "Leyendo página")).then(function (res) {
              if (!res.sheets.length) throw UserError("El PDF no tiene texto que pasar a Excel", "Parece un documento escaneado (solo imágenes).");
              return buildXlsx(res.sheets, phase(report, 0.85, 1, "Creando Excel")).then(function (blob) {
                return { blob: blob, name: base + ".xlsx", note: res.sheets.length + (res.sheets.length === 1 ? " hoja" : " hojas") + (res.empty ? " · " + res.empty + " página(s) sin texto omitida(s)" : "") };
              });
            }).then(done);
          }
          return fromOpenPdf(pdf, base, target, report).then(done);
        });
      }

      // ---------- WORD ----------
      if (kind === "docx") {
        return docxToModel(buf, phase(report, 0, 0.3, "Leyendo Word")).then(function (res) {
          res.model.title = base;
          if (target === "xlsx") return buildXlsx(modelToSheets(res.model), phase(report, 0.3, 1, "Creando Excel")).then(function (blob) { return { blob: blob, name: base + ".xlsx" }; });
          var toPdf = modelToPdf(res.model, phase(report, 0.3, target === "pdf" ? 1 : 0.55, "Creando PDF"));
          if (target === "pdf") return toPdf.then(function (blob) { return { blob: blob, name: base + ".pdf" }; });
          return toPdf.then(function (pdfBlob) { return viaPdf(pdfBlob, base, target, phase(report, 0.55, 1, "")); });
        });
      }

      // ---------- EXCEL ----------
      if (kind === "xlsx") {
        return xlsxToModel(buf, phase(report, 0, 0.25, "Leyendo Excel")).then(function (res) {
          res.model.title = base;
          var note = res.sheets + (res.sheets === 1 ? " hoja" : " hojas") + (res.truncated ? " · recortado a " + BRAND.limits.maxSheetRows.toLocaleString("es") + " filas por hoja" : "");
          if (target === "docx") return buildDocx(res.model, phase(report, 0.25, 1, "Creando Word")).then(function (blob) { return { blob: blob, name: base + ".docx", note: note }; });
          var toPdf = modelToPdf(res.model, phase(report, 0.25, target === "pdf" ? 1 : 0.55, "Creando PDF"));
          if (target === "pdf") return toPdf.then(function (blob) { return { blob: blob, name: base + ".pdf", note: note }; });
          return toPdf.then(function (pdfBlob) { return viaPdf(pdfBlob, base, target, phase(report, 0.55, 1, "")); });
        });
      }

      // ---------- POWERPOINT ----------
      if (kind === "pptx") {
        return pptxToSlides(buf, phase(report, 0, 0.35, "Leyendo diapositiva")).then(function (res) {
          var bits = [res.slides.length + " diapositivas"];
          if (res.notes.charts) bits.push(res.notes.charts + " gráfico(s) no incluido(s)");
          if (res.notes.skippedImages) bits.push(res.notes.skippedImages + " imagen(es) en formato no compatible");
          var note = bits.join(" · ");
          if (target === "docx") {
            var m = slidesToModel(res.slides); m.title = base;
            return buildDocx(m, phase(report, 0.35, 1, "Creando Word")).then(function (blob) { return { blob: blob, name: base + ".docx", note: note }; });
          }
          var toPdf = slidesToPdf(res.slides, phase(report, 0.35, target === "pdf" ? 1 : 0.6, "Creando PDF"));
          if (target === "pdf") return toPdf.then(function (blob) { return { blob: blob, name: base + ".pdf", note: note }; });
          return toPdf.then(function (pdfBlob) { return viaPdf(pdfBlob, base, target, phase(report, 0.6, 1, "")); }).then(function (r) { r.note = note; return r; });
        });
      }
      throw UserError("Conversión no disponible", "Elige otro formato de salida.");
    });
  }

  /* ===========================================================
     13. INTERFAZ: AVISOS (toasts)
     =========================================================== */
  var ICON_X = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round"><path d="M6 6l12 12M18 6 6 18"/></svg>';
  function toast(type, title, text, ms) {
    var host = $("[data-toasts]");
    if (!host) return;
    var el = document.createElement("div");
    el.className = "toast toast--" + type;
    el.innerHTML = "<div><strong>" + escHTML(title) + "</strong>" + (text ? "<p>" + escHTML(text) + "</p>" : "") + "</div>" +
      '<button class="icon-btn" type="button" aria-label="Cerrar aviso">' + ICON_X + "</button>";
    var close = function () { if (!el.isConnected) return; el.classList.add("is-leaving"); setTimeout(function () { el.remove(); }, 300); };
    el.querySelector("button").addEventListener("click", close);
    host.appendChild(el);
    while (host.children.length > 4) host.firstElementChild.remove();
    setTimeout(close, ms || (type === "error" ? 9000 : 5000));
  }

  /* ===========================================================
     14. INTERFAZ: LISTA DE ARCHIVOS, BOTÓN «CONVERTIR A» Y COLA
     -----------------------------------------------------------
     Flujo único y simple: el usuario sube archivos de cualquier tipo →
     aparece el botón «Convertir a» con la lista de formatos posibles para
     lo que subió → elige uno y se convierte todo → descarga.
     =========================================================== */
  var jobs = [], seq = 0, running = 0, batchActive = false;
  var listEl, toolCard, convertBar;

  /* Configuración de la página (cada herramienta propone su formato) */
  var PAGE = { target: null }, dzTitleHTML = "";
  var ORDER = ["pdf", "docx", "xlsx", "pptx", "jpg", "png", "webp"];
  var TARGET_NAME = { pdf: "PDF", docx: "Word", xlsx: "Excel", pptx: "PowerPoint", png: "PNG", jpg: "JPG", webp: "WebP" };
  var TARGET_DESC = {
    pdf: "Listo para enviar o imprimir", docx: "Documento editable (.docx)", xlsx: "Hoja de cálculo (.xlsx)",
    pptx: "Presentación (.pptx)", jpg: "Imagen liviana, ideal para fotos", png: "Imagen sin pérdida, con transparencia",
    webp: "Imagen moderna y muy liviana"
  };
  var TAG = { pdf: "PDF", docx: "DOC", xlsx: "XLS", pptx: "PPT", png: "PNG", jpg: "JPG", webp: "WEBP" };

  function limitFor(kind) { return kind === "pdf" ? BRAND.limits.pdf : IMAGE_KINDS[kind] ? BRAND.limits.image : BRAND.limits.office; }
  function isBusy(j) { return j.state === "queued" || j.state === "working"; }
  function anyBusy() { return running > 0 || jobs.some(isBusy); }

  function addFiles(list) {
    var files = Array.prototype.slice.call(list || []);
    if (!files.length) return;
    Promise.all(files.map(function (f) {
      return detectKind(f).then(function (d) { return { file: f, det: d }; }).catch(function () { return { file: f, det: { kind: null, reason: "No se pudo leer el archivo." } }; });
    })).then(function (res) {
      var added = 0;
      res.forEach(function (r) {
        var f = r.file, d = r.det;
        if (!d.kind) { toast("error", d.encrypted ? "«" + f.name + "» está protegido" : "«" + f.name + "» no es compatible", d.reason); return; }
        if (MODE === "merge" && d.kind !== "pdf") { toast("error", "«" + f.name + "» no es un PDF", "Para unir solo se aceptan PDF. Para mezclar imágenes o Word usa «Crear PDF»."); return; }
        if (!f.size) { toast("error", "«" + f.name + "» está vacío", "El archivo no tiene contenido."); return; }
        if (f.size > limitFor(d.kind)) { toast("error", "«" + f.name + "» es demasiado grande", "El máximo para este tipo es ~" + fmtBytes(limitFor(d.kind)) + " para que tu navegador no se quede sin memoria."); return; }
        var job = { id: ++seq, file: f, kind: d.kind, target: null, state: "ready", progress: 0, label: "", result: null, error: null, password: "", skipMsg: "" };
        jobs.push(job);
        listEl.appendChild(rowFor(job));
        mergeResult = null;
        added++;
      });
      if (MODE !== "convert") jobs.forEach(paint);
      refreshChrome();
      if (added) {
        var b = $("[data-merge-run]:not([hidden])") || $("[data-convert-main]:not([hidden])") || $("[data-convert-toggle]");
        if (b && !b.disabled) try { b.focus({ preventScroll: true }); } catch (e) { /* sin importancia */ }
      }
    });
  }

  function rowFor(job) {
    var li = document.createElement("li");
    li.className = "file-row";
    li.dataset.id = job.id;
    li.innerHTML =
      '<span class="c-icon"><span class="ficon ficon--' + job.kind + '" aria-label="' + KIND_LABEL[job.kind] + '">' + KIND_LABEL[job.kind] + "</span></span>" +
      '<div class="c-name"><div class="fname">' + escHTML(job.file.name) + '</div><div class="fnote" data-note></div>' +
        '<div class="pw-box" data-pw hidden><input type="password" placeholder="Contraseña del PDF" autocomplete="off" aria-label="Contraseña del PDF">' +
        '<button class="btn btn--ink btn--sm" type="button" data-unlock>Desbloquear</button></div></div>' +
      '<span class="c-size fsize">' + fmtBytes(job.file.size) + "</span>" +
      '<div class="c-progress"><div class="progress" role="progressbar" aria-label="Progreso" aria-valuemin="0" aria-valuemax="100" aria-valuenow="0"><i></i></div><div class="pct" data-pct></div></div>' +
      '<div class="c-action row-action">' + (MODE !== "convert"
        ? '<span class="order-btns"><button class="icon-btn" type="button" data-up aria-label="Subir ' + escHTML(job.file.name) + '">' + ICON_UP + '</button><button class="icon-btn" type="button" data-down aria-label="Bajar ' + escHTML(job.file.name) + '">' + ICON_DN + "</button></span>"
        : '<button class="btn btn--sm" type="button" data-action hidden></button>') + "</div>" +
      '<span class="c-rm"><button class="icon-btn" type="button" data-remove aria-label="Quitar ' + escHTML(job.file.name) + '">' + ICON_X + "</button></span>";
    job.el = li;
    paint(job);
    return li;
  }

  var ICON_UP = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m6 15 6-6 6 6"/></svg>';
  var ICON_DN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="m6 9 6 6 6-6"/></svg>';
  var ICON_DOWN = '<svg viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.4" stroke-linecap="round" stroke-linejoin="round"><path d="M12 4v11M7 10l5 5 5-5M5 20h14"/></svg>';

  /* Pinta el estado de una fila (único sitio que toca su DOM). */
  function paint(job) {
    var li = job.el;
    if (!li) return;
    var busy = isBusy(job);
    li.className = "file-row row--" + job.state;
    var btn = $("[data-action]", li), rm = $("[data-remove]", li);
    var bar = $(".progress", li), pct = $("[data-pct]", li), note = $("[data-note]", li), pw = $("[data-pw]", li);
    rm.disabled = busy;
    var p = Math.round((job.progress || 0) * 100);
    var full = job.state === "done" || job.state === "error" || job.state === "skip";
    bar.firstElementChild.style.width = (full ? 100 : p) + "%";
    bar.setAttribute("aria-valuenow", job.state === "done" ? 100 : p);
    var to = job.target ? TARGET_NAME[job.target] : "";
    if (job.state === "working") pct.textContent = p + " % · " + (job.label || "Convirtiendo");
    else if (job.state === "queued") pct.textContent = "En cola · a " + to;
    else if (job.state === "done") pct.textContent = "Listo";
    else if (job.state === "error") pct.textContent = "No se pudo convertir";
    else if (job.state === "skip") pct.textContent = "Omitido";
    else pct.textContent = "Listo para convertir";
    if (MODE !== "convert") return paintMergeRow(job, pct, note, pw, busy);
    btn.hidden = !(job.state === "done" || job.state === "error");
    if (job.state === "done") {
      btn.className = "btn btn--ok btn--sm";
      btn.innerHTML = ICON_DOWN + "Descargar <span>(" + fmtBytes(job.result.blob.size) + ")</span>";
      note.className = "fnote is-ok"; note.textContent = "→ " + job.result.name + (job.result.note ? " · " + job.result.note : "");
    } else if (job.state === "error") {
      btn.className = "btn btn--soft btn--sm"; btn.textContent = "Reintentar";
      note.className = "fnote is-err"; note.textContent = job.error ? job.error.title + ". " + job.error.detail : "";
    } else if (job.state === "skip") {
      note.className = "fnote"; note.textContent = job.skipMsg;
    } else {
      note.className = "fnote"; note.textContent = job.password && job.state === "ready" ? "🔓 Contraseña lista" : "";
    }
    pw.hidden = !(job.state === "error" && job.error && job.error.code === "password");
  }

  function paintMergeRow(job, pct, note, pw, busy) {
    var pos = jobs.indexOf(job);
    var up = $("[data-up]", job.el), dn = $("[data-down]", job.el);
    up.disabled = busy || anyBusy() || pos <= 0; dn.disabled = busy || anyBusy() || pos >= jobs.length - 1;
    var p = Math.round((job.progress || 0) * 100);
    if (job.state === "working") pct.textContent = p + " % · " + (job.label || "Procesando");
    else if (job.state === "queued") pct.textContent = "En espera";
    else if (job.state === "done") pct.textContent = "Añadido · " + job.pages + (job.pages === 1 ? " página" : " páginas");
    else if (job.state === "error") pct.textContent = "No se pudo añadir";
    else pct.textContent = "Posición " + (pos + 1);
    if (job.state === "error") { note.className = "fnote is-err"; note.textContent = job.error ? job.error.title + ". " + job.error.detail : ""; }
    else { note.className = "fnote"; note.textContent = job.kind === "pdf" ? "" : "Se convertirá a PDF antes de unirlo"; }
    pw.hidden = true;
  }

  /* Actualización de progreso agrupada por fotograma (no satura el navegador). */
  var dirty = new Set(), rafPending = false;
  function schedulePaint(job) {
    dirty.add(job);
    if (rafPending) return;
    rafPending = true;
    var flush = function () { if (!rafPending) return; rafPending = false; dirty.forEach(paint); dirty.clear(); };
    // En una pestaña en segundo plano el navegador pausa los fotogramas: se usa un temporizador de respaldo.
    if (document.hidden) setTimeout(flush, 250);
    else { requestAnimationFrame(flush); setTimeout(flush, 500); }
  }

  function enqueue(job) {
    if (isBusy(job)) return;
    job.state = "queued"; job.progress = 0; job.result = null; job.error = null;
    paint(job);
    pump();
    refreshChrome();
  }
  function pump() {
    while (running < BRAND.concurrency) {
      var next = jobs.filter(function (j) { return j.state === "queued"; })[0];
      if (!next) break;
      run(next);
    }
    if (!running && batchActive) {
      batchActive = false;
      var ok = jobs.filter(function (j) { return j.state === "done"; }).length;
      var bad = jobs.filter(function (j) { return j.state === "error"; }).length;
      if (ok > 1 && !bad) toast("ok", "¡Listo! " + ok + " archivos convertidos", "Descárgalos uno a uno o todos juntos en un ZIP.");
      else if (bad) toast("error", "Conversión terminada con avisos", ok + " listo(s) · " + bad + " con error. Revisa el mensaje de cada archivo.");
    }
    refreshChrome();
  }
  function run(job) {
    running++;
    job.state = "working"; job.label = "Preparando el motor (solo la primera vez)"; job.progress = 0.01;
    paint(job);
    var last = 0;
    var report = function (f, label) {
      if (jobs.indexOf(job) < 0) return;
      if (f >= last) { job.progress = Math.min(0.99, f); last = f; }
      if (label) job.label = label;
      schedulePaint(job);
    };
    // Se cede el turno antes de empezar para que la interfaz pinte el estado «Convirtiendo».
    yieldUI().then(function () { return convertJob(job, report); }).then(function (res) {
      job.result = res; job.state = "done"; job.progress = 1;
    }).catch(function (err) {
      console.warn("[conversión] " + job.file.name, err);
      job.error = humanize(err, job.kind); job.state = "error";
      toast("error", job.error.title, "«" + job.file.name + "»: " + job.error.detail);
    }).then(function () {
      running--;
      if (jobs.indexOf(job) >= 0) paint(job);
      pump();
    });
  }

  function removeJob(job) {
    if (isBusy(job)) return;
    jobs = jobs.filter(function (j) { return j !== job; });
    if (job.el) job.el.remove();
    mergeResult = null;
    if (MODE !== "convert") jobs.forEach(paint);
    refreshChrome();
  }

  /* Formatos posibles para los archivos cargados (con cuántos sirve cada uno). */
  function convertOptions() {
    var counts = {};
    jobs.forEach(function (j) { targetsFor(j.kind).forEach(function (t) { counts[t] = (counts[t] || 0) + 1; }); });
    var opts = ORDER.filter(function (t) { return counts[t]; }).map(function (t) { return { t: t, n: counts[t] }; });
    if (PAGE.target) opts.sort(function (a, b) { return (b.t === PAGE.target) - (a.t === PAGE.target); });
    return opts;
  }
  function imageJobs() { return jobs.filter(function (j) { return IMAGE_KINDS[j.kind]; }); }

  function renderConvertList() {
    var host = $("[data-convert-list]");
    if (!host) return;
    var opts = convertOptions(), total = jobs.length, html = "<h4>Convertir a</h4>";
    opts.forEach(function (o) {
      var extra = o.n < total ? " · sirve para " + o.n + " de " + total + " archivos" : "";
      html += '<button class="convert-opt" type="button" role="menuitem" data-to="' + o.t + '">' +
        '<span class="ticon ticon--' + o.t + '">' + TAG[o.t] + "</span>" +
        "<span>" + TARGET_NAME[o.t] + "<small>" + TARGET_DESC[o.t] + extra + "</small></span>" +
        (o.t === PAGE.target ? '<span class="pill-sun" style="margin-left:auto">Recomendado</span>' : "") + "</button>";
    });
    var imgs = imageJobs().length;
    if (imgs >= 2) {
      html += "<h4>Varias imágenes</h4>" +
        '<button class="convert-opt" type="button" role="menuitem" data-merge><span class="ticon ticon--pdf">PDF</span>' +
        "<span>Un solo PDF con todas las imágenes<small>" + imgs + " imágenes, en el orden de la lista</small></span></button>";
    }
    host.innerHTML = html;
  }

  function setMenuOpen(open) {
    var host = $("[data-convert-list]"), toggle = $("[data-convert-toggle]");
    if (!host || !toggle) return;
    if (open) renderConvertList();
    host.hidden = !open;
    toggle.setAttribute("aria-expanded", open ? "true" : "false");
    if (open) { var first = $(".convert-opt", host); if (first) first.focus(); }
  }

  /* Convierte TODOS los archivos al formato elegido (los que no pueden, se omiten con un aviso). */
  function convertAllTo(t) {
    setMenuOpen(false);
    var any = false;
    jobs.forEach(function (j) {
      if (isBusy(j)) return;
      if (j.kind === t) { j.state = "skip"; j.skipMsg = "Ya es un archivo " + TARGET_NAME[t] + "."; paint(j); return; }
      if (targetsFor(j.kind).indexOf(t) < 0) { j.state = "skip"; j.skipMsg = "Un archivo " + KIND_LABEL[j.kind] + " no se puede convertir a " + TARGET_NAME[t] + "."; paint(j); return; }
      if (j.error && j.error.code === "password" && !j.password) { j.target = t; paint(j); return; }
      j.target = t; any = true;
      j.state = "ready";
      enqueue(j);
    });
    if (any) batchActive = true;
    refreshChrome();
  }

  function mergeImages() {
    setMenuOpen(false);
    var imgs = imageJobs();
    if (imgs.length < 2) return;
    var main = $("[data-convert-toggle]");
    convertBar.dataset.merging = "1";
    refreshChrome();
    var label = $("[data-convert-label]");
    imagesToPdf(imgs.map(function (j) { return j.file; }), function (f) {
      if (label) label.textContent = "Uniendo… " + Math.round(f * 100) + " %";
    }).then(function (blob) {
      var name = "imagenes-unidas.pdf";
      toast("ok", "PDF creado con " + imgs.length + " imágenes", "Se está descargando «" + name + "» (" + fmtBytes(blob.size) + ").");
      saveBlob(blob, name);
      afterDownload("Descargando «" + name + "»");
    }).catch(function (e) {
      var h = humanize(e, "imagen"); toast("error", h.title, h.detail);
    }).then(function () { delete convertBar.dataset.merging; refreshChrome(); if (main) main.focus(); });
  }

  /* Estado de la barra «Convertir a» y de la tarjeta según la lista. */
  function refreshChrome() {
    var has = jobs.length > 0;
    toolCard.classList.toggle("has-files", has);
    convertBar.hidden = !has;
    $("[data-dz-title]").innerHTML = has ? "Añadir más archivos" : dzTitleHTML;
    if (!has) return;
    var busy = anyBusy() || convertBar.dataset.merging === "1";
    var done = jobs.filter(function (j) { return j.state === "done"; });
    $("[data-files-count]").textContent = jobs.length === 1 ? "1 archivo" : jobs.length + " archivos";
    $("[data-clear-all]").disabled = busy;
    var zip = $("[data-zip-all]");
    zip.hidden = done.length < 2; zip.disabled = busy;
    if (MODE !== "convert") { refreshMergeChrome(busy); return; }
    var opts = convertOptions();
    var main = $("[data-convert-main]"), toggle = $("[data-convert-toggle]"), label = $("[data-convert-label]");
    var pageOk = PAGE.target && opts.some(function (o) { return o.t === PAGE.target; });
    main.hidden = !pageOk;
    toggle.classList.toggle("btn--caret", !!pageOk);
    toggle.setAttribute("aria-label", pageOk ? "Elegir otro formato" : "Convertir a");
    main.disabled = toggle.disabled = busy || !opts.length;
    if (busy) {
      if (convertBar.dataset.merging !== "1") {
        if (pageOk) main.innerHTML = '<span class="spinner" aria-hidden="true"></span>Convirtiendo…';
        else label.textContent = "Convirtiendo…";
      }
    } else {
      if (pageOk) main.textContent = "Convertir a " + TARGET_NAME[PAGE.target];
      label.textContent = pageOk ? "" : "Convertir a";
    }
    $("[data-convert-caret-text]").hidden = !!pageOk;
    if (!$("[data-convert-list]").hidden) renderConvertList();
  }

  /* ===========================================================
     14b. HERRAMIENTAS «UNIR PDF» Y «CREAR PDF»
     -----------------------------------------------------------
     En estas páginas no se elige formato: los archivos se ordenan con
     las flechas y se juntan en UN solo PDF, en el orden de la lista.
       · merge  → solo PDF (se copian las páginas originales, texto incluido)
       · build  → imágenes, Word, Excel, PowerPoint y PDF: lo que no es PDF
                  se convierte primero y después se une todo.
     El motor pdf-lib 1.17.1 (vendorizado) copia las páginas sin rasterizar.
     =========================================================== */
  var MODE = "convert", mergeResult = null;

  function moveJob(job, dir) {
    var i = jobs.indexOf(job), j = i + dir;
    if (i < 0 || j < 0 || j >= jobs.length || anyBusy()) return;
    jobs.splice(i, 1); jobs.splice(j, 0, job);
    if (dir < 0) listEl.insertBefore(job.el, jobs[j + 1].el);
    else listEl.insertBefore(job.el, jobs[j - 1].el.nextSibling);
    mergeResult = null;
    jobs.forEach(paint);
    refreshChrome();
    var b = $(dir < 0 ? "[data-up]" : "[data-down]", job.el);
    if (b && !b.disabled) b.focus();
  }

  /* Un archivo cualquiera → bytes de un PDF (los PDF pasan tal cual). */
  function toPdfBytes(job, report) {
    var f = job.file, kind = job.kind;
    if (kind === "pdf") return readAsArrayBuffer(f);
    var pdfBlob;
    if (IMAGE_KINDS[kind]) pdfBlob = imagesToPdf([f], report);
    else pdfBlob = readAsArrayBuffer(f).then(function (buf) {
      if (kind === "docx") return docxToModel(buf, function (x) { report(x * 0.4); }).then(function (r) { return modelToPdf(r.model, function (x) { report(0.4 + x * 0.6); }); });
      if (kind === "xlsx") return xlsxToModel(buf, function (x) { report(x * 0.3); }).then(function (r) { return modelToPdf(r.model, function (x) { report(0.3 + x * 0.7); }); });
      if (kind === "pptx") return pptxToSlides(buf, function (x) { report(x * 0.4); }).then(function (r) { return slidesToPdf(r.slides, function (x) { report(0.4 + x * 0.6); }); });
      throw UserError("Formato no compatible", "Este archivo no se puede añadir al PDF.");
    });
    return pdfBlob.then(readAsArrayBuffer);
  }

  function runMerge() {
    if (anyBusy() || !jobs.length) return;
    if (MODE === "merge" && jobs.length < 2) { toast("info", "Añade al menos 2 PDF", "Para unir necesitas dos o más archivos."); return; }
    mergeResult = null;
    jobs.forEach(function (j) { j.state = "queued"; j.progress = 0; j.error = null; j.label = ""; paint(j); });
    running++;
    refreshChrome();
    var label = $("[data-merge-run]");
    var PDFLib, out, total = jobs.length, i = 0, failed = [];
    Engines.pdflib().then(function (lib) {
      PDFLib = lib;
      return PDFLib.PDFDocument.create();
    }).then(function (doc) {
      out = doc;
      function next() {
        if (i >= jobs.length) return Promise.resolve();
        var job = jobs[i++];
        job.state = "working"; job.label = job.kind === "pdf" ? "Leyendo PDF" : "Convirtiendo a PDF"; job.progress = 0.05; paint(job);
        if (label) label.innerHTML = '<span class="spinner" aria-hidden="true"></span>Procesando ' + i + " de " + total + "…";
        var report = function (f) { job.progress = 0.05 + Math.min(0.9, (f || 0) * 0.9); schedulePaint(job); };
        return toPdfBytes(job, report).then(function (bytes) {
          job.label = "Añadiendo páginas"; schedulePaint(job);
          return PDFLib.PDFDocument.load(bytes, { updateMetadata: false });
        }).then(function (src) {
          return out.copyPages(src, src.getPageIndices()).then(function (pages) {
            pages.forEach(function (p) { out.addPage(p); });
            job.state = "done"; job.progress = 1;
            job.pages = pages.length;
            paint(job);
          });
        }).catch(function (err) {
          console.warn("[unir] " + job.file.name, err);
          var msg = String((err && err.message) || "");
          job.error = /encrypt/i.test(msg) || (err && err.name === "EncryptedPDFError")
            ? UserError("PDF protegido con contraseña", "Ábrelo, guárdalo sin contraseña y vuelve a añadirlo.")
            : humanize(err, job.kind);
          job.state = "error"; paint(job); failed.push(job);
        }).then(yieldUI).then(next);
      }
      return next();
    }).then(function () {
      if (!out.getPageCount()) throw UserError("No se pudo crear el PDF", "Ninguno de los archivos se pudo añadir.");
      if (label) label.innerHTML = '<span class="spinner" aria-hidden="true"></span>Guardando el PDF…';
      out.setProducer(BRAND.name); out.setCreator(BRAND.name);
      return yieldUI().then(function () { return out.save(); });
    }).then(function (bytes) {
      var base = baseName(jobs[0].file.name);
      var name = (MODE === "merge" ? base + "-unido" : base + "-completo") + ".pdf";
      mergeResult = { blob: new Blob([bytes], { type: MIME.pdf }), name: name, pages: out.getPageCount() };
      if (failed.length) toast("error", "PDF creado sin " + failed.length + " archivo(s)", "Revisa el mensaje en cada archivo marcado en rojo.");
      else toast("ok", "¡PDF listo!", mergeResult.pages + " páginas en «" + name + "». Pulsa «Descargar».");
    }).catch(function (err) {
      var h = humanize(err, "pdf"); toast("error", h.title, h.detail);
    }).then(function () {
      running--;
      refreshChrome();
      var dl = $("[data-merge-result]");
      if (dl && !dl.hidden) try { dl.focus({ preventScroll: true }); } catch (e) { /* sin importancia */ }
    });
  }

  /* Partes de la interfaz que cambian en estas dos herramientas */
  function refreshMergeChrome(busy) {
    var run = $("[data-merge-run]"), res = $("[data-merge-result]");
    $(".convert-menu").hidden = true;
    $("[data-zip-all]").hidden = true;
    run.hidden = false;
    run.disabled = busy || !jobs.length || (MODE === "merge" && jobs.length < 2);
    if (!busy) run.textContent = MODE === "merge" ? "Unir " + jobs.length + " PDF" : "Crear PDF con " + (jobs.length === 1 ? "1 archivo" : jobs.length + " archivos");
    res.hidden = !mergeResult || busy;
    if (mergeResult) res.innerHTML = ICON_DOWN + "Descargar PDF <span>(" + mergeResult.pages + " págs. · " + fmtBytes(mergeResult.blob.size) + ")</span>";
  }

  /* ===========================================================
     15. PUBLICIDAD: pop-up tras descargar + notificación de esquina
     =========================================================== */
  var lastModal = 0;
  function afterDownload(text) {
    var cfg = BRAND.ads;
    if (!cfg.modalOnDownload) return;
    var now = Date.now();
    if (cfg.modalCooldownSec && now - lastModal < cfg.modalCooldownSec * 1000) return;
    lastModal = now;
    // La descarga YA empezó; el anuncio aparece justo después y nunca la bloquea.
    setTimeout(function () { openModal(text); }, cfg.modalDelayMs);
  }
  function openModal(text) {
    var d = $("[data-ad-modal]");
    if (!d) return;
    $("[data-modal-text]", d).textContent = text || "Tu descarga ha comenzado";
    try { if (typeof d.showModal === "function") { if (!d.open) d.showModal(); } else d.setAttribute("open", ""); }
    catch (e) { d.setAttribute("open", ""); }
  }
  function initAds() {
    var d = $("[data-ad-modal]");
    if (d) {
      var close = function () { try { d.close(); } catch (e) { d.removeAttribute("open"); } };
      $$("[data-modal-close]", d).forEach(function (b) { b.addEventListener("click", close); });
      d.addEventListener("click", function (e) { if (e.target === d) close(); }); // clic fuera de la tarjeta
      d.addEventListener("cancel", function () { /* Esc cierra (comportamiento nativo) */ });
    }
    var corner = $("[data-corner-ad]");
    if (!corner) return;
    var dismissed = false;
    try { dismissed = sessionStorage.getItem("cl_corner_closed") === "1"; } catch (e) {}
    if (dismissed) return;
    setTimeout(function () {
      corner.hidden = false;
      requestAnimationFrame(function () { requestAnimationFrame(function () { corner.classList.add("is-visible"); }); });
    }, BRAND.ads.cornerDelayMs);
    $("[data-corner-close]", corner).addEventListener("click", function () {
      corner.classList.remove("is-visible");
      setTimeout(function () { corner.hidden = true; }, 450);
      try { sessionStorage.setItem("cl_corner_closed", "1"); } catch (e) {}
    });
  }

  /* ===========================================================
     16. EVENTOS DE LA HERRAMIENTA
     =========================================================== */
  function initTool() {
    toolCard = $("[data-tool]"); listEl = $("[data-rows]"); convertBar = $("[data-convert-bar]");
    var dz = $("[data-dropzone]"), input = $("[data-file-input]");
    PAGE.target = toolCard.dataset.target || null;
    MODE = toolCard.dataset.mode || "convert";
    dzTitleHTML = $("[data-dz-title]").innerHTML;
    // Safari no puede crear WebP: en las páginas «a WebP» se avisa con claridad en lugar de fallar.
    if (PAGE.target === "webp" && !canExportWebP) {
      var warn = $("[data-webp-warning]");
      if (warn) warn.hidden = false;
    }

    input.addEventListener("change", function () { addFiles(input.files); input.value = ""; });
    var depth = 0;
    ["dragenter", "dragover"].forEach(function (ev) {
      document.addEventListener(ev, function (e) {
        if (!e.dataTransfer || Array.prototype.indexOf.call(e.dataTransfer.types || [], "Files") < 0) return;
        e.preventDefault();
        if (ev === "dragenter") depth++;
        dz.classList.add("is-over");
      });
    });
    document.addEventListener("dragleave", function () { depth = Math.max(0, depth - 1); if (!depth) dz.classList.remove("is-over"); });
    document.addEventListener("drop", function (e) {
      if (!e.dataTransfer || !e.dataTransfer.files || !e.dataTransfer.files.length) return;
      e.preventDefault(); depth = 0; dz.classList.remove("is-over");
      addFiles(e.dataTransfer.files);
    });
    document.addEventListener("paste", function (e) {
      var items = (e.clipboardData && e.clipboardData.files) || [];
      if (items.length) { e.preventDefault(); addFiles(items); }
    });

    // Lista de archivos: descargar, reintentar, quitar, desbloquear
    listEl.addEventListener("click", function (e) {
      var job = jobOf(e.target);
      if (!job) return;
      if (e.target.closest("[data-remove]")) return removeJob(job);
      if (e.target.closest("[data-up]")) return moveJob(job, -1);
      if (e.target.closest("[data-down]")) return moveJob(job, 1);
      if (e.target.closest("[data-unlock]")) {
        var inp = $("[data-pw] input", job.el);
        if (!inp.value) { inp.focus(); return; }
        job.password = inp.value; inp.value = "";
        job.state = "ready"; job.error = null;
        if (job.target) { batchActive = true; enqueue(job); } else paint(job);
        return;
      }
      if (e.target.closest("[data-action]")) {
        if (job.state === "done") {
          saveBlob(job.result.blob, job.result.name);
          afterDownload("Descargando «" + job.result.name + "»");
        } else if (job.state === "error" && job.target) {
          enqueue(job);
        }
      }
    });
    listEl.addEventListener("keydown", function (e) {
      if (e.key === "Enter" && e.target.matches("[data-pw] input")) { e.preventDefault(); $("[data-unlock]", e.target.closest("[data-id]")).click(); }
    });

    // Botón «Convertir a» y su lista
    $("[data-convert-main]").addEventListener("click", function () { if (PAGE.target) convertAllTo(PAGE.target); });
    var mr = $("[data-merge-run]"), md = $("[data-merge-result]");
    if (mr) mr.addEventListener("click", runMerge);
    if (md) md.addEventListener("click", function () {
      if (!mergeResult) return;
      saveBlob(mergeResult.blob, mergeResult.name);
      afterDownload("Descargando «" + mergeResult.name + "»");
    });
    $("[data-convert-toggle]").addEventListener("click", function () { setMenuOpen($("[data-convert-list]").hidden); });
    $("[data-convert-list]").addEventListener("click", function (e) {
      var opt = e.target.closest("[data-to]");
      if (opt) return convertAllTo(opt.dataset.to);
      if (e.target.closest("[data-merge]")) mergeImages();
    });
    $("[data-convert-list]").addEventListener("keydown", function (e) {
      var items = $$(".convert-opt", e.currentTarget), i = items.indexOf(document.activeElement);
      if (e.key === "ArrowDown") { e.preventDefault(); (items[i + 1] || items[0]).focus(); }
      if (e.key === "ArrowUp") { e.preventDefault(); (items[i - 1] || items[items.length - 1]).focus(); }
    });
    document.addEventListener("keydown", function (e) {
      if (e.key === "Escape" && !$("[data-convert-list]").hidden) { setMenuOpen(false); $("[data-convert-toggle]").focus(); }
    });
    document.addEventListener("click", function (e) {
      var host = $("[data-convert-list]");
      if (host.hidden || e.target.closest(".convert-menu")) return;
      setMenuOpen(false);
    });

    $("[data-zip-all]").addEventListener("click", function () {
      var done = jobs.filter(function (j) { return j.state === "done"; });
      if (!done.length) return;
      var btn = this, old = btn.innerHTML;
      btn.disabled = true; btn.innerHTML = '<span class="spinner" aria-hidden="true"></span>Preparando ZIP…';
      blobsToZip(done.map(function (j) { return { blob: j.result.blob, name: j.result.name }; }), function (f) {
        btn.innerHTML = '<span class="spinner" aria-hidden="true"></span>Preparando ZIP… ' + Math.round(f * 100) + " %";
      }).then(function (zip) {
        saveBlob(zip, BRAND.slug + "-" + new Date().toISOString().slice(0, 10) + ".zip");
        afterDownload("Descargando " + done.length + " archivos en un ZIP");
      }).catch(function (e) {
        console.warn(e); toast("error", "No se pudo crear el ZIP", "Prueba a descargar los archivos uno por uno.");
      }).then(function () { btn.innerHTML = old; refreshChrome(); });
    });
    $("[data-clear-all]").addEventListener("click", function () {
      if (anyBusy()) return;
      jobs.forEach(function (j) { if (j.el) j.el.remove(); });
      jobs = [];
      mergeResult = null;
      setMenuOpen(false);
      refreshChrome();
    });
  }
  function jobOf(el) {
    var row = el.closest && el.closest("[data-id]");
    if (!row) return null;
    var id = +row.dataset.id;
    return jobs.filter(function (j) { return j.id === id; })[0] || null;
  }

  /* ===========================================================
     17. TEMA CLARO / OSCURO
     =========================================================== */
  function initTheme() {
    var root = document.documentElement, saved = null;
    try { saved = localStorage.getItem("cl_theme"); } catch (e) {}
    var sysDark = window.matchMedia && matchMedia("(prefers-color-scheme: dark)").matches;
    if (!root.dataset.theme) root.dataset.theme = saved || (sysDark ? "dark" : "light");
    var btn = $("[data-theme-toggle]");
    if (btn) btn.addEventListener("click", function () {
      var next = root.dataset.theme === "dark" ? "light" : "dark";
      root.dataset.theme = next;
      try { localStorage.setItem("cl_theme", next); } catch (e) {}
    });
  }

  /* ===========================================================
     18. ARRANQUE + RED DE SEGURIDAD GLOBAL
     =========================================================== */
  window.addEventListener("error", function (e) {
    console.warn("[error global]", e.error || e.message);
  });
  window.addEventListener("unhandledrejection", function (e) {
    // Solo se registra: los errores de conversión ya se muestran en su fila y
    // un aviso aquí podría dispararse por scripts ajenos (p. ej. anuncios).
    console.warn("[promesa sin capturar]", e.reason);
  });

  /* ===========================================================
     17b. MENÚ «HERRAMIENTAS» (tipo panel) Y MENÚ MÓVIL
     =========================================================== */
  function initMenu() {
    var btn = $("[data-menu-btn]"), panel = $("[data-mega]");
    if (!btn || !panel) return;
    var extra = $$("[data-menu-alt]");
    function setOpen(open) {
      panel.hidden = !open;
      btn.setAttribute("aria-expanded", open ? "true" : "false");
      extra.forEach(function (b) { b.setAttribute("aria-expanded", open ? "true" : "false"); });
    }
    function toggle(e) { e.preventDefault(); setOpen(panel.hidden); }
    btn.addEventListener("click", toggle);
    extra.forEach(function (b) { b.addEventListener("click", toggle); });
    document.addEventListener("keydown", function (e) { if (e.key === "Escape" && !panel.hidden) { setOpen(false); btn.focus(); } });
    document.addEventListener("click", function (e) {
      if (panel.hidden) return;
      if (panel.contains(e.target) || btn.contains(e.target) || extra.some(function (b) { return b.contains(e.target); })) return;
      setOpen(false);
    });
  }

  function boot() {
    safe(initTheme, "initTheme");
    safe(initMenu, "initMenu");
    if ($("[data-tool]")) safe(initTool, "initTool");
    safe(initAds, "initAds");
    safe(function () { var y = $("[data-year]"); if (y) y.textContent = new Date().getFullYear(); }, "year");
    // Navegadores muy antiguos: aviso claro en vez de botones muertos.
    if (!window.Promise || !window.Blob || !window.URL || !document.createElement("canvas").getContext) {
      toast("error", "Tu navegador es demasiado antiguo", "Actualiza Chrome, Edge, Firefox o Safari para usar el convertidor.", 20000);
    }
    document.documentElement.classList.add("is-ready");
  }
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", boot);
  else boot();
})();
  
