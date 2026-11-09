// ============================================================
// core.js — утилиты, иконки, цвета, элементарные UI-компоненты
// ============================================================
const {
  useState,
  useEffect,
  useRef,
  useCallback,
  useMemo
} = React;
function uid() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function nowISO() {
  return new Date().toISOString();
}
function escapeHtml(s) {
  return String(s || "").replace(/[&<>"']/g, c => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#39;"
  })[c]);
}
function stripHtml(html) {
  const d = document.createElement("div");
  d.innerHTML = html || "";
  return (d.textContent || "").replace(/\s+/g, " ").trim();
}

// Экспорт в .doc — тот же приём, что и в «Конструкторе уроков» (русский язык):
// HTML в MS-Office-обёртке, который Word открывает как обычный документ.
function wordDocFullHtml(titleText, bodyHtml, opts) {
  opts = opts || {};
  return `<html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
<head><meta charset="utf-8"><title>${escapeHtml(titleText)}</title>
<style>
body{font-family:Calibri,Arial,sans-serif;font-size:14px;line-height:1.5}
h1{font-family:Georgia,serif;font-size:20px}
.q{margin-bottom:16px}
.qnum{font-weight:bold}
.answer{color:#132c53;margin-top:4px}
.note{background:#FBF3E4;padding:6px 10px;border-radius:6px;margin-top:4px;font-size:12px}
${opts.pageCss || ""}
</style></head>
<body>${opts.noTitle ? "" : `<h1>${escapeHtml(titleText)}</h1>`}${opts.pageCss ? `<div class="Section1">${bodyHtml}</div>` : bodyHtml}</body></html>`;
}
function downloadWord(filename, titleText, bodyHtml, opts) {
  const html = wordDocFullHtml(titleText, bodyHtml, opts);
  const blob = new Blob(["\ufeff", html], {
    type: "application/msword"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
}

// Экспорт в PDF — тот же приём, что и в «Конструкторе уроков» (русский язык): прячем всё,
// кроме служебной области, и вызываем системную печать — пользователь сохраняет как PDF.
function ensurePrintArea() {
  if (document.getElementById("lc-print-area")) return;
  const div = document.createElement("div");
  div.id = "lc-print-area";
  document.body.appendChild(div);
  const style = document.createElement("style");
  style.textContent = `
    #lc-print-area { display: none; }
    @media print {
      body > *:not(#lc-print-area) { display: none !important; }
      #lc-print-area { display: block !important; }
      #lc-print-area * { visibility: visible !important; }
    }
  `;
  document.head.appendChild(style);
}
function printAsPdf(bodyHtml) {
  ensurePrintArea();
  document.getElementById("lc-print-area").innerHTML = bodyHtml;
  setTimeout(() => window.print(), 80);
}

/* ============================================================
   ПЕЧАТЬ ЛИСТАМИ: один и тот же материал повторяется N раз на листе A4 (карточки, которые режут и раздают).
   Сетка [столбцов, строк]; поля листа 8 мм. Перенесено из «Конструктора уроков» (русский язык).
   ============================================================ */
const CARD_LAYOUTS = {
  1: [1, 1],
  2: [1, 2],
  4: [2, 2],
  6: [2, 3],
  8: [2, 4]
};
const CARD_PAGE_W_MM = 194;
const CARD_PAGE_H_MM = 277;
function cardsPageHtml(contentHtml, copies, fontPx, clip) {
  if (copies === 1) return `<div class="lc-card-cell" style="font-size:${fontPx}px;line-height:1.4;${clip ? "max-height:277mm;overflow:hidden;" : ""}">${contentHtml}</div>`;
  const [cols, rows] = CARD_LAYOUTS[copies] || [1, 1];
  const cellW = CARD_PAGE_W_MM / cols,
    cellH = CARD_PAGE_H_MM / rows;
  const cell = `<td style="width:${cellW}mm;height:${cellH}mm;vertical-align:top;border:1px dashed #8a97a5;padding:3mm;"><div class="lc-card-cell" style="height:${(cellH - 7).toFixed(1)}mm;overflow:${clip ? "hidden" : "visible"};font-size:${fontPx}px;line-height:1.4;">${contentHtml}</div></td>`;
  const row = `<tr style="height:${cellH}mm">${cell.repeat(cols)}</tr>`;
  return `<table style="border-collapse:collapse;table-layout:fixed;width:${CARD_PAGE_W_MM}mm;">${row.repeat(rows)}</table>`;
}
function cardCellSize(copies) {
  const [cols, rows] = CARD_LAYOUTS[copies] || [1, 1];
  return {
    w: CARD_PAGE_W_MM / cols - 6,
    h: CARD_PAGE_H_MM / rows - 7
  };
}
// Картинка, вписанная в прямоугольник boxW×boxH мм (размер в px — так и Word, и печать понимают одинаково).
function fitImgHtml(src, size, boxWmm, boxHmm) {
  if (!size) return `<img src="${src}" style="max-width:100%;" />`;
  const mmToPx = 3.7795;
  const scale = Math.min(boxWmm * mmToPx / size.w, boxHmm * mmToPx / size.h);
  return `<img src="${src}" width="${Math.round(size.w * scale)}" height="${Math.round(size.h * scale)}" style="display:block;margin:0 auto;" />`;
}
const CARDS_WORD_PAGE_CSS = `@page Section1 { size: 210mm 297mm; margin: 8mm 8mm 8mm 8mm; mso-page-orientation: portrait; } div.Section1 { page: Section1; }`;
// Печать листа: поля 8 мм и размер A4 действуют только на время печати.
function printCardsPdf(bodyHtml) {
  ensurePrintArea();
  const old = document.getElementById("lc-cards-page-style");
  if (old) old.remove();
  const st = document.createElement("style");
  st.id = "lc-cards-page-style";
  st.textContent = "@page { size: A4 portrait; margin: 8mm; }";
  document.head.appendChild(st);
  document.getElementById("lc-print-area").innerHTML = `<div style="font-family:Arial,sans-serif;color:#000;">${bodyHtml}</div>`;
  const cleanup = () => {
    st.remove();
    window.removeEventListener("afterprint", cleanup);
  };
  window.addEventListener("afterprint", cleanup);
  setTimeout(() => window.print(), 120);
}
// Word не видит картинки по адресу litfile:// — для скачивания их встраиваем в документ (data:).
async function inlineLitfileImages(html) {
  const doc = new DOMParser().parseFromString("<body>" + html + "</body>", "text/html");
  const imgs = Array.from(doc.querySelectorAll('img[src^="litfile:"]'));
  for (const img of imgs) {
    try {
      const r = await fetch(img.getAttribute("src"));
      const blob = await r.blob();
      const dataUrl = await new Promise((res, rej) => {
        const fr = new FileReader();
        fr.onload = () => res(fr.result);
        fr.onerror = rej;
        fr.readAsDataURL(blob);
      });
      img.setAttribute("src", dataUrl);
    } catch (e) {/* оставляем как есть */}
  }
  return doc.body.innerHTML;
}
// Выбор «сколько раз на листе» и размера текста + предпросмотр листа; используется окнами печати материала и ДЗ.
function CardsOptions({
  copies,
  setCopies,
  fontPx,
  setFontPx,
  noFont
}) {
  const chip = active => ({
    cursor: "pointer",
    borderColor: active ? "var(--accent)" : undefined,
    background: active ? "var(--accent-soft)" : undefined,
    color: active ? "var(--accent)" : undefined,
    fontWeight: active ? 700 : 600,
    padding: "6px 12px"
  });
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontWeight: 700,
      marginBottom: 4
    }
  }, "\u0421\u043A\u043E\u043B\u044C\u043A\u043E \u0440\u0430\u0437 \u043D\u0430 \u043E\u0434\u043D\u043E\u043C \u043B\u0438\u0441\u0442\u0435 A4"), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      marginBottom: 10
    }
  }, Object.keys(CARD_LAYOUTS).map(n => /*#__PURE__*/React.createElement("button", {
    key: n,
    type: "button",
    className: "chip",
    style: chip(copies === +n),
    onClick: () => setCopies(+n),
    title: `${CARD_LAYOUTS[n][0]} × ${CARD_LAYOUTS[n][1]}`
  }, n))), !noFont && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontWeight: 700,
      marginBottom: 4
    }
  }, "\u0420\u0430\u0437\u043C\u0435\u0440 \u0442\u0435\u043A\u0441\u0442\u0430"), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6
    }
  }, [11, 12, 14, 16, 18, 22, 26].map(px => /*#__PURE__*/React.createElement("button", {
    key: px,
    type: "button",
    className: "chip",
    style: chip(fontPx === px),
    onClick: () => setFontPx(px)
  }, px)))));
}
function CardsPreview({
  html,
  copies,
  fontPx,
  onOverflow
}) {
  const ref = useRef(null);
  useEffect(() => {
    const cells = ref.current ? Array.from(ref.current.querySelectorAll(".lc-card-cell")) : [];
    onOverflow(copies > 1 && cells.some(c => c.scrollHeight > c.clientHeight + 2));
  }, [html, copies, fontPx]);
  return /*#__PURE__*/React.createElement("div", {
    style: {
      flex: "0 0 auto"
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontWeight: 700,
      marginBottom: 4
    }
  }, "\u041F\u0440\u0435\u0434\u043F\u0440\u043E\u0441\u043C\u043E\u0442\u0440 \u043B\u0438\u0441\u0442\u0430"), /*#__PURE__*/React.createElement("div", {
    style: {
      border: "1px solid var(--line)",
      background: "#fff",
      boxShadow: "0 2px 10px rgba(19,44,83,.15)",
      width: 330,
      height: 440,
      overflow: "hidden"
    }
  }, /*#__PURE__*/React.createElement("div", {
    ref: ref,
    style: {
      zoom: 0.4,
      width: "210mm",
      padding: "8mm",
      color: "#000",
      fontFamily: "Arial,sans-serif"
    },
    dangerouslySetInnerHTML: {
      __html: cardsPageHtml(html, copies, fontPx, true)
    }
  })));
}
function safeFileName(s) {
  return String(s || "").replace(/[\\/:*?"<>|]/g, "").trim().slice(0, 80) || "uprazhnenie";
}

// Переводит HTML в список строк с сохранением переносов (в отличие от stripHtml,
// который схлопывает все пробелы/переносы в один пробел) — нужно для PPTX, где
// каждая строка стихотворения должна остаться отдельным абзацем.
function htmlToLines(html) {
  const d = document.createElement("div");
  d.innerHTML = html || "";
  Array.from(d.querySelectorAll("br")).forEach(br => br.parentNode.insertBefore(document.createTextNode("\n"), br));
  Array.from(d.querySelectorAll("p,div,li,h1,h2,h3")).forEach(el => el.appendChild(document.createTextNode("\n")));
  return (d.textContent || "").split("\n").map(l => l.replace(/[ \t]+/g, " ").trim());
}

// ---- Минимальный ZIP (без сжатия) — только для сборки .pptx на лету, без внешних библиотек ----
const CRC32_TABLE = (() => {
  const t = [];
  for (let n = 0; n < 256; n++) {
    let c = n;
    for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ c >>> 1 : c >>> 1;
    t[n] = c >>> 0;
  }
  return t;
})();
function crc32(bytes) {
  let crc = 0 ^ -1;
  for (let i = 0; i < bytes.length; i++) crc = crc >>> 8 ^ CRC32_TABLE[(crc ^ bytes[i]) & 0xFF];
  return (crc ^ -1) >>> 0;
}
function strToBytes(str) {
  return new TextEncoder().encode(str);
}
function makeZip(files) {
  const localParts = [];
  const centralParts = [];
  let offset = 0;
  files.forEach(({
    name,
    data
  }) => {
    const nameBytes = strToBytes(name);
    const crc = crc32(data);
    const size = data.length;
    const local = new Uint8Array(30 + nameBytes.length);
    const lv = new DataView(local.buffer);
    lv.setUint32(0, 0x04034b50, true);
    lv.setUint16(4, 20, true);
    lv.setUint16(6, 0, true);
    lv.setUint16(8, 0, true);
    lv.setUint16(10, 0, true);
    lv.setUint16(12, 0x21, true);
    lv.setUint32(14, crc, true);
    lv.setUint32(18, size, true);
    lv.setUint32(22, size, true);
    lv.setUint16(26, nameBytes.length, true);
    lv.setUint16(28, 0, true);
    local.set(nameBytes, 30);
    localParts.push(local, data);
    const central = new Uint8Array(46 + nameBytes.length);
    const cv = new DataView(central.buffer);
    cv.setUint32(0, 0x02014b50, true);
    cv.setUint16(4, 20, true);
    cv.setUint16(6, 20, true);
    cv.setUint16(8, 0, true);
    cv.setUint16(10, 0, true);
    cv.setUint16(12, 0, true);
    cv.setUint16(14, 0x21, true);
    cv.setUint32(16, crc, true);
    cv.setUint32(20, size, true);
    cv.setUint32(24, size, true);
    cv.setUint16(28, nameBytes.length, true);
    cv.setUint16(30, 0, true);
    cv.setUint16(32, 0, true);
    cv.setUint16(34, 0, true);
    cv.setUint16(36, 0, true);
    cv.setUint32(38, 0, true);
    cv.setUint32(42, offset, true);
    central.set(nameBytes, 46);
    centralParts.push(central);
    offset += local.length + data.length;
  });
  const centralStart = offset;
  const centralSize = centralParts.reduce((n, p) => n + p.length, 0);
  const end = new Uint8Array(22);
  const ev = new DataView(end.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, files.length, true);
  ev.setUint16(10, files.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, centralStart, true);
  ev.setUint16(20, 0, true);
  const allParts = [...localParts, ...centralParts, end];
  const out = new Uint8Array(allParts.reduce((n, p) => n + p.length, 0));
  let pos = 0;
  allParts.forEach(p => {
    out.set(p, pos);
    pos += p.length;
  });
  return out;
}
function showToast(msg) {
  const el = document.getElementById("toast");
  if (!el) return;
  el.textContent = msg;
  el.classList.add("show");
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => el.classList.remove("show"), 2200);
}

// ---- Цвета ----
function hexToHsl(hex) {
  let r = parseInt(hex.slice(1, 3), 16) / 255,
    g = parseInt(hex.slice(3, 5), 16) / 255,
    b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b),
    min = Math.min(r, g, b);
  let h,
    s,
    l = (max + min) / 2;
  if (max === min) {
    h = s = 0;
  } else {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    switch (max) {
      case r:
        h = (g - b) / d + (g < b ? 6 : 0);
        break;
      case g:
        h = (b - r) / d + 2;
        break;
      case b:
        h = (r - g) / d + 4;
        break;
    }
    h /= 6;
  }
  return [h * 360, s * 100, l * 100];
}
function hslToHex(h, s, l) {
  h /= 360;
  s /= 100;
  l /= 100;
  let r, g, b;
  if (s === 0) {
    r = g = b = l;
  } else {
    const hue2rgb = (p, q, t) => {
      if (t < 0) t += 1;
      if (t > 1) t -= 1;
      if (t < 1 / 6) return p + (q - p) * 6 * t;
      if (t < 1 / 2) return q;
      if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
      return p;
    };
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
    const p = 2 * l - q;
    r = hue2rgb(p, q, h + 1 / 3);
    g = hue2rgb(p, q, h);
    b = hue2rgb(p, q, h - 1 / 3);
  }
  const toHex = x => {
    const v = Math.round(x * 255).toString(16);
    return v.length === 1 ? "0" + v : v;
  };
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}
function contrastTextColor(hex) {
  if (!hex) return "#1E2530";
  const [,, l] = hexToHsl(hex);
  return l > 60 ? "#1E2530" : "#FFFFFF";
}
function generateShades(baseHex, count) {
  const [h] = hexToHsl(baseHex);
  const out = [];
  for (let i = 0; i < count; i++) {
    const t = count === 1 ? 0.5 : i / (count - 1);
    const l = 78 - t * 52;
    out.push(hslToHex(h, 62, Math.max(20, l)));
  }
  return out;
}
const SWATCHES = ["#132c53", "#2E5C9A", "#3B7DBF", "#2A9D8F", "#5C6BC0", "#7E57C2", "#8E4A86", "#C1554B", "#C2663F", "#C98A2B", "#B7791F", "#5E7A3A", "#8AA13E", "#4E6E8E", "#8C7B6B"];
const HIGHLIGHT_COLORS = ["#FFF3B0", "#C9F0D8", "#FBD8CE", "#D6E4FF", "#EBD6FF"];

// ---- Иконки (инлайн SVG, в стиле line-icons) ----
function Icon({
  name,
  size = 15
}) {
  const common = {
    width: size,
    height: size,
    viewBox: "0 0 24 24",
    fill: "none",
    stroke: "currentColor",
    strokeWidth: 2,
    strokeLinecap: "round",
    strokeLinejoin: "round"
  };
  const paths = {
    plus: /*#__PURE__*/React.createElement("path", {
      d: "M12 5v14M5 12h14"
    }),
    search: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("circle", {
      cx: "11",
      cy: "11",
      r: "7"
    }), /*#__PURE__*/React.createElement("path", {
      d: "m21 21-4.3-4.3"
    })),
    trash: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M4 7h16"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M9 7V5a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v2"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M6 7l1 13a1 1 0 0 0 1 1h8a1 1 0 0 0 1-1l1-13"
    })),
    chev: /*#__PURE__*/React.createElement("path", {
      d: "m9 6 6 6-6 6"
    }),
    left: /*#__PURE__*/React.createElement("path", {
      d: "m15 6-6 6 6 6"
    }),
    close: /*#__PURE__*/React.createElement("path", {
      d: "M18 6 6 18M6 6l12 12"
    }),
    tag: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M20.59 13.41 11 3.83A2 2 0 0 0 9.59 3.2H4a1 1 0 0 0-1 1v5.59a2 2 0 0 0 .59 1.41l9.58 9.59a2 2 0 0 0 2.83 0l4.59-4.59a2 2 0 0 0 0-2.83Z"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "7.5",
      cy: "7.5",
      r: "1.2"
    })),
    image: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("rect", {
      x: "3",
      y: "3",
      width: "18",
      height: "18",
      rx: "2"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "8.5",
      cy: "8.5",
      r: "1.5"
    }), /*#__PURE__*/React.createElement("path", {
      d: "m21 15-5-5L5 21"
    })),
    music: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M9 18V5l12-2v13"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "6",
      cy: "18",
      r: "3"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "18",
      cy: "16",
      r: "3"
    })),
    video: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("rect", {
      x: "2",
      y: "5",
      width: "14",
      height: "14",
      rx: "2"
    }), /*#__PURE__*/React.createElement("path", {
      d: "m22 8-6 4 6 4Z"
    })),
    note: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M15 3H6a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h12a2 2 0 0 0 2-2V9Z"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M15 3v6h6"
    })),
    save: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M19 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11l5 5v11a2 2 0 0 1-2 2Z"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M17 21v-8H7v8M7 3v5h8"
    })),
    folder: /*#__PURE__*/React.createElement("path", {
      d: "M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2Z"
    }),
    eye: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7-10-7-10-7Z"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "12",
      cy: "12",
      r: "3"
    })),
    eyeOff: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M17.94 17.94A10.94 10.94 0 0 1 12 19c-6.5 0-10-7-10-7a18.6 18.6 0 0 1 4.22-5.06M9.9 4.24A9.12 9.12 0 0 1 12 4c6.5 0 10 7 10 7a18.5 18.5 0 0 1-2.16 3.19m-6.72-1.07a3 3 0 1 1-4.24-4.24"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M1 1l22 22"
    })),
    link: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M9 17H7A5 5 0 0 1 7 7h2"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M15 7h2a5 5 0 1 1 0 10h-2"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M8 12h8"
    })),
    user: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("circle", {
      cx: "12",
      cy: "8",
      r: "4"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M4 21c0-4 4-7 8-7s8 3 8 7"
    })),
    book: /*#__PURE__*/React.createElement("path", {
      d: "M4 4.5A2.5 2.5 0 0 1 6.5 2H20v18H6.5A2.5 2.5 0 0 0 4 22.5Z M4 4.5v18"
    }),
    grip: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("circle", {
      cx: "9",
      cy: "6",
      r: "1.4",
      fill: "currentColor",
      stroke: "none"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "9",
      cy: "12",
      r: "1.4",
      fill: "currentColor",
      stroke: "none"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "9",
      cy: "18",
      r: "1.4",
      fill: "currentColor",
      stroke: "none"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "15",
      cy: "6",
      r: "1.4",
      fill: "currentColor",
      stroke: "none"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "15",
      cy: "12",
      r: "1.4",
      fill: "currentColor",
      stroke: "none"
    }), /*#__PURE__*/React.createElement("circle", {
      cx: "15",
      cy: "18",
      r: "1.4",
      fill: "currentColor",
      stroke: "none"
    })),
    check: /*#__PURE__*/React.createElement("path", {
      d: "M20 6 9 17l-5-5"
    }),
    edit: /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("path", {
      d: "M12 20h9"
    }), /*#__PURE__*/React.createElement("path", {
      d: "M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"
    }))
  };
  return /*#__PURE__*/React.createElement("svg", common, paths[name] || null);
}

// ---- UI-атомы ----
function Btn({
  kind = "default",
  sm,
  disabled,
  onClick,
  children,
  title,
  type = "button"
}) {
  const cls = ["btn", kind !== "default" ? kind : "", sm ? "sm" : ""].filter(Boolean).join(" ");
  return /*#__PURE__*/React.createElement("button", {
    type: type,
    className: cls,
    disabled: disabled,
    onClick: onClick,
    title: title
  }, children);
}
function Badge({
  tag,
  onClick,
  active,
  onRemove,
  small
}) {
  if (!tag) return null;
  const style = {
    background: tag.color || "#888",
    color: contrastTextColor(tag.color || "#888")
  };
  return /*#__PURE__*/React.createElement("span", {
    className: "chip",
    "aria-pressed": active === true ? "true" : undefined,
    style: onClick ? {
      cursor: "pointer"
    } : undefined,
    onClick: onClick
  }, /*#__PURE__*/React.createElement("span", {
    className: "cdot",
    style: style
  }), tag.name, tag.locked && /*#__PURE__*/React.createElement("span", {
    className: "badge-lock",
    title: "\u041F\u043E\u0441\u0442\u043E\u044F\u043D\u043D\u044B\u0439 \u0442\u0435\u0433"
  }, "\uD83D\uDD12"), onRemove && /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: e => {
      e.stopPropagation();
      onRemove();
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  })));
}
function Field({
  label,
  children
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, label && /*#__PURE__*/React.createElement("label", null, label), children);
}
function ColorPicker({
  value,
  onChange,
  colors = SWATCHES
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "pal"
  }, colors.map(c => /*#__PURE__*/React.createElement("div", {
    key: c,
    className: "sw",
    style: {
      background: c
    },
    "aria-pressed": value === c ? "true" : undefined,
    onClick: () => onChange(c),
    title: c
  })));
}
function TagSearchBar({
  value,
  onChange,
  placeholder
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "searchbar"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "search"
  }), /*#__PURE__*/React.createElement("input", {
    value: value,
    onChange: e => onChange(e.target.value),
    placeholder: placeholder || "Поиск по тегам и названию…"
  }), value && /*#__PURE__*/React.createElement("span", {
    className: "x",
    style: {
      cursor: "pointer"
    },
    onClick: () => onChange("")
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 13
  })));
}
function ViewModeToggle({
  value,
  onChange,
  options
}) {
  // options: [{id,label}]
  return /*#__PURE__*/React.createElement("div", {
    className: "viewtoggle"
  }, options.map(o => /*#__PURE__*/React.createElement("button", {
    key: o.id,
    "aria-pressed": value === o.id ? "true" : undefined,
    onClick: () => onChange(o.id)
  }, o.label)));
}

// Electron не поддерживает нативные window.prompt()/window.confirm() как
// блокирующие диалоги (клики по кнопкам «создать/удалить» иначе не работают
// в собранном приложении) — используем собственный модальный диалог.
let __dialogSetter = null;
function DialogHost() {
  const [state, setState] = useState(null);
  useEffect(() => {
    __dialogSetter = setState;
    return () => {
      __dialogSetter = null;
    };
  }, []);
  if (!state) return null;
  const finish = result => {
    setState(null);
    state.resolve(result);
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    style: {
      zIndex: 3000
    },
    onClick: () => finish(state.kind === "prompt" ? null : false)
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 440
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, state.message)), state.kind === "prompt" && /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("input", {
    autoFocus: true,
    value: state.value,
    onChange: e => setState(s => ({
      ...s,
      value: e.target.value
    })),
    onKeyDown: e => {
      if (e.key === "Enter") finish(state.value);
      if (e.key === "Escape") finish(null);
    },
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "9px 11px",
      fontSize: 15
    }
  })), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => finish(state.kind === "prompt" ? null : false)
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: () => finish(state.kind === "prompt" ? state.value : true)
  }, state.kind === "prompt" ? "ОК" : state.okLabel || "Да, удалить"))));
}
function Confirm(msg, okLabel) {
  return new Promise(resolve => {
    if (!__dialogSetter) {
      resolve(true);
      return;
    }
    __dialogSetter({
      kind: "confirm",
      message: msg,
      resolve,
      okLabel
    });
  });
}
function PromptStr(msg, def) {
  return new Promise(resolve => {
    if (!__dialogSetter) {
      resolve(null);
      return;
    }
    __dialogSetter({
      kind: "prompt",
      message: msg,
      value: def || "",
      resolve
    });
  });
}
;
// ============================================================
// dataModel.js — форма данных, миграция, общий tag-движок
// ============================================================

const LOCKED_GENRE_SEED = [{
  id: "lyric",
  name: "Лирика",
  color: "#3B7DBF",
  locked: true
}, {
  id: "prose",
  name: "Эпос",
  color: "#2A9D8F",
  locked: true
}, {
  id: "drama",
  name: "Драма",
  color: "#8E4A86",
  locked: true
}];
const HISTORY_THEORY_KIND = [{
  id: "history",
  name: "История",
  color: "#C98A2B",
  locked: true
}, {
  id: "theory",
  name: "Теория",
  color: "#132c53",
  locked: true
}];
// Постоянный вид упражнения «Прогулка по тропам» — узнаём его по стабильному id,
// а не по имени (имя можно переименовать, id остаётся тем же).
const TROPE_WALK_TYPE_ID = "trope-walk";
const TROPE_VIEW_MODES = [{
  id: "flat",
  label: "Общий список"
}, {
  id: "authors",
  label: "По авторам"
}, {
  id: "classes",
  label: "По классам",
  groupBy: "classTag"
}, {
  id: "tropes",
  label: "По тропам",
  groupBy: "tropeTag"
}];
function emptyCourse() {
  return {
    lectureHtml: "",
    attachments: [],
    questions: [],
    notes: []
  };
}
const CLASSES = ["5", "6", "7", "8", "9", "10", "11"];
function defaultState() {
  return {
    tags: {
      genre: LOCKED_GENRE_SEED.map(t => ({
        ...t
      })),
      literaryGenre: [],
      questionType: [],
      classTag: [],
      tropeTag: [],
      historyTheoryKind: HISTORY_THEORY_KIND.map(t => ({
        ...t
      })),
      exerciseType: [{
        id: TROPE_WALK_TYPE_ID,
        name: "Прогулка по тропам",
        color: "#B7791F",
        categoryId: null,
        order: -1,
        slots: [],
        locked: true
      }],
      lessonType: [],
      literaryDirection: []
    },
    personalities: [],
    works: [],
    theoryHistory: [],
    categories: [],
    examples: [],
    lessons: [],
    lessonSystems: [],
    courses: [],
    reserves: [],
    materialsForArticles: [],
    classLetters: [],
    // литеры классов {id, grade, name} — например, 5А/5Б внутри 5 класса
    courseSubgroups: [],
    // подгруппы курсов {id, courseId, name}
    classSettings: [],
    // настройки литеры/курса {id, kind, value, classLetterId, gradingSystem, periodMode, requiredMarks}
    seatingLayouts: [],
    // рассадка по партам {id, kind, value, classLetterId, files, visionLimits, seatFurther}
    calendar: {
      vacations: [],
      students: [],
      marks: [],
      journalDoneDates: []
    },
    schedule: {
      mode: "weekly",
      weekAAnchor: null,
      slots: [],
      undo: null
    },
    // «Моё расписание»
    favorites: {
      classes: [],
      courses: []
    } // избранные классы/курсы: показываются первыми, на них открываются «Уроки» и «Планирование»
  };
}
function migrateEntity(e, extra) {
  return {
    id: e.id || uid(),
    title: e.title ?? e.name ?? "",
    authorId: e.authorId ?? null,
    kind: e.kind ?? null,
    color: e.color ?? null,
    tagIds: Array.isArray(e.tagIds) ? e.tagIds : [],
    // Для каких классов/курсов материал доступен при составлении уроков (пусто — всем) и где отключён вручную.
    forClasses: Array.isArray(e.forClasses) ? e.forClasses.map(String) : [],
    forCourseIds: Array.isArray(e.forCourseIds) ? e.forCourseIds : [],
    hiddenFor: Array.isArray(e.hiddenFor) ? e.hiddenFor.filter(x => typeof x === "string") : [],
    alwaysFor: Array.isArray(e.alwaysFor) ? e.alwaysFor.filter(x => typeof x === "string") : [],
    course: {
      lectureHtml: e.course?.lectureHtml || "",
      attachments: Array.isArray(e.course?.attachments) ? e.course.attachments : [],
      questions: Array.isArray(e.course?.questions) ? e.course.questions.map(q => ({
        id: q.id || uid(),
        text: q.text || "",
        order: q.order ?? 0,
        questionTypeTagId: q.questionTypeTagId ?? null
      })) : [],
      notes: Array.isArray(e.course?.notes) ? e.course.notes : []
    },
    createdAt: e.createdAt || nowISO(),
    updatedAt: e.updatedAt || e.createdAt || nowISO(),
    ...extra
  };
}
function migrateExample(e) {
  return {
    id: e.id || uid(),
    typeTagId: e.typeTagId ?? null,
    questionTypeTagId: e.questionTypeTagId ?? null,
    text: e.text || "",
    note: e.note || "",
    tagIds: Array.isArray(e.tagIds) ? e.tagIds : [],
    attachments: Array.isArray(e.attachments) ? e.attachments : [],
    slotId: e.slotId ?? null,
    assignedStudentIds: Array.isArray(e.assignedStudentIds) ? e.assignedStudentIds : [],
    // Только для «Прогулки по тропам»: автор (ссылка на карточку в «Персоналиях» —
    // при сохранении заводится автоматически, если такой ещё нет) и название произведения.
    // Искомые тропы и класс — обычные теги в tagIds (namespaces tropeTag / classTag).
    authorId: e.authorId ?? null,
    workTitle: e.workTitle || "",
    // Области (класс/литера/курс), где стихотворение отключено вручную — ключи вида "c:6", "l:6:<id литеры>", "k:<id курса>".
    hiddenFor: Array.isArray(e.hiddenFor) ? e.hiddenFor.filter(x => typeof x === "string") : [],
    alwaysFor: Array.isArray(e.alwaysFor) ? e.alwaysFor.filter(x => typeof x === "string") : [],
    forClasses: Array.isArray(e.forClasses) ? e.forClasses.map(String) : [],
    forCourseIds: Array.isArray(e.forCourseIds) ? e.forCourseIds : [],
    // Вопросы к самому упражнению (как «Ход урока» у персоналий и теории).
    questions: Array.isArray(e.questions) ? e.questions.map((q, i) => ({
      id: q.id || uid(),
      text: q.text || "",
      order: q.order ?? i,
      questionTypeTagId: q.questionTypeTagId ?? null
    })) : [],
    createdAt: e.createdAt || nowISO(),
    updatedAt: e.updatedAt || e.createdAt || nowISO()
  };
}
function migrateLessonItem(it) {
  // На упражнение/вопрос можно вызвать сразу нескольких учеников — у каждого своя запись
  // в students[] (имя, отметка, markId). Старый формат (одиночные student/grade/…
  // прямо на item) конвертируется в этот массив один раз при загрузке файла.
  const students = Array.isArray(it.students) ? it.students : it.student ? [{
    id: uid(),
    name: it.student,
    grade: it.grade || "",
    gradeScale: it.gradeScale || "5",
    gradeExcused: it.gradeExcused == null ? null : it.gradeExcused,
    markId: it.markId || null
  }] : [];
  return {
    id: it.id || uid(),
    type: it.type,
    // "exercise" | "theory" | "personality" | "work" | "question"
    refId: it.refId ?? null,
    // для exercise/theory/personality/work — id сущности
    entityType: it.entityType ?? null,
    // для question — "theory"|"personality"|"work"
    entityId: it.entityId ?? null,
    // для question — id сущности-источника
    questionId: it.questionId ?? null,
    // для question — id вопроса в course.questions сущности
    note: it.note || "",
    // Вопрос «только в уроке» (не добавлен в карточку): { text, questionTypeTagId }
    local: it.local && typeof it.local === "object" ? {
      text: it.local.text || "",
      questionTypeTagId: it.local.questionTypeTagId ?? null
    } : null,
    students
  };
}
function migrateLesson(l) {
  return {
    id: l.id || uid(),
    class: l.class ?? null,
    courseId: l.courseId ?? null,
    classLetterId: l.classLetterId ?? null,
    copyGroupId: l.copyGroupId ?? null,
    systemId: l.systemId ?? null,
    topic: l.topic || "",
    lessonTypeTagIds: Array.isArray(l.lessonTypeTagIds) ? l.lessonTypeTagIds : [],
    items: Array.isArray(l.items) ? l.items.map(migrateLessonItem) : [],
    homework: l.homework || "",
    homeworkAttachments: Array.isArray(l.homeworkAttachments) ? l.homeworkAttachments : [],
    completed: !!l.completed,
    archived: !!l.archived,
    plannedDate: l.plannedDate || "",
    actualDate: l.actualDate || "",
    excludeFromSchedule: !!l.excludeFromSchedule,
    // номер урока в расписании действителен только пока урок стоит на той дате, для которой он назначен
    lessonNo: l.lessonNo || null,
    lessonNoDate: l.lessonNoDate || "",
    needsWork: !!l.needsWork,
    // Урок-дубликат: id исходного урока (в списках пишется «(дубликат)»)
    duplicateOf: l.duplicateOf ?? null,
    // Примечания-«стикеры» к карточкам в уроке: { "personality:<id>": "текст" }; к упражнениям и вопросам — item.note
    groupNotes: l.groupNotes && typeof l.groupNotes === "object" && !Array.isArray(l.groupNotes) ? l.groupNotes : {},
    order: l.order ?? 0,
    createdAt: l.createdAt || nowISO(),
    updatedAt: l.updatedAt || l.createdAt || nowISO()
  };
}
function migrateLessonSystem(s) {
  return {
    id: s.id || uid(),
    title: s.title || "",
    color: s.color || SWATCHES[0],
    colorLinkKind: s.colorLinkKind ?? null,
    colorLinkId: s.colorLinkId ?? null,
    class: s.class ?? null,
    courseId: s.courseId ?? null,
    // Литера системы (id из classLetters). null — «общая» система, видна во всех литерах класса.
    // Так для 6Б и 6Г можно вести свои, независимые системы уроков.
    classLetterId: s.classLetterId ?? null,
    copyGroupId: s.copyGroupId ?? null,
    archived: !!s.archived,
    // выставляется автоматически (normalizeSystemArchive): все уроки системы пройдены
    lessonIds: Array.isArray(s.lessonIds) ? s.lessonIds : [],
    collapsed: !!s.collapsed,
    order: s.order ?? 0,
    createdAt: s.createdAt || nowISO()
  };
}
function migrateState(raw) {
  const base = defaultState();
  const d = {
    ...base,
    ...(raw || {})
  };
  d.tags = {
    ...base.tags,
    ...(raw && raw.tags ? raw.tags : {})
  };
  // Гарантируем наличие постоянных тегов рода и раздела история/теория. Имя/locked синхронизируем
  // с сидом при каждой загрузке — так переименование постоянного тега (например, «Проза» → «Эпос»)
  // применяется и к уже сохранённым файлам, а не только к новым проектам.
  LOCKED_GENRE_SEED.forEach(seed => {
    const idx = d.tags.genre.findIndex(t => t.id === seed.id);
    if (idx === -1) d.tags.genre.unshift({
      ...seed
    });else d.tags.genre[idx] = {
      ...d.tags.genre[idx],
      name: seed.name,
      locked: true
    };
  });
  d.tags.historyTheoryKind = HISTORY_THEORY_KIND.map(seed => {
    const found = (raw?.tags?.historyTheoryKind || []).find(t => t.id === seed.id);
    return found ? {
      ...seed,
      ...found,
      locked: true
    } : {
      ...seed
    };
  });
  if (!Array.isArray(d.tags.literaryGenre)) d.tags.literaryGenre = [];
  if (!Array.isArray(d.tags.questionType)) d.tags.questionType = [];
  if (!Array.isArray(d.tags.classTag)) d.tags.classTag = [];
  if (!Array.isArray(d.tags.tropeTag)) d.tags.tropeTag = [];
  if (!Array.isArray(d.tags.exerciseType)) d.tags.exerciseType = [];
  d.tags.exerciseType = d.tags.exerciseType.map(t => ({
    id: t.id || uid(),
    name: t.name || "",
    color: t.color || SWATCHES[0],
    categoryId: t.categoryId ?? null,
    order: t.order ?? 0,
    slots: Array.isArray(t.slots) ? t.slots : [],
    locked: !!t.locked,
    // Области (ключи "c:6", "l:6:<id>", "k:<id>"), где ВСЕ упражнения этого вида предлагаются всегда. Для «Поэтической
    // пятиминутки» и «Поделись книгой» по умолчанию — 5–7 классы (заводится один раз, потом редактируется как угодно).
    alwaysFor: Array.isArray(t.alwaysFor) ? t.alwaysFor.filter(x => typeof x === "string") : /^\s*(поэтическая пятиминутка|поделись книгой)\s*$/i.test(t.name || "") ? ["c:5", "c:6", "c:7"] : []
  }));
  // «Прогулка по тропам» — встроенный вид упражнения (со своей карточкой, группировкой
  // по автору/классу/тропу и экспортом) — заводим его один раз, если ещё не создан.
  if (!d.tags.exerciseType.find(t => t.id === TROPE_WALK_TYPE_ID)) {
    d.tags.exerciseType.unshift({
      id: TROPE_WALK_TYPE_ID,
      name: "Прогулка по тропам",
      color: "#B7791F",
      categoryId: null,
      order: -1,
      slots: [],
      locked: true
    });
  } else {
    const idx = d.tags.exerciseType.findIndex(t => t.id === TROPE_WALK_TYPE_ID);
    d.tags.exerciseType[idx] = {
      ...d.tags.exerciseType[idx],
      locked: true
    };
  }
  if (!Array.isArray(d.tags.lessonType)) d.tags.lessonType = [];
  if (!Array.isArray(d.tags.literaryDirection)) d.tags.literaryDirection = [];
  d.personalities = (d.personalities || []).map(p => migrateEntity(p, {}));
  d.works = (d.works || []).map(w => migrateEntity(w, {}));
  d.theoryHistory = (d.theoryHistory || []).map(t => migrateEntity(t, {
    kind: t.kind === "theory" ? "theory" : "history"
  }));
  d.categories = (d.categories || []).map(c => ({
    id: c.id || uid(),
    name: c.name || "",
    color: c.color || SWATCHES[0],
    order: c.order ?? 0
  }));
  d.examples = (d.examples || []).map(migrateExample);
  d.lessons = (d.lessons || []).map(migrateLesson);
  d.lessonSystems = (d.lessonSystems || []).map(migrateLessonSystem);
  d.courses = (d.courses || []).map(c => ({
    id: c.id || uid(),
    name: c.name || "",
    color: c.color || SWATCHES[Math.floor(Math.random() * SWATCHES.length)]
  }));
  d.reserves = Array.isArray(d.reserves) ? d.reserves : [];
  d.materialsForArticles = d.materialsForArticles || [];
  d.classLetters = (d.classLetters || []).map(cl => ({
    id: cl.id || uid(),
    grade: cl.grade,
    name: cl.name || ""
  }));
  d.courseSubgroups = (d.courseSubgroups || []).map(g => ({
    id: g.id || uid(),
    courseId: g.courseId,
    name: g.name || ""
  }));
  d.classSettings = (d.classSettings || []).map(s => ({
    id: s.id || uid(),
    kind: s.kind,
    value: s.value,
    classLetterId: s.classLetterId || null,
    gradingSystem: s.gradingSystem || "5",
    periodMode: s.periodMode || "quarter",
    requiredMarks: s.requiredMarks || null
  }));
  d.seatingLayouts = (d.seatingLayouts || []).map(l => ({
    id: l.id || uid(),
    kind: l.kind,
    value: l.value,
    classLetterId: l.classLetterId || null,
    files: Array.isArray(l.files) ? l.files : [[], [], []],
    visionLimits: Array.isArray(l.visionLimits) ? l.visionLimits : [],
    seatFurther: Array.isArray(l.seatFurther) ? l.seatFurther : [],
    // группы «нежелательно рядом»; старые пары (avoidPairs), если были, становятся группами из двух
    avoidGroups: [...(Array.isArray(l.avoidGroups) ? l.avoidGroups : []).map(g => ({
      id: g.id || uid(),
      members: Array.from(new Set(g.members || []))
    })), ...(Array.isArray(l.avoidPairs) ? l.avoidPairs : []).map(p => ({
      id: uid(),
      members: [p.a, p.b].filter(Boolean)
    }))].filter(g => g.members.length >= 2)
  }));
  {
    // Раньше системы не имели литеры и показывались во всех литерах класса. Для систем, у которых
    // поле классовой литеры ещё не задавалось, определяем литеру по их урокам (если все уроки
    // системы — одной литеры, система становится системой этой литеры). Смешанные и пустые системы
    // остаются «общими». Поле сохраняется в файле, поэтому дальше можно менять вручную.
    (raw && Array.isArray(raw.lessonSystems) ? raw.lessonSystems : []).forEach(rawSys => {
      if (rawSys && Object.prototype.hasOwnProperty.call(rawSys, "classLetterId")) return;
      const sys = d.lessonSystems.find(x => x.id === rawSys.id);
      if (!sys) return;
      const letters = new Set(sys.lessonIds.map(id => d.lessons.find(l => l.id === id)).filter(Boolean).map(l => l.classLetterId || null));
      if (letters.size === 1) sys.classLetterId = Array.from(letters)[0];
    });
  }
  {
    // Самолечение: урок с systemId, которого на самом деле нет в lessonIds ни одной системы,
    // становится "невидимым" — не отдельный (systemId задан), но и не внутри системы. Раньше
    // так ломались копии в другие литеры, сделанные из урока внутри системы. Возвращаем таким
    // урокам статус отдельных, чтобы они снова появились в списке "Уроки".
    d.lessons.forEach(l => {
      if (l.systemId && !d.lessonSystems.some(s => s.id === l.systemId && s.lessonIds.includes(l.id))) l.systemId = null;
    });
  }
  {
    // Расстановка блоков в списке «Уроки» (отдельные уроки и системы вперемешку) — выбор учителя, поэтому
    // при загрузке её НЕЛЬЗЯ перестраивать (раньше здесь системы принудительно поднимались наверх, а отдельные
    // уроки выстраивались по датам — и вся ручная расстановка пропадала при каждом запуске). Исправляем только
    // совпадающие значения order внутри одного вида списка (старые файлы, где уроки и системы считались раздельно).
    const views = [];
    d.lessons.forEach(l => {
      if (!views.some(x => sameLessonGroup(x, l))) views.push({
        class: l.class,
        courseId: l.courseId,
        classLetterId: l.classLetterId || null
      });
    });
    views.forEach(sample => {
      const entries = topLevelEntries(d, sample);
      const orders = entries.map(e => e.order);
      if (new Set(orders).size === orders.length) return;
      entries.sort((a, b) => a.order - b.order || (a.type === b.type ? 0 : a.type === "system" ? -1 : 1));
      entries.forEach((e, idx) => {
        const obj = e.type === "lesson" ? d.lessons.find(x => x.id === e.id) : d.lessonSystems.find(x => x.id === e.id);
        if (obj) obj.order = idx;
      });
    });
  }
  [d.personalities, d.works, d.theoryHistory, d.examples].forEach(arr => arr.forEach(e => {
    if (Array.isArray(e.forClasses) && e.forClasses.length) {
      e.forClasses.forEach(c => {
        const id = classTagIdForGrade(d, String(c));
        if (id && !e.tagIds.includes(id)) e.tagIds.push(id);
      });
      e.forClasses = [];
    }
  }));
  d.calendar = d.calendar || {
    vacations: [],
    students: [],
    marks: []
  };
  d.calendar.vacations = (d.calendar.vacations || []).map(v => ({
    id: v.id || uid(),
    name: v.name || "Каникулы",
    start: v.start || "",
    end: v.end || v.start || ""
  }));
  d.calendar.students = (d.calendar.students || []).map(s => ({
    id: s.id || uid(),
    name: s.name || "",
    class: s.courseId ? null : s.class || CLASSES[0],
    courseId: s.courseId || null,
    classLetterId: s.classLetterId || null,
    courseSubgroupId: s.courseSubgroupId || null,
    // закрытые («Начать новую четверть») четверти и «не вызывать» (домашнее обучение)
    closedQ: Array.isArray(s.closedQ) ? s.closedQ : [],
    noCall: !!s.noCall
  }));
  d.calendar.marks = (d.calendar.marks || []).map(m => ({
    id: m.id || uid(),
    studentId: m.studentId,
    value: m.value || 0,
    date: m.date || "",
    reason: m.reason || "",
    scale: m.scale || "5",
    excused: m.excused == null ? null : m.excused,
    lessonId: m.lessonId || null,
    exerciseId: m.exerciseId || null,
    questionId: m.questionId || null,
    posted: !!m.posted // «выставлена» в школьный электронный журнал
  }));
  {
    const sc = raw && raw.schedule || {};
    d.schedule = {
      mode: sc.mode === "biweekly" ? "biweekly" : "weekly",
      weekAAnchor: sc.weekAAnchor || null,
      slots: (sc.slots || []).filter(x => x && x.value && x.dow && x.num).map(x => ({
        id: x.id || uid(),
        week: x.week === "B" ? "B" : "A",
        dow: +x.dow,
        num: +x.num,
        kind: x.kind === "course" ? "course" : "class",
        value: String(x.value),
        classLetterId: x.classLetterId || null
      })),
      undo: sc.undo && Array.isArray(sc.undo.items) ? sc.undo : null
    };
  }
  d.favorites = {
    classes: [],
    courses: [],
    ...(raw?.favorites || {})
  };
  d.favorites.classes = Array.isArray(d.favorites.classes) ? d.favorites.classes : [];
  d.favorites.courses = Array.isArray(d.favorites.courses) ? d.favorites.courses : [];
  d.calendar.periodMarks = Array.isArray(d.calendar.periodMarks) ? d.calendar.periodMarks : []; // четвертные/годовые/экзаменационные
  d.calendar.journalFileSnapshots = Array.isArray(d.calendar.journalFileSnapshots) ? d.calendar.journalFileSnapshots : [];
  // Отмеченные учителем как «отметки перенесены в журнал» столбцы-даты в «Журнале» —
  // просто список ключей вида "класс/курс:значение:фильтр литеры:дата", красится светло-серым.
  d.calendar.journalDoneDates = Array.isArray(d.calendar.journalDoneDates) ? d.calendar.journalDoneDates : [];
  // Старый флажок «отметки перенесены в журнал» у даты → флаг «выставлена» у самих отметок (один раз, затем список пуст)
  if (d.calendar.journalDoneDates.length) {
    d.calendar.journalDoneDates.forEach(key => {
      const p = String(key).split(":");
      if (p.length < 4) return;
      const [kd, val, lf, dt] = p;
      const ids = new Set(d.calendar.students.filter(st => (kd === "course" ? st.courseId === val : st.class === val && !st.courseId) && (kd !== "class" || lf === "all" || (lf === "none" ? !st.classLetterId : st.classLetterId === lf))).map(st => st.id));
      d.calendar.marks.forEach(m => {
        if (ids.has(m.studentId) && m.date === dt) m.posted = true;
      });
    });
    d.calendar.journalDoneDates = [];
  }
  return d;
}

// Пройденный урок уходит в архив (из списка «Уроки»), но остаётся в календаре и не теряется для журнала.
function setLessonCompleted(l, on) {
  l.completed = on;
  l.archived = on;
  if (on && !l.actualDate) l.actualDate = l.plannedDate || todayIso();
}
// «В игре» — урок, который ещё идёт по плану ИЛИ уже пройден (пройденные тоже архивные, но нужны календарю, «Сегодня» и расписанию).
function isLessonInPlay(l) {
  return !l.archived || !!l.completed;
}
// Система в архиве, когда в ней есть уроки и все они пройдены; стоит вернуть любой урок — система снова в списке.
function normalizeSystemArchive(d) {
  const byId = new Map(d.lessons.map(l => [l.id, l]));
  (d.lessonSystems || []).forEach(sy => {
    const ls = (sy.lessonIds || []).map(id => byId.get(id)).filter(Boolean);
    sy.archived = ls.length > 0 && ls.every(l => l.completed);
  });
}
function removeStudentFromSeating(d, id) {
  (d.seatingLayouts || []).forEach(l => {
    l.files.forEach(file => file.forEach(desk => {
      desk.seats = desk.seats.map(sid => sid === id ? null : sid);
    }));
    l.visionLimits = (l.visionLimits || []).filter(v => v.studentId !== id);
    l.seatFurther = (l.seatFurther || []).filter(sid => sid !== id);
    l.avoidGroups = (l.avoidGroups || []).map(g => ({
      ...g,
      members: g.members.filter(m => m !== id)
    })).filter(g => g.members.length >= 2);
  });
}
function removeStudentCascade(d, studentId) {
  d.calendar.students = d.calendar.students.filter(s => s.id !== studentId);
  d.calendar.marks = d.calendar.marks.filter(m => m.studentId !== studentId);
  d.calendar.periodMarks = (d.calendar.periodMarks || []).filter(p => p.studentId !== studentId);
  removeStudentFromSeating(d, studentId);
  (d.examples || []).forEach(e => {
    e.assignedStudentIds = (e.assignedStudentIds || []).filter(id => id !== studentId);
  });
}
function nsTagIds(tagIds, ns, data) {
  return (tagIds || []).filter(id => (data.tags[ns] || []).some(t => t.id === id));
}
function withNsTagIds(tagIds, ns, data, newIds) {
  return [...(tagIds || []).filter(id => !(data.tags[ns] || []).some(t => t.id === id)), ...newIds];
}
function findTagAnywhere(data, id) {
  for (const ns of Object.keys(data.tags)) {
    const t = (data.tags[ns] || []).find(x => x.id === id);
    if (t) return t;
  }
  return null;
}
function ensureTag(d, ns, name) {
  name = (name || "").trim();
  if (!name) return null;
  let tag = d.tags[ns].find(t => t.name.toLowerCase() === name.toLowerCase());
  if (tag) return tag.id;
  const id = uid();
  const color = SWATCHES[Math.floor(Math.random() * SWATCHES.length)];
  d.tags[ns].push({
    id,
    name,
    color
  });
  return id;
}

// Находит персоналию по имени (без учёта регистра) или заводит новую карточку в «Персоналиях» —
// используется, когда автор указывается прямо в упражнении (например, в «Прогулке по тропам»).
function ensurePersonality(d, name) {
  name = (name || "").trim();
  if (!name) return null;
  const found = d.personalities.find(p => (p.title || "").trim().toLowerCase() === name.toLowerCase());
  if (found) return found.id;
  const p = migrateEntity({
    title: name
  }, {});
  d.personalities.push(p);
  return p.id;
}
function removeTagCascade(d, ns, id) {
  d.tags[ns] = d.tags[ns].filter(t => t.id !== id);
  const stripFrom = arr => arr.forEach(e => {
    if (e.tagIds) e.tagIds = e.tagIds.filter(x => x !== id);
  });
  stripFrom(d.personalities);
  stripFrom(d.works);
  stripFrom(d.theoryHistory);
  stripFrom(d.examples || []);
  if (ns === "questionType") {
    const stripQ = arr => arr.forEach(e => {
      (e.course?.questions || []).forEach(q => {
        if (q.questionTypeTagId === id) q.questionTypeTagId = null;
      });
    });
    stripQ(d.personalities);
    stripQ(d.works);
    stripQ(d.theoryHistory);
    (d.examples || []).forEach(e => {
      if (e.questionTypeTagId === id) e.questionTypeTagId = null;
    });
  }
  if (ns === "exerciseType") {
    (d.examples || []).forEach(e => {
      if (e.typeTagId === id) {
        e.typeTagId = null;
        e.slotId = null;
      }
    });
  }
  if (ns === "lessonType") {
    (d.lessons || []).forEach(l => {
      l.lessonTypeTagIds = (l.lessonTypeTagIds || []).filter(x => x !== id);
    });
  }
}

// Двусторонняя синхронизация цвета между карточкой персоналии/теории и блоком системы уроков.
function syncColorToLessonSystems(d, entityKind, entityId, color) {
  (d.lessonSystems || []).forEach(s => {
    if (s.colorLinkKind === entityKind && s.colorLinkId === entityId) s.color = color;
  });
}
function syncColorFromLessonSystem(d, system) {
  if (!system.colorLinkKind || !system.colorLinkId) return;
  const arr = d[system.colorLinkKind];
  if (!arr) return;
  const idx = arr.findIndex(e => e.id === system.colorLinkId);
  if (idx >= 0) arr[idx] = {
    ...arr[idx],
    color: system.color
  };
}

// ---- Вложения ----
async function pickAndSaveOne(projectPath, kind) {
  if (!window.litAPI) return null;
  const picked = await window.litAPI.pickFiles(kind, false);
  if (!picked || !picked.ok || !picked.paths || !picked.paths[0]) return null;
  const filePath = picked.paths[0];
  const name = filePath.split(/[\\/]/).pop();
  const saved = await window.litAPI.saveAttachment(projectPath, filePath, name);
  if (!saved || !saved.ok) {
    showToast(saved && saved.error || "Не удалось сохранить вложение");
    return null;
  }
  return {
    id: saved.id,
    relPath: saved.relPath,
    name: saved.name,
    size: saved.size,
    kind
  };
}
function fileUrl(relPath) {
  return window.litAPI ? window.litAPI.fileUrl(relPath) : "";
}
// Скачивание вложения (litfile:) через временную ссылку на blob — программа предложит выбрать, куда сохранить файл.
window.__litProjectPath = null; // путь открытого проекта (ставит App) — нужен, чтобы вернуть потерявшийся файл вложения
async function downloadAttachment(href, name, retried) {
  if (window.litTabletMode) {
    showToast("Файлы-вложения (презентации и т.п.) хранятся на ноутбуке и на планшете не открываются");
    return;
  }
  try {
    let r = null;
    try {
      r = await fetch(href);
    } catch (e) {
      r = null;
    }
    if (!r || !r.ok) {
      // Файла нет в папке «<проект>_files»: чаще всего её не перенесли вместе с файлом проекта на другой компьютер.
      if (!retried && window.litAPI && window.litAPI.restoreAttachment && window.__litProjectPath) {
        const rel = decodeURIComponent(href.replace(/^litfile:\/*/, ""));
        if (await Confirm("Файл «" + (name || rel) + "» не найден в папке вложений рядом с файлом проекта (папка «…_files»). Вероятно, она не перенесена на этот компьютер вместе с проектом. Указать, где лежит этот файл?", "Указать файл")) {
          const res = await window.litAPI.restoreAttachment(window.__litProjectPath, rel, name);
          if (res && res.ok) {
            window.dispatchEvent(new Event("lit-attachments-changed"));
            return downloadAttachment(href, name, true);
          }
          if (res && res.error) showToast("Не удалось вернуть файл: " + res.error);
        }
        return;
      }
      throw new Error(r ? "HTTP " + r.status : "файл не найден");
    }
    const blob = await r.blob();
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = name || "файл";
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 15000);
  } catch (e) {
    showToast("Не удалось скачать файл: " + (e && e.message ? e.message : e));
  }
}
document.addEventListener("click", ev => {
  const a = ev.target && ev.target.closest ? ev.target.closest('a[href^="litfile:"]') : null;
  if (!a) return;
  ev.preventDefault();
  ev.stopPropagation();
  const blk = a.closest(".lc-block");
  const nm = blk && blk.querySelector(".blk-name") && blk.querySelector(".blk-name").textContent || a.textContent.replace(/^⬇\s*/, "").trim();
  downloadAttachment(a.getAttribute("href"), nm);
}, true);

// ---- Выделение текста: обёртка Range в span, без потери контента при выделениях, задевающих несколько узлов ----
function wrapRangeInSpan(range, className, dataAttrs) {
  const span = document.createElement("span");
  span.className = className;
  Object.entries(dataAttrs || {}).forEach(([k, v]) => span.setAttribute(k, v));
  try {
    range.surroundContents(span);
  } catch (e) {
    const frag = range.extractContents();
    span.appendChild(frag);
    range.insertNode(span);
  }
  return span;
}
function unwrapSpanById(container, attrName, attrValue) {
  if (!container) return false;
  const el = container.querySelector(`[${attrName}="${CSS.escape(String(attrValue))}"]`);
  if (!el) return false;
  const parent = el.parentNode;
  while (el.firstChild) parent.insertBefore(el.firstChild, el);
  parent.removeChild(el);
  return true;
}
function htmlHasAttr(html, attrName, attrValue) {
  if (!html) return false;
  try {
    const dp = new DOMParser();
    const doc = dp.parseFromString(html, "text/html");
    return !!doc.querySelector(`[${attrName}="${CSS.escape(String(attrValue))}"]`);
  } catch (e) {
    return false;
  }
}

// Все фрагменты текста, привязанные к вопросу qId (может быть несколько — каждый со своим
// data-anchor-id, подсвечиваются вместе по общему data-qa-id).
function qaFragmentsOf(lectureHtml, qId) {
  if (!lectureHtml) return [];
  try {
    const doc = new DOMParser().parseFromString(lectureHtml, "text/html");
    return Array.from(doc.querySelectorAll(`[data-qa-id="${CSS.escape(qId)}"]`)).map(el => ({
      anchorId: el.getAttribute("data-anchor-id") || "",
      text: el.textContent || ""
    }));
  } catch (e) {
    return [];
  }
}
;
// ============================================================
// richEditor.js — редактор в стиле Notion: форматирование, блоки
// изображение/аудио/видео, заметки/примечания к выделению.
// ============================================================
const {
  forwardRef,
  useImperativeHandle
} = React;
function ToolBtn({
  onClick,
  active,
  title,
  children
}) {
  return /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "toolbtn" + (active ? " active" : ""),
    title: title,
    onMouseDown: e => e.preventDefault(),
    onClick: onClick
  }, children);
}
function getClosestTable() {
  const sel = window.getSelection();
  if (!sel || !sel.anchorNode) return null;
  const el = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode;
  return el ? el.closest("table") : null;
}
function getClosestCell() {
  const sel = window.getSelection();
  if (!sel || !sel.anchorNode) return null;
  const el = sel.anchorNode.nodeType === 3 ? sel.anchorNode.parentElement : sel.anchorNode;
  return el ? el.closest("td,th") : null;
}
const RichEditor = forwardRef(function RichEditor(props, ref) {
  const {
    html,
    onChange,
    projectPath,
    compact,
    notesHidden,
    onAddNoteClick,
    onBlockClick,
    onNoteAnchorClick,
    placeholder,
    update,
    sourceInfo
  } = props;
  const editorRef = useRef(null);
  const savedRangeRef = useRef(null);
  const mountedHtmlRef = useRef(null);
  useEffect(() => {
    if (editorRef.current && mountedHtmlRef.current === null) {
      editorRef.current.innerHTML = html || "";
      mountedHtmlRef.current = html || "";
    }
  }, []);
  const saveHtml = () => {
    if (editorRef.current) onChange(editorRef.current.innerHTML);
  };
  const focusEd = () => editorRef.current && editorRef.current.focus();
  const exec = (cmd, value) => {
    focusEd();
    document.execCommand(cmd, false, value);
    saveHtml();
  };
  const execHighlight = color => {
    focusEd();
    if (!document.execCommand("hiliteColor", false, color)) document.execCommand("backColor", false, color);
    saveHtml();
  };
  const onSelectionCapture = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount && !sel.isCollapsed) {
      const r = sel.getRangeAt(0);
      if (editorRef.current && editorRef.current.contains(r.commonAncestorContainer)) {
        savedRangeRef.current = r.cloneRange();
      }
    }
  };
  useImperativeHandle(ref, () => ({
    getHtml: () => editorRef.current ? editorRef.current.innerHTML : html || "",
    hasSelection: () => !!savedRangeRef.current,
    getContainer: () => editorRef.current,
    wrapSelection: (className, dataAttrs) => {
      if (!savedRangeRef.current) return null;
      let span = null;
      try {
        span = wrapRangeInSpan(savedRangeRef.current, className, dataAttrs);
      } catch (e) {
        span = null;
      }
      savedRangeRef.current = null;
      saveHtml();
      return span;
    }
  }));
  async function sendToMaterials() {
    if (!savedRangeRef.current) {
      showToast("Сначала выделите фрагмент текста");
      return;
    }
    const div = document.createElement("div");
    div.appendChild(savedRangeRef.current.cloneContents());
    const excerptHtml = div.innerHTML;
    if (!stripHtml(excerptHtml).trim()) {
      showToast("Пустое выделение");
      return;
    }
    const articleName = await PromptStr("Название статьи (можно уточнить позже):", "");
    const genreTagIds = sourceInfo?.genreTagIds || [];
    update(d => {
      d.materialsForArticles.push({
        id: uid(),
        excerptHtml,
        note: "",
        articleName: (articleName || "").trim(),
        tagIds: [...genreTagIds],
        sourceSection: sourceInfo?.section || "",
        sourceEntityId: sourceInfo?.entityId || null,
        sourceEntityTitle: sourceInfo?.entityTitle || "",
        authorId: sourceInfo?.authorId || null,
        workId: sourceInfo?.workId || null,
        createdAt: nowISO()
      });
      return d;
    });
    showToast("Добавлено в «Материалы для статьи»");
  }
  async function insertMedia(kind) {
    if (!projectPath) {
      showToast("Сначала сохраните файл проекта — вложения хранятся рядом с ним");
      return;
    }
    const att = await pickAndSaveOne(projectPath, kind);
    if (!att) return;
    const url = fileUrl(att.relPath);
    let inner = "";
    if (kind === "image") inner = `<img src="${url}" alt="${escapeHtml(att.name)}">`;else if (kind === "audio") inner = `<audio controls src="${url}"></audio>`;else if (kind === "video") inner = `<video controls src="${url}"></video>`;else inner = `<a href="${url}" target="_blank">${escapeHtml(att.name)}</a>`;
    const blockHtml = `<div class="lc-block" contenteditable="false" data-block-id="${att.id}"><span class="blk-del" data-del-block="1">✕</span>${inner}<div class="blk-name">${escapeHtml(att.name)}</div></div><p><br></p>`;
    focusEd();
    document.execCommand("insertHTML", false, blockHtml);
    saveHtml();
  }
  async function insertTable() {
    const rows = parseInt((await PromptStr("Сколько строк?", "3")) || "0", 10);
    const cols = parseInt((await PromptStr("Сколько столбцов?", "3")) || "0", 10);
    if (!rows || !cols) return;
    let t = '<table><tbody>';
    for (let r = 0; r < rows; r++) {
      t += "<tr>";
      for (let c = 0; c < cols; c++) t += "<td>&nbsp;</td>";
      t += "</tr>";
    }
    t += "</tbody></table><p><br></p>";
    focusEd();
    document.execCommand("insertHTML", false, t);
    saveHtml();
  }
  function addRow() {
    const table = getClosestTable();
    if (!table) return;
    const firstRow = table.querySelector("tr");
    const cols = firstRow ? firstRow.children.length : 1;
    const tr = document.createElement("tr");
    for (let c = 0; c < cols; c++) {
      const td = document.createElement("td");
      td.innerHTML = "&nbsp;";
      tr.appendChild(td);
    }
    table.querySelector("tbody, table").appendChild ? (table.querySelector("tbody") || table).appendChild(tr) : table.appendChild(tr);
    saveHtml();
  }
  function addCol() {
    const table = getClosestTable();
    if (!table) return;
    table.querySelectorAll("tr").forEach(tr => {
      const td = document.createElement("td");
      td.innerHTML = "&nbsp;";
      tr.appendChild(td);
    });
    saveHtml();
  }
  function delRow() {
    const cell = getClosestCell();
    const table = getClosestTable();
    if (!cell || !table) return;
    const tr = cell.closest("tr");
    if (tr) tr.remove();
    saveHtml();
  }
  function delCol() {
    const cell = getClosestCell();
    const table = getClosestTable();
    if (!cell || !table) return;
    const idx = Array.prototype.indexOf.call(cell.parentElement.children, cell);
    table.querySelectorAll("tr").forEach(tr => {
      if (tr.children[idx]) tr.children[idx].remove();
    });
    saveHtml();
  }
  function delTable() {
    const table = getClosestTable();
    if (!table) return;
    table.remove();
    saveHtml();
  }
  async function handleEditorClick(e) {
    const del = e.target.closest(".blk-del");
    if (del) {
      const block = del.closest(".lc-block");
      if (block && (await Confirm("Удалить этот блок?"))) {
        block.remove();
        saveHtml();
      }
      return;
    }
    const note = e.target.closest(".lc-note");
    if (note) {
      onNoteAnchorClick && onNoteAnchorClick(note.getAttribute("data-note-id"));
      return;
    }
  }
  return /*#__PURE__*/React.createElement("div", {
    className: "editorwrap" + (notesHidden ? " notes-hidden" : "")
  }, /*#__PURE__*/React.createElement("div", {
    className: "edtoolbar"
  }, /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0416\u0438\u0440\u043D\u044B\u0439",
    onClick: () => exec("bold")
  }, /*#__PURE__*/React.createElement("b", null, "\u0416")), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u041A\u0443\u0440\u0441\u0438\u0432",
    onClick: () => exec("italic")
  }, /*#__PURE__*/React.createElement("i", null, "\u041A")), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u041F\u043E\u0434\u0447\u0451\u0440\u043A\u043D\u0443\u0442\u044B\u0439",
    onClick: () => exec("underline")
  }, /*#__PURE__*/React.createElement("u", null, "\u0427")), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0417\u0430\u0447\u0451\u0440\u043A\u043D\u0443\u0442\u044B\u0439",
    onClick: () => exec("strikeThrough")
  }, /*#__PURE__*/React.createElement("s", null, "\u0417")), /*#__PURE__*/React.createElement("div", {
    className: "sep"
  }), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    title: "\u0421\u0442\u0438\u043B\u044C",
    defaultValue: "",
    onChange: e => {
      if (e.target.value) exec("formatBlock", e.target.value);
      e.target.value = "";
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0421\u0442\u0438\u043B\u044C\u2026"), /*#__PURE__*/React.createElement("option", {
    value: "p"
  }, "\u041E\u0431\u044B\u0447\u043D\u044B\u0439 \u0442\u0435\u043A\u0441\u0442"), /*#__PURE__*/React.createElement("option", {
    value: "h2"
  }, "\u0417\u0430\u0433\u043E\u043B\u043E\u0432\u043E\u043A 2"), /*#__PURE__*/React.createElement("option", {
    value: "h3"
  }, "\u0417\u0430\u0433\u043E\u043B\u043E\u0432\u043E\u043A 3"), /*#__PURE__*/React.createElement("option", {
    value: "blockquote"
  }, "\u0426\u0438\u0442\u0430\u0442\u0430")), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    title: "\u0428\u0440\u0438\u0444\u0442",
    defaultValue: "",
    onChange: e => {
      if (e.target.value) exec("fontName", e.target.value);
      e.target.value = "";
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0428\u0440\u0438\u0444\u0442\u2026"), /*#__PURE__*/React.createElement("option", {
    value: "'PT Sans',sans-serif"
  }, "PT Sans"), /*#__PURE__*/React.createElement("option", {
    value: "'Playfair Display',serif"
  }, "Playfair Display"), /*#__PURE__*/React.createElement("option", {
    value: "Georgia,serif"
  }, "Georgia"), /*#__PURE__*/React.createElement("option", {
    value: "'Courier New',monospace"
  }, "Courier New")), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    title: "\u0420\u0430\u0437\u043C\u0435\u0440",
    defaultValue: "",
    onChange: e => {
      if (e.target.value) exec("fontSize", e.target.value);
      e.target.value = "";
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0420\u0430\u0437\u043C\u0435\u0440\u2026"), /*#__PURE__*/React.createElement("option", {
    value: "2"
  }, "\u041C\u0435\u043B\u043A\u0438\u0439"), /*#__PURE__*/React.createElement("option", {
    value: "3"
  }, "\u041E\u0431\u044B\u0447\u043D\u044B\u0439"), /*#__PURE__*/React.createElement("option", {
    value: "5"
  }, "\u041A\u0440\u0443\u043F\u043D\u044B\u0439"), /*#__PURE__*/React.createElement("option", {
    value: "7"
  }, "\u041E\u0447\u0435\u043D\u044C \u043A\u0440\u0443\u043F\u043D\u044B\u0439")), /*#__PURE__*/React.createElement("div", {
    className: "sep"
  }), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0421\u043F\u0438\u0441\u043E\u043A",
    onClick: () => exec("insertUnorderedList")
  }, "\u2022\u2261"), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u041D\u0443\u043C\u0435\u0440\u043E\u0432\u0430\u043D\u043D\u044B\u0439 \u0441\u043F\u0438\u0441\u043E\u043A",
    onClick: () => exec("insertOrderedList")
  }, "1\u2261"), /*#__PURE__*/React.createElement("input", {
    type: "color",
    title: "\u0426\u0432\u0435\u0442 \u0442\u0435\u043A\u0441\u0442\u0430",
    onChange: e => exec("foreColor", e.target.value),
    style: {
      width: 26,
      height: 26,
      border: "none",
      background: "none",
      padding: 0
    }
  }), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u041E\u0447\u0438\u0441\u0442\u0438\u0442\u044C \u0444\u043E\u0440\u043C\u0430\u0442\u0438\u0440\u043E\u0432\u0430\u043D\u0438\u0435",
    onClick: () => exec("removeFormat")
  }, "\u232B"), /*#__PURE__*/React.createElement("div", {
    className: "sep"
  }), /*#__PURE__*/React.createElement("div", {
    className: "hlrow",
    title: "\u041C\u0430\u0440\u043A\u0435\u0440"
  }, HIGHLIGHT_COLORS.map(c => /*#__PURE__*/React.createElement("div", {
    key: c,
    className: "hlsw",
    style: {
      background: c
    },
    onMouseDown: e => e.preventDefault(),
    onClick: () => execHighlight(c)
  })), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0423\u0431\u0440\u0430\u0442\u044C \u043C\u0430\u0440\u043A\u0435\u0440",
    onClick: () => execHighlight("transparent")
  }, "\u2715")), /*#__PURE__*/React.createElement("div", {
    className: "sep"
  }), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0422\u0430\u0431\u043B\u0438\u0446\u0430",
    onClick: insertTable
  }, "\u25A6"), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "+ \u0441\u0442\u0440\u043E\u043A\u0430",
    onClick: addRow
  }, "\u25A6+\u0440"), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "+ \u0441\u0442\u043E\u043B\u0431\u0435\u0446",
    onClick: addCol
  }, "\u25A6+\u0441"), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u2212 \u0441\u0442\u0440\u043E\u043A\u0430",
    onClick: delRow
  }, "\u25A6\u2212\u0440"), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u2212 \u0441\u0442\u043E\u043B\u0431\u0435\u0446",
    onClick: delCol
  }, "\u25A6\u2212\u0441"), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C \u0442\u0430\u0431\u043B\u0438\u0446\u0443",
    onClick: delTable
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  })), /*#__PURE__*/React.createElement("div", {
    className: "sep"
  }), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0418\u0437\u043E\u0431\u0440\u0430\u0436\u0435\u043D\u0438\u0435",
    onClick: () => insertMedia("image")
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "image",
    size: 14
  })), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0410\u0443\u0434\u0438\u043E",
    onClick: () => insertMedia("audio")
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "music",
    size: 14
  })), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0412\u0438\u0434\u0435\u043E",
    onClick: () => insertMedia("video")
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "video",
    size: 14
  })), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0424\u0430\u0439\u043B",
    onClick: () => insertMedia("any")
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 14
  })), onAddNoteClick && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "sep"
  }), /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0437\u0430\u043C\u0435\u0442\u043A\u0443 \u043A \u0432\u044B\u0434\u0435\u043B\u0435\u043D\u0438\u044E",
    onClick: onAddNoteClick
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "note",
    size: 14
  }), " \u0417\u0430\u043C\u0435\u0442\u043A\u0430")), update && sourceInfo && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(ToolBtn, {
    title: "\u041E\u0442\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u0432\u044B\u0434\u0435\u043B\u0435\u043D\u0438\u0435 \u0432 \xAB\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B \u0434\u043B\u044F \u0441\u0442\u0430\u0442\u044C\u0438\xBB",
    onClick: sendToMaterials
  }, "\uD83D\uDCF0 \u0412 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B"))), /*#__PURE__*/React.createElement("div", {
    ref: editorRef,
    className: "editor-area" + (compact ? " compact" : ""),
    contentEditable: true,
    suppressContentEditableWarning: true,
    "data-placeholder": placeholder || "",
    onInput: saveHtml,
    onBlur: saveHtml,
    onMouseUp: onSelectionCapture,
    onKeyUp: onSelectionCapture,
    onClick: handleEditorClick
  }));
});
;
// ============================================================
// tagUI.js — выбор/создание тегов, экран управления тегами,
// панель "Ход урока" (вопросный план + заметки + подсветка).
// ============================================================

function TagPicker({
  data,
  update,
  ns,
  selectedIds,
  onChange,
  placeholder,
  allowCreate = true
}) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const tags = data.tags[ns] || [];
  const selected = tags.filter(t => selectedIds.includes(t.id));
  const rest = tags.filter(t => !selectedIds.includes(t.id));
  const matches = q ? rest.filter(t => t.name.toLowerCase().includes(q.toLowerCase())) : rest.slice(0, 8);
  const add = id => {
    onChange([...selectedIds, id]);
    setQ("");
  };
  const createNew = () => {
    const id = uid();
    const color = SWATCHES[Math.floor(Math.random() * SWATCHES.length)];
    update(d => {
      d.tags[ns].push({
        id,
        name: q.trim(),
        color
      });
      return d;
    });
    onChange([...selectedIds, id]);
    setQ("");
  };
  return /*#__PURE__*/React.createElement("div", null, selected.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 6
    }
  }, selected.map(t => /*#__PURE__*/React.createElement(Badge, {
    key: t.id,
    tag: t,
    onRemove: () => onChange(selectedIds.filter(x => x !== t.id))
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative"
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: q,
    placeholder: placeholder || "Добавить тег…",
    onChange: e => {
      setQ(e.target.value);
      setOpen(true);
    },
    onFocus: () => setOpen(true),
    onBlur: () => setTimeout(() => setOpen(false), 150),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      width: "100%"
    }
  }), open && (matches.length > 0 || allowCreate && q.trim()) && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "absolute",
      zIndex: 20,
      top: "100%",
      left: 0,
      right: 0,
      background: "var(--surface)",
      border: "1px solid var(--line)",
      borderRadius: 8,
      boxShadow: "var(--shadow)",
      marginTop: 4,
      maxHeight: 180,
      overflow: "auto"
    }
  }, matches.map(t => /*#__PURE__*/React.createElement("div", {
    key: t.id,
    style: {
      padding: "7px 10px",
      cursor: "pointer"
    },
    onMouseDown: () => add(t.id)
  }, /*#__PURE__*/React.createElement(Badge, {
    tag: t
  }))), allowCreate && q.trim() && !tags.some(t => t.name.toLowerCase() === q.trim().toLowerCase()) && /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "7px 10px",
      cursor: "pointer",
      color: "var(--accent)",
      fontWeight: 600
    },
    onMouseDown: createNew
  }, "+ \u0421\u043E\u0437\u0434\u0430\u0442\u044C \xAB", q.trim(), "\xBB"))));
}
function SingleTagSelect({
  data,
  update,
  ns,
  value,
  onChange,
  placeholder
}) {
  const tags = data.tags[ns] || [];
  return /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: value || "",
    onChange: async e => {
      const v = e.target.value;
      if (v === "__new__") {
        const name = await PromptStr("Название нового тега:");
        if (name && name.trim()) {
          const id = uid();
          const color = SWATCHES[Math.floor(Math.random() * SWATCHES.length)];
          update(d => {
            d.tags[ns].push({
              id,
              name: name.trim(),
              color
            });
            return d;
          });
          onChange(id);
        }
      } else onChange(v || null);
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, placeholder || "Не выбрано"), tags.map(t => /*#__PURE__*/React.createElement("option", {
    key: t.id,
    value: t.id
  }, t.name)), /*#__PURE__*/React.createElement("option", {
    value: "__new__"
  }, "+ \u0421\u043E\u0437\u0434\u0430\u0442\u044C \u043D\u043E\u0432\u044B\u0439\u2026"));
}
const TAG_NS_LABELS = {
  genre: {
    title: "Род литературы",
    hint: "Лирика / проза / драма — базовые нельзя удалить, остальные можно добавлять и удалять"
  },
  literaryGenre: {
    title: "Жанр литературы",
    hint: "Ода, повесть, комедия и т.д. — полностью в вашем управлении"
  },
  questionType: {
    title: "Вид вопроса",
    hint: "Используется в «Ходе урока» и в упражнениях"
  },
  classTag: {
    title: "Классы",
    hint: "Пригодится для упражнений и уроков"
  },
  tropeTag: {
    title: "Тропы",
    hint: "Метафора, эпитет, олицетворение и т.д. — для упражнения «Прогулка по тропам»"
  },
  lessonType: {
    title: "Тип урока",
    hint: "Используется в карточке урока"
  },
  literaryDirection: {
    title: "Литературное направление",
    hint: "Классицизм, романтизм, реализм и т.д."
  }
};
function TagAdminGroup({
  data,
  update,
  ns
}) {
  const info = TAG_NS_LABELS[ns];
  const [name, setName] = useState("");
  const [color, setColor] = useState(SWATCHES[0]);
  const tags = data.tags[ns] || [];
  const add = () => {
    if (!name.trim()) return;
    update(d => {
      d.tags[ns].push({
        id: uid(),
        name: name.trim(),
        color
      });
      return d;
    });
    setName("");
  };
  const remove = async id => {
    const t = tags.find(x => x.id === id);
    if (t && t.locked) return;
    if (!(await Confirm(`Удалить тег «${t ? t.name : ""}»? Он будет снят со всех карточек.`))) return;
    update(d => {
      removeTagCascade(d, ns, id);
      return d;
    });
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-head"
  }, /*#__PURE__*/React.createElement("h2", null, info.title)), /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, info.hint), tags.map(t => /*#__PURE__*/React.createElement("div", {
    key: t.id,
    className: "tk-row"
  }, /*#__PURE__*/React.createElement("span", {
    className: "tk-dot",
    style: {
      background: t.color
    }
  }), /*#__PURE__*/React.createElement("span", {
    className: "tk-name"
  }, t.name), t.locked ? /*#__PURE__*/React.createElement("span", {
    className: "badge-lock"
  }, "\uD83D\uDD12 \u043F\u043E\u0441\u0442\u043E\u044F\u043D\u043D\u044B\u0439") : /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => remove(t.id)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  })))), /*#__PURE__*/React.createElement("div", {
    className: "addbox",
    style: {
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: name,
    onChange: e => setName(e.target.value),
    placeholder: "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u043D\u043E\u0432\u043E\u0433\u043E \u0442\u0435\u0433\u0430",
    onKeyDown: e => {
      if (e.key === "Enter") add();
    }
  }), /*#__PURE__*/React.createElement(ColorPicker, {
    value: color,
    onChange: setColor
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: add
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C"))));
}
function TagsAdminTab({
  data,
  update
}) {
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement(TagAdminGroup, {
    data: data,
    update: update,
    ns: "genre"
  }), /*#__PURE__*/React.createElement(TagAdminGroup, {
    data: data,
    update: update,
    ns: "literaryGenre"
  }), /*#__PURE__*/React.createElement(TagAdminGroup, {
    data: data,
    update: update,
    ns: "questionType"
  }), /*#__PURE__*/React.createElement(TagAdminGroup, {
    data: data,
    update: update,
    ns: "classTag"
  }), /*#__PURE__*/React.createElement(TagAdminGroup, {
    data: data,
    update: update,
    ns: "tropeTag"
  }), /*#__PURE__*/React.createElement(TagAdminGroup, {
    data: data,
    update: update,
    ns: "lessonType"
  }), /*#__PURE__*/React.createElement(TagAdminGroup, {
    data: data,
    update: update,
    ns: "literaryDirection"
  }));
}

// ---- Ход урока: вопросный план + заметки + подсветка вопрос→фрагмент ----
function CourseOfLesson({
  data,
  update,
  course,
  onChange,
  projectPath,
  editorRef,
  sourceInfo
}) {
  const [notesHidden, setNotesHidden] = useState(false);
  const [hoveredQ, setHoveredQ] = useState(null);
  const [newQ, setNewQ] = useState("");
  const setLecture = html => onChange(c => ({
    ...c,
    lectureHtml: html
  }));
  const setField = patch => onChange(c => ({
    ...c,
    ...patch
  }));
  const highlightAnchor = (attr, id, cls, on) => {
    const container = editorRef.current && editorRef.current.getContainer();
    if (!container) return;
    container.querySelectorAll(`[${attr}="${CSS.escape(id)}"]`).forEach(el => el.classList.toggle(cls, on));
  };
  const highlightQuestion = (qId, on) => highlightAnchor("data-qa-id", qId, "qa-highlight", on);
  const highlightNote = (noteId, on) => highlightAnchor("data-note-id", noteId, "note-highlight", on);
  const addQuestion = () => {
    if (!newQ.trim()) return;
    const q = {
      id: uid(),
      text: newQ.trim(),
      order: (course.questions || []).length,
      questionTypeTagId: null
    };
    setField({
      questions: [...(course.questions || []), q]
    });
    setNewQ("");
  };
  const updateQuestion = (id, patch) => setField({
    questions: course.questions.map(q => q.id === id ? {
      ...q,
      ...patch
    } : q)
  });
  const removeQuestion = async id => {
    if (!(await Confirm("Удалить вопрос?"))) return;
    setField({
      questions: course.questions.filter(q => q.id !== id),
      notes: (course.notes || []).filter(n => !(n.target?.kind === "question" && n.target.questionId === id))
    });
  };
  const moveQuestion = (id, dir) => {
    const qs = [...course.questions].sort((a, b) => a.order - b.order);
    const idx = qs.findIndex(q => q.id === id);
    const swapIdx = idx + dir;
    if (swapIdx < 0 || swapIdx >= qs.length) return;
    const tmp = qs[idx].order;
    qs[idx].order = qs[swapIdx].order;
    qs[swapIdx].order = tmp;
    setField({
      questions: qs
    });
  };
  const linkToText = qId => {
    if (!editorRef.current || !editorRef.current.hasSelection()) {
      showToast("Сначала выделите фрагмент в тексте слева");
      return;
    }
    editorRef.current.wrapSelection("lc-qa-anchor", {
      "data-qa-id": qId,
      "data-anchor-id": uid()
    });
  };
  const addNoteToQuestion = async qId => {
    const text = await PromptStr("Текст заметки к вопросу:");
    if (text == null || !text.trim()) return;
    setField({
      notes: [...(course.notes || []), {
        id: uid(),
        text: text.trim(),
        target: {
          kind: "question",
          questionId: qId
        }
      }]
    });
  };
  const addStandaloneNote = async () => {
    const text = await PromptStr("Текст заметки:");
    if (text == null || !text.trim()) return;
    setField({
      notes: [...(course.notes || []), {
        id: uid(),
        text: text.trim(),
        target: {
          kind: "standalone"
        }
      }]
    });
  };
  const onAddNoteFromSelection = async () => {
    if (!editorRef.current || !editorRef.current.hasSelection()) {
      showToast("Сначала выделите фрагмент текста");
      return;
    }
    const text = await PromptStr("Текст заметки к выделенному фрагменту:");
    if (text == null || !text.trim()) return;
    const noteId = uid();
    const span = editorRef.current.wrapSelection("lc-note", {
      "data-note-id": noteId
    });
    if (!span) return;
    setField({
      notes: [...(course.notes || []), {
        id: noteId,
        text: text.trim(),
        target: {
          kind: "span",
          anchorId: noteId
        }
      }]
    });
  };
  const removeNote = id => setField({
    notes: (course.notes || []).filter(n => n.id !== id)
  });
  const onNoteAnchorClick = noteId => {
    const n = (course.notes || []).find(x => x.id === noteId);
    if (n) showToast("Заметка: " + n.text);
  };
  const linkNoteToText = noteId => {
    if (!editorRef.current || !editorRef.current.hasSelection()) {
      showToast("Сначала выделите фрагмент в тексте слева");
      return;
    }
    const span = editorRef.current.wrapSelection("lc-note", {
      "data-note-id": noteId
    });
    if (!span) return;
    setField({
      notes: course.notes.map(n => n.id === noteId ? {
        ...n,
        target: {
          kind: "span",
          anchorId: noteId
        }
      } : n)
    });
  };
  const unlinkNote = noteId => {
    const container = editorRef.current && editorRef.current.getContainer();
    if (container && unwrapSpanById(container, "data-note-id", noteId)) setLecture(container.innerHTML);
    setField({
      notes: course.notes.map(n => n.id === noteId ? {
        ...n,
        target: {
          kind: "standalone"
        }
      } : n)
    });
  };
  const unlinkFragment = (qId, anchorId) => {
    const container = editorRef.current && editorRef.current.getContainer();
    if (!container) return;
    const removed = anchorId ? unwrapSpanById(container, "data-anchor-id", anchorId) : unwrapSpanById(container, "data-qa-id", qId);
    if (removed) setLecture(container.innerHTML);
  };
  const sortedQuestions = [...(course.questions || [])].sort((a, b) => a.order - b.order);
  const sideNotes = (course.notes || []).filter(n => n.target?.kind !== "question");
  const questionAnswerFragment = qId => qaFragmentsOf(course.lectureHtml, qId).map(f => f.text).join(" … ");
  const buildWorksheetHtml = forTeacher => {
    if (sortedQuestions.length === 0) return "<p><i>Вопросы ещё не добавлены.</i></p>";
    return sortedQuestions.map((q, i) => {
      let html = `<div class="q"><div class="qnum">${i + 1}. ${escapeHtml(q.text)}</div>`;
      if (forTeacher) {
        const frag = questionAnswerFragment(q.id);
        if (frag) html += `<div class="answer">Ответ (из текста): ${escapeHtml(frag)}</div>`;
        (course.notes || []).filter(n => n.target?.kind === "question" && n.target.questionId === q.id).forEach(n => {
          html += `<div class="note">Примечание: ${escapeHtml(n.text)}</div>`;
        });
      } else {
        html += `<div style="border-bottom:1px solid #999;height:44px;margin-top:6px;"></div>`;
      }
      html += `</div>`;
      return html;
    }).join("");
  };
  const exportWorksheet = forTeacher => {
    const title = (sourceInfo?.entityTitle || "Рабочий лист").trim() || "Рабочий лист";
    const suffix = forTeacher ? "учитель" : "ученик";
    downloadWord(`${title} — рабочий лист (${suffix}).doc`, title, buildWorksheetHtml(forTeacher));
  };
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      justifyContent: "flex-end",
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportWorksheet(false),
    title: "\u0412\u043E\u043F\u0440\u043E\u0441\u044B \u0431\u0435\u0437 \u043E\u0442\u0432\u0435\u0442\u043E\u0432 \u2014 \u0434\u043B\u044F \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u041B\u0438\u0441\u0442 \u0434\u043B\u044F \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportWorksheet(true),
    title: "\u0412\u043E\u043F\u0440\u043E\u0441\u044B \u0441 \u043E\u0442\u0432\u0435\u0442\u0430\u043C\u0438 \u0438\u0437 \u0442\u0435\u043A\u0441\u0442\u0430 \u0438 \u043F\u0440\u0438\u043C\u0435\u0447\u0430\u043D\u0438\u044F\u043C\u0438 \u2014 \u0434\u043B\u044F \u0443\u0447\u0438\u0442\u0435\u043B\u044F"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u041B\u0438\u0441\u0442 \u0434\u043B\u044F \u0443\u0447\u0438\u0442\u0435\u043B\u044F"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setNotesHidden(h => !h)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: notesHidden ? "eye" : "eyeOff",
    size: 14
  }), " ", notesHidden ? "Показать заметки и примечания" : "Скрыть заметки и примечания")), /*#__PURE__*/React.createElement("div", {
    className: "course-grid"
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 6
    }
  }, "\u041B\u0435\u043A\u0446\u0438\u044F / \u0442\u0435\u043A\u0441\u0442"), /*#__PURE__*/React.createElement(RichEditor, {
    html: course.lectureHtml,
    onChange: setLecture,
    projectPath: projectPath,
    ref: editorRef,
    notesHidden: notesHidden,
    onAddNoteClick: onAddNoteFromSelection,
    onNoteAnchorClick: onNoteAnchorClick,
    update: update,
    sourceInfo: sourceInfo
  })), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 6
    }
  }, "\u0425\u043E\u0434 \u0443\u0440\u043E\u043A\u0430 \u2014 \u0432\u043E\u043F\u0440\u043E\u0441\u043D\u044B\u0439 \u043F\u043B\u0430\u043D"), /*#__PURE__*/React.createElement("div", {
    className: "panel",
    style: {
      marginBottom: 12
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, sortedQuestions.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "muted",
    style: {
      fontSize: 13
    }
  }, "\u0412\u043E\u043F\u0440\u043E\u0441\u043E\u0432 \u043F\u043E\u043A\u0430 \u043D\u0435\u0442."), sortedQuestions.map((q, i) => {
    const fragments = qaFragmentsOf(course.lectureHtml, q.id);
    return /*#__PURE__*/React.createElement("div", {
      key: q.id,
      className: "qrow" + (hoveredQ === q.id ? " hovered" : ""),
      onMouseEnter: () => {
        setHoveredQ(q.id);
        highlightQuestion(q.id, true);
      },
      onMouseLeave: () => {
        setHoveredQ(null);
        highlightQuestion(q.id, false);
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row",
      style: {
        alignItems: "flex-start"
      }
    }, /*#__PURE__*/React.createElement("span", {
      className: "muted",
      style: {
        fontSize: 12,
        minWidth: 18
      }
    }, i + 1, "."), /*#__PURE__*/React.createElement(AutoTextarea, {
      value: q.text,
      onChange: v => updateQuestion(q.id, {
        text: v
      }),
      style: {
        flex: 1,
        minWidth: 0,
        border: "none",
        background: "transparent",
        fontWeight: 600,
        fontSize: 13.5,
        fontFamily: "inherit",
        lineHeight: 1.4,
        padding: "2px 0"
      }
    }), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => moveQuestion(q.id, -1),
      title: "\u0412\u044B\u0448\u0435"
    }, "\u2191"), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => moveQuestion(q.id, 1),
      title: "\u041D\u0438\u0436\u0435"
    }, "\u2193"), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => removeQuestion(q.id),
      title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "trash",
      size: 13
    }))), /*#__PURE__*/React.createElement("div", {
      className: "qmeta"
    }, /*#__PURE__*/React.createElement(SingleTagSelect, {
      data: data,
      update: update,
      ns: "questionType",
      value: q.questionTypeTagId,
      onChange: v => updateQuestion(q.id, {
        questionTypeTagId: v
      }),
      placeholder: "\u0412\u0438\u0434 \u0432\u043E\u043F\u0440\u043E\u0441\u0430\u2026"
    }), /*#__PURE__*/React.createElement(Btn, {
      kind: fragments.length > 0 ? "gold" : "ghost",
      sm: true,
      onClick: () => linkToText(q.id),
      title: "\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0444\u0440\u0430\u0433\u043C\u0435\u043D\u0442 \u043B\u0435\u043A\u0446\u0438\u0438 \u0441\u043B\u0435\u0432\u0430 \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "link",
      size: 12
    }), " ", fragments.length > 0 ? `Привязано: ${fragments.length}` : "Привязать фрагмент"), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => addNoteToQuestion(q.id),
      title: "\u0417\u0430\u043C\u0435\u0442\u043A\u0430 \u043A \u0432\u043E\u043F\u0440\u043E\u0441\u0443"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "note",
      size: 12
    }))), fragments.length > 0 && /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        marginTop: 6,
        gap: 4
      }
    }, fragments.map((f, fi) => /*#__PURE__*/React.createElement("span", {
      key: f.anchorId || fi,
      className: "chip",
      style: {
        fontSize: 11
      }
    }, f.text.slice(0, 26), f.text.length > 26 ? "…" : "", /*#__PURE__*/React.createElement("span", {
      className: "x",
      onClick: () => unlinkFragment(q.id, f.anchorId)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "close",
      size: 10
    }))))), (course.notes || []).filter(n => n.target?.kind === "question" && n.target.questionId === q.id).map(n => /*#__PURE__*/React.createElement("div", {
      key: n.id,
      className: "noterow",
      style: {
        marginTop: 6
      }
    }, n.text, /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => removeNote(n.id),
      title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "close",
      size: 11
    })))));
  }), /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      marginTop: 8
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: newQ,
    onChange: e => setNewQ(e.target.value),
    placeholder: "\u041D\u043E\u0432\u044B\u0439 \u0432\u043E\u043F\u0440\u043E\u0441\u2026",
    onKeyDown: e => {
      if (e.key === "Enter") addQuestion();
    },
    style: {
      flex: 1,
      minWidth: 0,
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px"
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: addQuestion
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C")))), /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-head"
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      flex: 1
    }
  }, "\u0417\u0430\u043C\u0435\u0442\u043A\u0438"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: addStandaloneNote
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u0410\u0432\u0442\u043E\u043D\u043E\u043C\u043D\u0430\u044F")), /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, sideNotes.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "muted",
    style: {
      fontSize: 13
    }
  }, "\u0417\u0430\u043C\u0435\u0442\u043E\u043A \u043F\u043E\u043A\u0430 \u043D\u0435\u0442."), sideNotes.map(n => {
    const linked = n.target?.kind === "span";
    return /*#__PURE__*/React.createElement("div", {
      key: n.id,
      className: "noterow",
      onMouseEnter: () => linked && highlightNote(n.id, true),
      onMouseLeave: () => linked && highlightNote(n.id, false)
    }, /*#__PURE__*/React.createElement("div", {
      className: "row",
      style: {
        alignItems: "flex-start"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1
      }
    }, n.text), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => removeNote(n.id)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "close",
      size: 11
    }))), /*#__PURE__*/React.createElement("div", {
      style: {
        marginTop: 6
      }
    }, /*#__PURE__*/React.createElement(Btn, {
      kind: linked ? "gold" : "ghost",
      sm: true,
      onClick: () => linked ? unlinkNote(n.id) : linkNoteToText(n.id),
      title: linked ? "Отвязать от текста" : "Выделите фрагмент лекции слева и нажмите"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "link",
      size: 12
    }), " ", linked ? "Отвязать от текста" : "Привязать к тексту")));
  }))))));
}
;
// ============================================================
// exercises.js — Упражнения: категории → виды (теги) → карточки,
// слоты («цепочки»), привязка к урокам.
// ============================================================

function AddToLessonPicker({
  data,
  update,
  itemType,
  refId
}) {
  const [scope, setScope] = useState("");
  const [lessonId, setLessonId] = useState("");
  const lessonsInScope = () => {
    if (!scope) return [];
    if (scope.startsWith("class:")) {
      const c = scope.slice(6);
      return data.lessons.filter(l => l.class === c && !l.archived);
    }
    const cid = scope.slice(7);
    return data.lessons.filter(l => l.courseId === cid && !l.archived);
  };
  const add = async () => {
    if (!lessonId) return;
    const l0 = data.lessons.find(x => x.id === lessonId);
    if (l0 && ITEM_TYPE_META[itemType]) {
      const ent0 = (data[ITEM_TYPE_META[itemType].dataKey] || []).find(e => e.id === refId);
      const why = materialHiddenReason(data, itemType, ent0, {
        class: l0.class,
        courseId: l0.courseId,
        classLetterId: l0.classLetterId || null
      }, l0.id);
      if (why && why !== "audience" && !(await Confirm("Эта карточка для «" + scopeLabelOf(data, scopeKeyOf(l0)) + "» — " + hiddenReasonText(why, scopeLabelOf(data, scopeKeyOf(l0))) + ". Всё равно добавить в этот урок?", "Да, добавить"))) return;
    }
    update(d => {
      const l = d.lessons.find(x => x.id === lessonId);
      if (l && !l.items.some(it => it.type === itemType && it.refId === refId)) l.items.push({
        id: uid(),
        type: itemType,
        refId,
        note: ""
      });
      return d;
    });
    showToast("Добавлено в урок");
  };
  return /*#__PURE__*/React.createElement(Field, {
    label: "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0432 \u0443\u0440\u043E\u043A"
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: scope,
    onChange: e => {
      setScope(e.target.value);
      setLessonId("");
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u041A\u043B\u0430\u0441\u0441 / \u043A\u0443\u0440\u0441\u2026"), CLASSES.map(c => /*#__PURE__*/React.createElement("option", {
    key: c,
    value: "class:" + c
  }, c, " \u043A\u043B\u0430\u0441\u0441")), data.courses.map(c => /*#__PURE__*/React.createElement("option", {
    key: c.id,
    value: "course:" + c.id
  }, c.name))), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: lessonId,
    onChange: e => setLessonId(e.target.value),
    disabled: !scope
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0423\u0440\u043E\u043A\u2026"), lessonsInScope().map(l => /*#__PURE__*/React.createElement("option", {
    key: l.id,
    value: l.id
  }, (l.class || "") + classLetterName(data, l.classLetterId), l.plannedDate ? " · " + formatDateRu(l.plannedDate).slice(0, 5) : "", " \xB7 ", l.topic || "Без темы", l.duplicateOf ? " (дубликат)" : ""))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: add,
    disabled: !lessonId
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C")));
}
function StudentPicker({
  data,
  update,
  cls,
  setCls,
  studentIds,
  onChange
}) {
  const [q, setQ] = useState("");
  const [open, setOpen] = useState(false);
  const roster = data.calendar.students.filter(s => s.class === cls);
  const assigned = studentIds.map(id => data.calendar.students.find(s => s.id === id)).filter(Boolean);
  const available = roster.filter(s => !studentIds.includes(s.id));
  const matches = q ? available.filter(s => s.name.toLowerCase().includes(q.toLowerCase())) : available;
  const add = id => {
    onChange([...studentIds, id]);
    setQ("");
  };
  const createNew = () => {
    const id = uid();
    update(d => {
      d.calendar.students.push({
        id,
        name: q.trim(),
        class: cls
      });
      return d;
    });
    onChange([...studentIds, id]);
    setQ("");
  };
  const pickRandom = () => {
    if (available.length === 0) {
      showToast(`В ${cls} классе нет свободных учеников — добавьте кого-то в список`);
      return;
    }
    const chosen = available[Math.floor(Math.random() * available.length)];
    add(chosen.id);
    showToast("Назначен: " + chosen.name);
  };
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 8
    }
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: cls,
    onChange: e => setCls(e.target.value)
  }, CLASSES.map(c => /*#__PURE__*/React.createElement("option", {
    key: c,
    value: c
  }, c, " \u043A\u043B\u0430\u0441\u0441"))), /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    sm: true,
    onClick: pickRandom,
    title: "\u041D\u0430\u0437\u043D\u0430\u0447\u0438\u0442\u044C \u0441\u043B\u0443\u0447\u0430\u0439\u043D\u043E\u0433\u043E \u0443\u0447\u0435\u043D\u0438\u043A\u0430 \u0438\u0437 \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u043E\u0433\u043E \u043A\u043B\u0430\u0441\u0441\u0430"
  }, "\uD83C\uDFB2 \u0421\u043B\u0443\u0447\u0430\u0439\u043D\u044B\u0439 \u0443\u0447\u0435\u043D\u0438\u043A")), assigned.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 6
    }
  }, assigned.map(s => /*#__PURE__*/React.createElement("span", {
    key: s.id,
    className: "chip"
  }, s.name, " \xB7 ", s.class, " \u043A\u043B.", /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => onChange(studentIds.filter(x => x !== s.id))
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  }))))), /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative"
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: q,
    placeholder: `Добавить ученика из ${cls} класса…`,
    onChange: e => {
      setQ(e.target.value);
      setOpen(true);
    },
    onFocus: () => setOpen(true),
    onBlur: () => setTimeout(() => setOpen(false), 150),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      width: "100%"
    }
  }), open && (matches.length > 0 || q.trim()) && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "absolute",
      zIndex: 20,
      top: "100%",
      left: 0,
      right: 0,
      background: "var(--surface)",
      border: "1px solid var(--line)",
      borderRadius: 8,
      boxShadow: "var(--shadow)",
      marginTop: 4,
      maxHeight: 180,
      overflow: "auto"
    }
  }, matches.map(s => /*#__PURE__*/React.createElement("div", {
    key: s.id,
    style: {
      padding: "7px 10px",
      cursor: "pointer"
    },
    onMouseDown: () => add(s.id)
  }, s.name)), q.trim() && !roster.some(s => s.name.toLowerCase() === q.trim().toLowerCase()) && /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "7px 10px",
      cursor: "pointer",
      color: "var(--accent)",
      fontWeight: 600
    },
    onMouseDown: createNew
  }, "+ \u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \xAB", q.trim(), "\xBB \u0432 ", cls, " \u043A\u043B\u0430\u0441\u0441"))));
}
function ExampleForm({
  data,
  update,
  projectPath,
  typeId,
  editId,
  onDone
}) {
  const existing = editId ? data.examples.find(e => e.id === editId) : null;
  const isTropeWalk = typeId === TROPE_WALK_TYPE_ID;
  const [text, setText] = useState(existing?.text || "");
  const [note, setNote] = useState(existing?.note || "");
  const [qType, setQType] = useState(existing?.questionTypeTagId || null);
  const [tagIds, setTagIds] = useState(existing?.tagIds || []);
  const [studentIds, setStudentIds] = useState(existing?.assignedStudentIds || []);
  const [cls, setCls] = useState(CLASSES[0]);
  const existingAuthor = existing?.authorId ? data.personalities.find(p => p.id === existing.authorId) : null;
  const [authorName, setAuthorName] = useState(existingAuthor?.title || "");
  const [workTitle, setWorkTitle] = useState(existing?.workTitle || "");
  const [forClasses, setForClasses] = useState(existing?.forClasses || []);
  const [forCourseIds, setForCourseIds] = useState(existing?.forCourseIds || []);
  const save = () => {
    let createdId = null;
    if (existing) {
      update(d => {
        const idx = d.examples.findIndex(e => e.id === editId);
        const authorId = isTropeWalk ? ensurePersonality(d, authorName) : d.examples[idx].authorId;
        d.examples[idx] = {
          ...d.examples[idx],
          text,
          note,
          questionTypeTagId: qType,
          tagIds,
          assignedStudentIds: studentIds,
          forClasses,
          forCourseIds,
          updatedAt: nowISO(),
          ...(isTropeWalk ? {
            authorId,
            workTitle
          } : {})
        };
        return d;
      });
    } else {
      update(d => {
        const authorId = isTropeWalk ? ensurePersonality(d, authorName) : null;
        const ex = migrateExample({
          typeTagId: typeId,
          text,
          note,
          questionTypeTagId: qType,
          tagIds,
          assignedStudentIds: studentIds,
          forClasses,
          forCourseIds,
          authorId,
          workTitle: isTropeWalk ? workTitle : ""
        });
        d.examples.push(ex);
        createdId = ex.id;
        return d;
      });
    }
    onDone(createdId);
  };
  return /*#__PURE__*/React.createElement("div", null, isTropeWalk && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 18,
      alignItems: "flex-start",
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 220,
      flex: 1
    }
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u0410\u0432\u0442\u043E\u0440"
  }, /*#__PURE__*/React.createElement("input", {
    list: "trope-walk-authors",
    value: authorName,
    onChange: e => setAuthorName(e.target.value),
    placeholder: "\u0424.\u0418.\u041E. \u0430\u0432\u0442\u043E\u0440\u0430 \u2014 \u0435\u0441\u043B\u0438 \u0442\u0430\u043A\u043E\u0433\u043E \u043D\u0435\u0442 \u0432 \xAB\u041F\u0435\u0440\u0441\u043E\u043D\u0430\u043B\u0438\u044F\u0445\xBB, \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0430 \u0441\u043E\u0437\u0434\u0430\u0441\u0442\u0441\u044F \u0441\u0430\u043C\u0430",
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      width: "100%"
    }
  }), /*#__PURE__*/React.createElement("datalist", {
    id: "trope-walk-authors"
  }, data.personalities.map(p => /*#__PURE__*/React.createElement("option", {
    key: p.id,
    value: p.title
  }))))), /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 220,
      flex: 1
    }
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u043F\u0440\u043E\u0438\u0437\u0432\u0435\u0434\u0435\u043D\u0438\u044F"
  }, /*#__PURE__*/React.createElement("input", {
    value: workTitle,
    onChange: e => setWorkTitle(e.target.value),
    placeholder: "\xAB\u0424\u043E\u043D\u0442\u0430\u043D\xBB",
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      width: "100%"
    }
  })))), /*#__PURE__*/React.createElement(Field, {
    label: isTropeWalk ? "Текст произведения (в котором нужно искать троп)" : "Текст упражнения"
  }, /*#__PURE__*/React.createElement(RichEditor, {
    html: text,
    onChange: setText,
    projectPath: projectPath,
    compact: true,
    update: update,
    sourceInfo: {
      section: "Упражнения",
      entityId: editId || null,
      entityTitle: (data.tags.exerciseType.find(t => t.id === typeId) || {}).name || "",
      genreTagIds: nsTagIds(tagIds, "genre", data)
    }
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u041F\u0440\u0438\u043C\u0435\u0447\u0430\u043D\u0438\u0435 (\u043D\u0435\u043E\u0431\u044F\u0437\u0430\u0442\u0435\u043B\u044C\u043D\u043E)"
  }, /*#__PURE__*/React.createElement("textarea", {
    value: note,
    onChange: e => setNote(e.target.value),
    rows: 2,
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: 8
    }
  })), !isTropeWalk && /*#__PURE__*/React.createElement(Field, {
    label: "\u0412\u0438\u0434 \u0432\u043E\u043F\u0440\u043E\u0441\u0430"
  }, /*#__PURE__*/React.createElement(SingleTagSelect, {
    data: data,
    update: update,
    ns: "questionType",
    value: qType,
    onChange: setQType
  })), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 18,
      alignItems: "flex-start"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 220,
      flex: 1
    }
  }, /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.genre.title
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "genre",
    selectedIds: nsTagIds(tagIds, "genre", data),
    onChange: ids => setTagIds(withNsTagIds(tagIds, "genre", data, ids))
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 220,
      flex: 1
    }
  }, /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.classTag.title
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "classTag",
    selectedIds: nsTagIds(tagIds, "classTag", data),
    onChange: ids => setTagIds(withNsTagIds(tagIds, "classTag", data, ids))
  }))), isTropeWalk && /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 220,
      flex: 1
    }
  }, /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.tropeTag.title + " (можно несколько)"
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "tropeTag",
    selectedIds: nsTagIds(tagIds, "tropeTag", data),
    onChange: ids => setTagIds(withNsTagIds(tagIds, "tropeTag", data, ids)),
    placeholder: "\u0418\u0441\u043A\u043E\u043C\u044B\u0439 \u0442\u0440\u043E\u043F\u2026"
  })))), /*#__PURE__*/React.createElement(Field, {
    label: "\u0414\u043B\u044F \u043A\u0430\u043A\u0438\u0445 \u043A\u0443\u0440\u0441\u043E\u0432 \u0434\u043E\u0441\u0442\u0443\u043F\u043D\u043E (\u043A\u043B\u0430\u0441\u0441\u044B \u0437\u0430\u0434\u0430\u044E\u0442\u0441\u044F \u043C\u0435\u0442\u043A\u0430\u043C\u0438 \xAB\u041A\u043B\u0430\u0441\u0441\u044B\xBB \u0432\u044B\u0448\u0435; \u043C\u0435\u0442\u043A\u0430 \u043A\u043B\u0430\u0441\u0441\u0430 = \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0435\u0442\u0441\u044F \u0442\u043E\u043B\u044C\u043A\u043E \u0432 \u044D\u0442\u043E\u043C \u043A\u043B\u0430\u0441\u0441\u0435; \u043F\u0443\u0441\u0442\u043E \u2014 \u0432\u0441\u0435\u043C)"
  }, /*#__PURE__*/React.createElement(AudiencePicker, {
    data: data,
    classes: forClasses,
    courseIds: forCourseIds,
    showClasses: false,
    onChange: v => {
      setForClasses(v.classes);
      setForCourseIds(v.courseIds);
    }
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u0423\u0447\u0435\u043D\u0438\u043A\u0438"
  }, /*#__PURE__*/React.createElement(StudentPicker, {
    data: data,
    update: update,
    cls: cls,
    setCls: setCls,
    studentIds: studentIds,
    onChange: setStudentIds
  })), /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      justifyContent: "flex-end",
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => onDone(null)
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: save
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C")));
}

// ---- «Прогулка по тропам»: экспорт одного упражнения на весь экран/лист — белый фон,
// крупный чёрный текст (автор+название сверху, текст произведения по центру). ----
function tropeWalkExportContent(ex, data) {
  const author = ex.authorId ? data.personalities.find(p => p.id === ex.authorId) : null;
  const heading = (author ? author.title : "Автор не указан") + (ex.workTitle ? `, «${ex.workTitle}»` : "");
  const lines = htmlToLines(ex.text).filter(l => l !== "");
  return {
    heading,
    bodyHtml: ex.text || "<p><i>Пусто</i></p>",
    lines
  };
}
function exportTropeWalkWord(ex, data) {
  const {
    heading,
    bodyHtml
  } = tropeWalkExportContent(ex, data);
  const html = `<div style="text-align:center;color:#000;background:#fff;">
    <h1 style="font-size:32px;">${escapeHtml(heading)}</h1>
    <div style="font-size:24px;line-height:1.6;font-weight:bold;">${bodyHtml}</div>
  </div>`;
  downloadWord(safeFileName(heading) + ".doc", heading, html);
}
function exportTropeWalkPdf(ex, data) {
  const {
    heading,
    bodyHtml
  } = tropeWalkExportContent(ex, data);
  printAsPdf(`<div style="min-height:100vh;display:flex;flex-direction:column;justify-content:center;align-items:center;background:#fff;color:#000;font-family:Arial,sans-serif;text-align:center;padding:40px;box-sizing:border-box;">
    <div style="font-size:40px;font-weight:800;margin-bottom:30px;">${escapeHtml(heading)}</div>
    <div style="font-size:28px;line-height:1.7;font-weight:bold;max-width:900px;">${bodyHtml}</div>
  </div>`);
}
function buildTropeWalkPptxXml(ex, data) {
  const {
    heading,
    lines
  } = tropeWalkExportContent(ex, data);
  const textRun = (text, sz) => `<a:r><a:rPr lang="ru-RU" sz="${sz}" b="1"><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Arial"/></a:rPr><a:t>${escapeHtml(text)}</a:t></a:r>`;
  const bodyParas = (lines.length ? lines : [""]).map(line => line ? `<a:p><a:pPr algn="ctr"/>${textRun(line, 2800)}</a:p>` : `<a:p><a:pPr algn="ctr"/><a:endParaRPr lang="ru-RU" sz="2800"/></a:p>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
<p:sp>
<p:nvSpPr><p:cNvPr id="2" name="Title"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="457200" y="274638"/><a:ext cx="11277600" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/>
<a:p><a:pPr algn="ctr"/>${textRun(heading, 4400)}</a:p>
</p:txBody>
</p:sp>
<p:sp>
<p:nvSpPr><p:cNvPr id="3" name="Body"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr/></p:nvSpPr>
<p:spPr><a:xfrm><a:off x="457200" y="1600200"/><a:ext cx="11277600" cy="4886325"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
<p:txBody><a:bodyPr wrap="square" anchor="ctr"><a:normAutofit/></a:bodyPr><a:lstStyle/>
${bodyParas}
</p:txBody>
</p:sp>
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sld>`;
}
const PPTX_CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>
<Override PartName="/ppt/slideLayouts/slideLayout1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>
<Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
</Types>`;
const PPTX_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>`;
const PPTX_CORE_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
<dc:title>Прогулка по тропам</dc:title>
<dc:creator>Конструктор уроков литературы</dc:creator>
<cp:lastModifiedBy>Конструктор уроков литературы</cp:lastModifiedBy>
</cp:coreProperties>`;
const PPTX_APP_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
<Application>Конструктор уроков литературы</Application>
<Slides>1</Slides>
</Properties>`;
const PPTX_PRESENTATION_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>
<p:sldIdLst><p:sldId id="256" r:id="rId2"/></p:sldIdLst>
<p:sldSz cx="12192000" cy="6858000" type="screen16x9"/>
<p:notesSz cx="6858000" cy="9144000"/>
</p:presentation>`;
const PPTX_PRESENTATION_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`;
const PPTX_THEME_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Simple">
<a:themeElements>
<a:clrScheme name="Simple">
<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>
<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="000000"/></a:dk2>
<a:lt2><a:srgbClr val="FFFFFF"/></a:lt2>
<a:accent1><a:srgbClr val="132C53"/></a:accent1>
<a:accent2><a:srgbClr val="B7791F"/></a:accent2>
<a:accent3><a:srgbClr val="2A9D8F"/></a:accent3>
<a:accent4><a:srgbClr val="3B7DBF"/></a:accent4>
<a:accent5><a:srgbClr val="8E4A86"/></a:accent5>
<a:accent6><a:srgbClr val="C1554B"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
</a:clrScheme>
<a:fontScheme name="Simple">
<a:majorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont>
<a:minorFont><a:latin typeface="Arial"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont>
</a:fontScheme>
<a:fmtScheme name="Simple">
<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst>
<a:lnStyleLst><a:ln w="6350"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="12700"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="19050"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst>
<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>
<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst>
</a:fmtScheme>
</a:themeElements>
</a:theme>`;
const PPTX_SLIDE_MASTER_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
<p:cSld>
<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>
</p:cSld>
<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>
<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rId1"/></p:sldLayoutIdLst>
</p:sldMaster>`;
const PPTX_SLIDE_MASTER_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"/>
</Relationships>`;
const PPTX_SLIDE_LAYOUT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="blank" preserve="1">
<p:cSld name="Blank">
<p:spTree>
<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>
</p:spTree>
</p:cSld>
<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>
</p:sldLayout>`;
const PPTX_SLIDE_LAYOUT_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="../slideMasters/slideMaster1.xml"/>
</Relationships>`;
const PPTX_SLIDE_RELS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout1.xml"/>
</Relationships>`;
function exportTropeWalkPptx(ex, data) {
  const {
    heading
  } = tropeWalkExportContent(ex, data);
  const files = [{
    name: "[Content_Types].xml",
    data: strToBytes(PPTX_CONTENT_TYPES_XML)
  }, {
    name: "_rels/.rels",
    data: strToBytes(PPTX_RELS_XML)
  }, {
    name: "docProps/core.xml",
    data: strToBytes(PPTX_CORE_XML)
  }, {
    name: "docProps/app.xml",
    data: strToBytes(PPTX_APP_XML)
  }, {
    name: "ppt/presentation.xml",
    data: strToBytes(PPTX_PRESENTATION_XML)
  }, {
    name: "ppt/_rels/presentation.xml.rels",
    data: strToBytes(PPTX_PRESENTATION_RELS_XML)
  }, {
    name: "ppt/theme/theme1.xml",
    data: strToBytes(PPTX_THEME_XML)
  }, {
    name: "ppt/slideMasters/slideMaster1.xml",
    data: strToBytes(PPTX_SLIDE_MASTER_XML)
  }, {
    name: "ppt/slideMasters/_rels/slideMaster1.xml.rels",
    data: strToBytes(PPTX_SLIDE_MASTER_RELS_XML)
  }, {
    name: "ppt/slideLayouts/slideLayout1.xml",
    data: strToBytes(PPTX_SLIDE_LAYOUT_XML)
  }, {
    name: "ppt/slideLayouts/_rels/slideLayout1.xml.rels",
    data: strToBytes(PPTX_SLIDE_LAYOUT_RELS_XML)
  }, {
    name: "ppt/slides/slide1.xml",
    data: strToBytes(buildTropeWalkPptxXml(ex, data))
  }, {
    name: "ppt/slides/_rels/slide1.xml.rels",
    data: strToBytes(PPTX_SLIDE_RELS_XML)
  }];
  const bytes = makeZip(files);
  const blob = new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = safeFileName(heading) + ".pptx";
  a.click();
  URL.revokeObjectURL(url);
}

// ---- Массовый импорт стихотворений для «Прогулки по тропам» из Word (.docx) / PDF / txt ----
// Файл читается в «строки» (абзац Word = строка; пустой абзац = пропуск между строфами), затем строки
// разбираются по шаблону: перед каждым стихотворением идут строки с метками «Автор:», «Название:»,
// «Класс:», «Тропы:», дальше — сам текст. Новое стихотворение начинается с метки «Автор:»/«Название:»
// (после уже набранного текста) или со строки-разделителя из ***, ---, ===, ###.
async function inflateRawBytes(bytes) {
  const stream = new Blob([bytes]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}
async function readZipEntryBytes(buf, wanted) {
  const u8 = new Uint8Array(buf),
    dv = new DataView(buf);
  let eocd = -1;
  for (let i = u8.length - 22; i >= Math.max(0, u8.length - 70000); i--) {
    if (dv.getUint32(i, true) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) throw new Error("Это не файл .docx (не найден архив)");
  const total = dv.getUint16(eocd + 10, true);
  let pos = dv.getUint32(eocd + 16, true);
  for (let i = 0; i < total; i++) {
    if (dv.getUint32(pos, true) !== 0x02014b50) break;
    const method = dv.getUint16(pos + 10, true),
      compSize = dv.getUint32(pos + 20, true);
    const nameLen = dv.getUint16(pos + 28, true),
      extraLen = dv.getUint16(pos + 30, true),
      commentLen = dv.getUint16(pos + 32, true);
    const local = dv.getUint32(pos + 42, true);
    const name = new TextDecoder().decode(u8.subarray(pos + 46, pos + 46 + nameLen));
    if (name === wanted) {
      const start = local + 30 + dv.getUint16(local + 26, true) + dv.getUint16(local + 28, true);
      const raw = u8.subarray(start, start + compSize);
      return method === 0 ? raw : await inflateRawBytes(raw);
    }
    pos += 46 + nameLen + extraLen + commentLen;
  }
  throw new Error("В архиве нет нужного файла: " + wanted);
}
async function docxToLines(buf) {
  const xml = new TextDecoder("utf-8").decode(await readZipEntryBytes(buf, "word/document.xml"));
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const out = [],
    meta = [];
  const textOf = (node, acc) => {
    node.childNodes.forEach(ch => {
      if (ch.nodeType !== 1) return;
      const n = ch.localName;
      if (n === "t") acc.s += ch.textContent;else if (n === "tab") acc.s += " ";else if (n === "br" || n === "cr") acc.s += "\n";else if (n === "instrText" || n === "delText") return;else textOf(ch, acc);
    });
  };
  const kids = (el, name) => Array.from(el.childNodes).filter(c => c.nodeType === 1 && c.localName === name);
  const flagOn = el => !/^(0|false|off)$/i.test(el.getAttribute("w:val") || el.getAttribute("val") || "");
  Array.from(doc.getElementsByTagName("*")).filter(e => e.localName === "p").forEach(p => {
    const acc = {
      s: ""
    };
    textOf(p, acc);
    // Признаки «это заголовок»: стиль «Заголовок…» или весь текст абзаца набран жирным
    let heading = false;
    kids(p, "pPr").forEach(pr => kids(pr, "pStyle").forEach(st => {
      if (/head|title|заголов|назван/i.test(st.getAttribute("w:val") || st.getAttribute("val") || "")) heading = true;
    }));
    let anyText = false,
      allBold = true;
    Array.from(p.getElementsByTagName("*")).filter(e => e.localName === "r").forEach(r => {
      if (!(r.textContent || "").trim()) return;
      anyText = true;
      const on = kids(r, "rPr").some(rp => kids(rp, "b").some(flagOn));
      if (!on) allBold = false;
    });
    const m = {
      bold: anyText && allBold,
      heading
    };
    acc.s.split("\n").forEach(l => {
      out.push(l.replace(/ /g, " ").replace(/[ \t]+$/g, ""));
      meta.push(m);
    });
  });
  out.meta = meta;
  out.src = "docx";
  return out;
}
let __pdfJsPromise = null;
function loadPdfJs() {
  if (window.pdfjsLib) return Promise.resolve(window.pdfjsLib);
  if (!__pdfJsPromise) __pdfJsPromise = new Promise((resolve, reject) => {
    const sc = document.createElement("script");
    sc.src = "vendor/pdf.min.js";
    sc.onload = () => {
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = "vendor/pdf.worker.min.js";
      resolve(window.pdfjsLib);
    };
    sc.onerror = () => reject(new Error("Не удалось загрузить модуль чтения PDF"));
    document.head.appendChild(sc);
  });
  return __pdfJsPromise;
}
async function pdfToLines(buf) {
  const lib = await loadPdfJs();
  const pdf = await lib.getDocument({
    data: new Uint8Array(buf)
  }).promise;
  const out = [];
  for (let p = 1; p <= pdf.numPages; p++) {
    const content = await (await pdf.getPage(p)).getTextContent();
    const rows = [];
    content.items.forEach(it => {
      if (typeof it.str !== "string") return;
      const y = it.transform[5],
        x = it.transform[4];
      let row = rows.find(r => Math.abs(r.y - y) < 2.5);
      if (!row) {
        row = {
          y,
          parts: []
        };
        rows.push(row);
      }
      row.parts.push({
        x,
        w: it.width || 0,
        s: it.str
      });
    });
    rows.sort((a, b) => b.y - a.y);
    const gaps = [];
    for (let i = 1; i < rows.length; i++) gaps.push(rows[i - 1].y - rows[i].y);
    const sorted = gaps.slice().sort((a, b) => a - b);
    const typical = sorted.length ? sorted[Math.floor(sorted.length / 2)] : 0;
    rows.forEach((row, i) => {
      if (i > 0 && typical > 0 && rows[i - 1].y - row.y > typical * 1.7) out.push("");
      row.parts.sort((a, b) => a.x - b.x);
      let line = "",
        prevEnd = null;
      row.parts.forEach(pt => {
        if (prevEnd !== null && pt.x - prevEnd > 1 && !/\s$/.test(line) && !/^\s/.test(pt.s)) line += " ";
        line += pt.s;
        prevEnd = pt.x + pt.w;
      });
      out.push(line.replace(/\s+$/g, ""));
    });
    out.push("");
  }
  out.src = "pdf";
  return out;
}
async function fileToLines(file) {
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  const buf = await file.arrayBuffer();
  const head = new Uint8Array(buf.slice(0, 4));
  if (head[0] === 0xD0 && head[1] === 0xCF) throw new Error("Старый формат Word (.doc) не читается — откройте файл в Word и сохраните как «Документ Word (.docx)»");
  if (ext === "docx" || head[0] === 0x50 && head[1] === 0x4B) return docxToLines(buf);
  if (ext === "pdf" || head[0] === 0x25 && head[1] === 0x50) return pdfToLines(buf);
  const text = new TextDecoder("utf-8").decode(buf).replace(/^\uFEFF/, "");
  if (/^\s*<(!doctype|html|\?xml)/i.test(text) || ext === "htm" || ext === "html" || ext === "doc") return htmlToLines(text);
  return text.split(/\r?\n/).map(l => l.replace(/\s+$/g, ""));
}
// ----- Другие источники: электронные книги fb2 / epub и готовые карточки в формате JSON -----
function xmlAll(root, name) {
  return Array.from(root.getElementsByTagName("*")).filter(e => e.localName === name);
}
// fb2: если в книге есть разметка <poem>/<stanza>/<v> — стихи берутся готовыми (почти без ошибок);
// иначе абзацы книги превращаются в строки и режутся обычным способом.
function fb2Parse(text) {
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.getElementsByTagName("parsererror").length) throw new Error("Файл fb2 повреждён или не читается");
  const tx = el => (el.textContent || "").replace(/\s+/g, " ").trim();
  const titleOf = el => {
    const t = Array.from(el.children).find(k => k.localName === "title");
    return t ? xmlAll(t, "p").map(tx).filter(Boolean).join(" ") : "";
  };
  const poems = [];
  xmlAll(doc, "poem").forEach(po => {
    let title = titleOf(po);
    if (!title) {
      let a = po.parentNode;
      while (a && a.nodeType === 1 && !title) {
        title = titleOf(a);
        a = a.parentNode;
      }
    }
    const lines = [];
    Array.from(po.children).filter(k => k.localName === "stanza").forEach(st => {
      const vs = xmlAll(st, "v").map(tx).filter(Boolean);
      if (!vs.length) return;
      if (lines.length) lines.push("");
      vs.forEach(v => lines.push(v));
    });
    if (!lines.length) return;
    const dateEl = Array.from(po.children).find(k => k.localName === "date");
    const t = poemCleanTitle(title);
    poems.push({
      author: "",
      title: t || poemAutoTitle(lines[0]),
      autoTitle: !t,
      classes: [],
      tropes: [],
      note: dateEl ? "Датировка: " + tx(dateEl) : "",
      lines
    });
  });
  const out = [],
    meta = [];
  const push = (l, m) => {
    out.push(l);
    meta.push(m || {});
  };
  const walk = el => {
    const n = el.localName;
    if (n === "binary" || n === "description") return;
    if (n === "empty-line") {
      push("");
      return;
    }
    if (n === "title") {
      xmlAll(el, "p").forEach(p => push(tx(p), {
        bold: true,
        heading: true
      }));
      push("");
      return;
    }
    if (n === "p" || n === "v") {
      const t = tx(el),
        st = xmlAll(el, "strong");
      const bold = !!t && st.length > 0 && st.map(tx).join(" ").replace(/\s+/g, " ").trim() === t;
      push(t, bold ? {
        bold: true
      } : {});
      return;
    }
    Array.from(el.children).forEach(walk);
    if (n === "stanza") push("");
  };
  if (!poems.length) Array.from(doc.documentElement.children).forEach(walk);
  out.meta = meta;
  out.src = "fb2";
  return {
    poems,
    lines: out
  };
}
async function epubToLines(buf) {
  const dec = b => new TextDecoder("utf-8").decode(b);
  const xml = t => new DOMParser().parseFromString(t, "application/xml");
  const container = xml(dec(await readZipEntryBytes(buf, "META-INF/container.xml")));
  const rf = xmlAll(container, "rootfile")[0];
  const opfPath = rf && rf.getAttribute("full-path");
  if (!opfPath) throw new Error("Файл epub не читается (нет описания книги)");
  const opf = xml(dec(await readZipEntryBytes(buf, opfPath)));
  const dir = opfPath.indexOf("/") >= 0 ? opfPath.slice(0, opfPath.lastIndexOf("/") + 1) : "";
  const items = {};
  xmlAll(opf, "item").forEach(it => {
    items[it.getAttribute("id")] = it.getAttribute("href");
  });
  const order = xmlAll(opf, "itemref").map(r => items[r.getAttribute("idref")]).filter(Boolean);
  const out = [],
    meta = [];
  const push = (l, m) => {
    out.push(l);
    meta.push(m || {});
  };
  const BLOCK = /^(p|div|h[1-6]|li|blockquote|section|article|body|pre)$/;
  const leaf = el => {
    const clone = el.cloneNode(true);
    Array.from(clone.querySelectorAll("br")).forEach(br => br.replaceWith("\n"));
    const raw = (clone.textContent || "").replace(/ /g, " ");
    if (!raw.trim()) {
      push("");
      return;
    }
    const parts = raw.split("\n").map(s => s.replace(/\s+/g, " ").trim());
    const heading = /^h[1-6]$/.test(el.localName) || /title|head|zagol|назв|заголов/i.test(el.className || "");
    const boldAll = el.children.length === 1 && /^(b|strong)$/.test(el.children[0].localName) && (el.textContent || "").trim() === (el.children[0].textContent || "").trim();
    if (heading) push("");
    parts.forEach(l => push(l, {
      bold: boldAll,
      heading
    }));
    if (parts.length > 1 || heading) push("");
  };
  const walk = el => {
    const kids = Array.from(el.children);
    if (!kids.some(k => BLOCK.test(k.localName))) {
      leaf(el);
      return;
    }
    kids.forEach(k => {
      if (BLOCK.test(k.localName)) walk(k);
    });
  };
  for (const href of order) {
    let bytes;
    try {
      bytes = await readZipEntryBytes(buf, dir + decodeURIComponent(href.split("#")[0]));
    } catch (e) {
      continue;
    }
    const html = new DOMParser().parseFromString(dec(bytes), "text/html");
    if (html.body) walk(html.body);
    push("");
    push("");
  }
  out.meta = meta;
  out.src = "epub";
  return out;
}
// Готовые карточки (JSON): терпимо относится к названиям ключей, русским и английским. Образец — кнопка «Образец JSON».
function poemGrades(v) {
  const set = new Set();
  String(v == null ? "" : v).replace(/(\d+)\s*[-–—]\s*(\d+)/g, (m, a, b) => {
    for (let i = +a; i <= +b && i <= 11; i++) set.add(String(i));
    return " ";
  }).replace(/\d+/g, m => {
    set.add(String(+m));
    return m;
  });
  return Array.from(set);
}
function poemsFromJson(root) {
  const pick = (o, keys) => {
    if (!o || typeof o !== "object") return undefined;
    for (const k of keys) if (o[k] !== undefined && o[k] !== null && o[k] !== "") return o[k];
    return undefined;
  };
  const str = v => (v == null ? "" : typeof v === "object" ? String(pick(v, ["name", "title", "название", "имя"]) || "") : String(v)).trim();
  const lineText = x => typeof x === "string" ? x : x && typeof x === "object" ? String(pick(x, ["text", "line", "строка", "текст", "t"]) || "") : "";
  let list = null;
  if (Array.isArray(root)) list = root;else if (root && typeof root === "object") {
    const k = ["poems", "стихотворения", "стихи", "cards", "карточки", "items", "works", "произведения", "data"].find(x => Array.isArray(root[x]));
    if (k) list = root[k];else if (pick(root, ["stanzas", "строфы", "text", "текст", "lines", "строки"]) !== undefined) list = [root];else {
      const any = Object.values(root).find(Array.isArray);
      if (any) list = any;
    }
  }
  if (!list) throw new Error("В файле JSON не найден список стихотворений");
  const rootAuthor = root && !Array.isArray(root) ? str(pick(root, ["author", "автор", "poet", "поэт"])) : "";
  const rangeText = v => {
    if (Array.isArray(v)) {
      const nums = v.map(Number).filter(n => !isNaN(n)).sort((a, b) => a - b);
      if (nums.length && nums.length === v.length) {
        const parts = [];
        let s = nums[0],
          p = nums[0];
        for (let i = 1; i <= nums.length; i++) {
          if (nums[i] === p + 1) {
            p = nums[i];
            continue;
          }
          parts.push(s === p ? String(s) : s + "–" + p);
          s = nums[i];
          p = nums[i];
        }
        return parts.join(", ");
      }
      return v.join(", ");
    }
    return String(v);
  };
  const poems = [];
  list.forEach(it => {
    if (!it || typeof it !== "object") return;
    let lines = [];
    const stanzas = pick(it, ["stanzas", "строфы"]);
    if (Array.isArray(stanzas)) {
      stanzas.forEach(st => {
        let ls = [];
        if (typeof st === "string") ls = st.split(/\r?\n/);else if (Array.isArray(st)) ls = st.map(lineText);else if (st && typeof st === "object") {
          const inner = pick(st, ["lines", "строки", "text", "текст"]);
          ls = Array.isArray(inner) ? inner.map(lineText) : String(inner || "").split(/\r?\n/);
        }
        ls = ls.map(l => l.replace(/\s+$/g, "")).filter(l => l.trim());
        if (!ls.length) return;
        if (lines.length) lines.push("");
        lines = lines.concat(ls);
      });
    } else {
      const flat = pick(it, ["lines", "строки"]);
      if (Array.isArray(flat)) lines = flat.map(lineText);else lines = String(pick(it, ["text", "текст"]) || "").split(/\r?\n/);
      lines = lines.map(l => l.replace(/\s+$/g, ""));
      while (lines.length && !lines[0].trim()) lines.shift();
      while (lines.length && !lines[lines.length - 1].trim()) lines.pop();
    }
    if (!lines.length) return;
    const title = poemCleanTitle(str(pick(it, ["title", "название", "заглавие", "name"])));
    const keyRaw = pick(it, ["tropes", "средства", "тропы", "figures", "key", "ключ"]);
    const tropes = [],
      keyLines = [];
    (Array.isArray(keyRaw) ? keyRaw : []).forEach(k => {
      const name = typeof k === "string" ? k : str(pick(k, ["name", "trope", "type", "название", "тип", "троп", "средство"]));
      if (!name.trim()) return;
      const nm = name.trim().charAt(0).toUpperCase() + name.trim().slice(1);
      if (!tropes.includes(nm)) tropes.push(nm);
      if (typeof k === "object") {
        const ln = pick(k, ["lines", "line", "строки", "строка", "stih"]);
        const quote = str(pick(k, ["quote", "text", "цитата", "фрагмент"]));
        const comment = str(pick(k, ["comment", "note", "пояснение", "комментарий", "description"]));
        if (ln !== undefined || quote || comment) keyLines.push("• " + nm + (ln !== undefined ? " (" + (Array.isArray(ln) && ln.length > 1 || /[-–,]/.test(String(ln)) ? "строки " : "строка ") + rangeText(ln) + ")" : "") + (quote ? ": «" + quote + "»" : "") + (comment ? " — " + comment : ""));
      }
    });
    const noteParts = [];
    const date = str(pick(it, ["date", "year", "датировка", "дата", "год"])),
      page = str(pick(it, ["page", "страница", "стр"])),
      own = str(pick(it, ["note", "заметка", "примечание"]));
    if (date) noteParts.push("Датировка: " + date);
    if (page) noteParts.push("Страница: " + page);
    if (own) noteParts.push(own);
    if (keyLines.length) noteParts.push("Ключ:\n" + keyLines.join("\n"));
    poems.push({
      author: str(pick(it, ["author", "автор", "poet", "поэт"])) || rootAuthor,
      title: title || poemAutoTitle(lines.find(l => l.trim())),
      autoTitle: !title,
      classes: poemGrades(pick(it, ["classes", "class", "классы", "класс"])),
      tropes,
      note: noteParts.join("\n"),
      lines
    });
  });
  return poems;
}
function importPoemsFromLines(lines) {
  let poems = parsePoemsFromLines(lines),
    mode = "template";
  if (!poems.length) {
    poems = splitCollectionLines(lines);
    mode = "collection";
  }
  return {
    poems,
    mode,
    dropped: poems.dropped || 0
  };
}
// Читает любой поддерживаемый файл и возвращает { poems, mode }.
async function fileToPoems(file) {
  const ext = (file.name.split(".").pop() || "").toLowerCase();
  const decodeText = async () => {
    const buf = await file.arrayBuffer();
    const probe = new TextDecoder("latin1").decode(buf.slice(0, 300));
    const enc = (probe.match(/encoding=["']([\w-]+)["']/i) || [])[1];
    try {
      return new TextDecoder(enc || "utf-8").decode(buf).replace(/^﻿/, "");
    } catch (e) {
      return new TextDecoder("utf-8").decode(buf).replace(/^﻿/, "");
    }
  };
  if (ext === "json") {
    let obj;
    try {
      obj = JSON.parse(await decodeText());
    } catch (e) {
      throw new Error("Файл JSON повреждён");
    }
    return {
      poems: poemsFromJson(obj),
      mode: "cards"
    };
  }
  if (ext === "epub") return importPoemsFromLines(await epubToLines(await file.arrayBuffer()));
  if (ext === "fb2") {
    const r = fb2Parse(await decodeText());
    return r.poems.length ? {
      poems: r.poems,
      mode: "collection"
    } : importPoemsFromLines(r.lines);
  }
  return importPoemsFromLines(await fileToLines(file));
}
function downloadPoemJsonSample() {
  const sample = {
    "автор": "Ф.И. Тютчев",
    "стихотворения": [{
      "название": "Фонтан",
      "классы": [5, 6],
      "датировка": "1836",
      "страница": 12,
      "строфы": [["Смотри, как облаком живым", "Фонтан сияющий клубится;", "Как пламенеет, как дробится", "Его на солнце влажный дым."], ["Лучом поднявшись к небу, он", "Коснулся высоты заветной —", "И снова пылью огнецветной", "Ниспасть на землю осуждён."]],
      "тропы": [{
        "название": "метафора",
        "строки": [1, 2],
        "цитата": "облаком живым",
        "комментарий": "фонтан сравнивается с облаком без слов «как» и «словно»"
      }, {
        "название": "олицетворение",
        "строки": [8],
        "цитата": "на землю осуждён"
      }, "эпитет"]
    }]
  };
  const blob = new Blob([JSON.stringify(sample, null, 2)], {
    type: "application/json"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Образец_карточек_стихотворений.json";
  a.click();
  URL.revokeObjectURL(url);
}
// ----- Разбор сборника без разметки: автор вводится отдельно, а границы стихотворений программа ищет сама -----
function poemAutoTitle(line) {
  const t = (line || "").trim().replace(/[\s,;:.!?…—–-]+$/g, "");
  return t.length > 60 ? t.slice(0, 57).trim() + "…" : t ? t + "…" : "";
}
function poemCleanTitle(t) {
  let r = (t || "").trim().replace(/^[*•·\s]+|[*•·\s]+$/g, "").replace(/^[«"“„'‘]+|[»"”“'’]+$/g, "").replace(/\.$/, "").trim();
  // ЗАГОЛОВОК ПРОПИСНЫМИ → «Заголовок прописными»
  if (r.length > 3 && /[A-Za-zА-Яа-яЁё]/.test(r) && r === r.toLocaleUpperCase("ru")) {
    r = r.toLocaleLowerCase("ru");
    r = r.charAt(0).toLocaleUpperCase("ru") + r.slice(1);
  }
  return r;
}
// Если в файле между КАЖДЫМ абзацем стоит пустая строка (так бывает в fb2 и Word), то одна пустая строка —
// просто разделитель строк, а граница строф — две подряд. Приводим такой текст к обычному виду.
function normalizeSpacing(lines) {
  const gaps = [];
  let run = 0,
    seen = false;
  lines.forEach(l => {
    if (l.trim() === "") {
      if (seen) run++;
    } else {
      if (seen) gaps.push(run);
      run = 0;
      seen = true;
    }
  });
  const ones = gaps.filter(g => g === 1).length,
    zeros = gaps.filter(g => g === 0).length,
    big = gaps.filter(g => g >= 2).length;
  if (gaps.length < 20 || big === 0 || ones < gaps.length * 0.5 || zeros > gaps.length * 0.15) return lines;
  const meta = lines.meta || null;
  const out = [],
    om = [];
  let i = 0;
  while (i < lines.length) {
    if (lines[i].trim() !== "") {
      out.push(lines[i]);
      om.push(meta ? meta[i] : null);
      i++;
      continue;
    }
    let j = i;
    while (j < lines.length && lines[j].trim() === "") j++;
    const g = j - i;
    if (out.length && j < lines.length) {
      for (let k = 0; k < g - 1; k++) {
        out.push("");
        om.push(null);
      }
    }
    i = j;
  }
  if (meta) out.meta = om;
  out.src = lines.src;
  return out;
}
function splitStanzas(lines) {
  const out = [];
  let cur = null;
  lines.forEach(l => {
    if (l.trim() === "") {
      cur = null;
      return;
    }
    if (!cur) {
      cur = [];
      out.push(cur);
    }
    cur.push(l);
  });
  return out;
}
function joinStanzas(sts) {
  const out = [];
  sts.forEach((s, i) => {
    if (i) out.push("");
    s.forEach(l => out.push(l));
  });
  return out;
}
// Делит весь текст файла на стихотворения. Признаки заголовка: строфа из одной короткой строки, которая
// не заканчивается запятой/точкой и после которой идёт обычная строфа; в Word — ещё и жирный шрифт/стиль «Заголовок».
// Строка «* * *» тоже начинает новое стихотворение (название берётся из первой строки).
function splitCollectionLines(rawLines) {
  rawLines = normalizeSpacing(rawLines);
  const meta = rawLines.meta || null;
  const isPdf = rawLines.src === "pdf";
  const starLine = l => l.replace(/\s+/g, "").length >= 3 && /^[\s*•·_#=~\-–—]+$/.test(l);
  const blocks = [];
  let cur = null,
    blanks = 0;
  rawLines.forEach((raw, i) => {
    const line = raw.replace(/ /g, " ").replace(/\s+$/g, "");
    if (isPdf && /^\s*\d{1,3}\s*$/.test(line)) return; // номера страниц в PDF
    if (line.trim() === "") {
      blanks++;
      cur = null;
      return;
    }
    if (starLine(line)) {
      blocks.push({
        lines: [line],
        blanksBefore: blanks,
        star: true
      });
      cur = null;
      blanks = 0;
      return;
    }
    const m = meta && meta[i];
    const lineBold = !!(m && (m.bold || m.heading));
    if (cur && meta && cur.lastBold !== lineBold) cur = null; // жирная строка (заголовок) — всегда отдельный блок
    if (!cur) {
      cur = {
        lines: [],
        blanksBefore: blanks,
        bold: true,
        heading: false
      };
      blocks.push(cur);
      blanks = 0;
    }
    cur.lastBold = lineBold;
    cur.lines.push(line);
    if (!m || !m.bold) cur.bold = false;
    if (m && m.heading) cur.heading = true;
  });
  const isStar = b => !!b.star;
  const isDateLike = t => /^[\s(\[]*(?:\d{1,2}[.\s]\d{1,2}[.\s])?\d{3,4}(?:\s*[-–—]\s*\d{2,4})?\s*(?:г\.?|гг\.?|год)?[\s)\].]*$/i.test(t);
  const oneLiners = blocks.filter(b => b.lines.length === 1 && !isStar(b));
  const boldTitles = oneLiners.filter(b => b.bold || b.heading).length;
  const boldMulti = blocks.filter(b => b.lines.length > 1 && b.bold).length;
  const useBold = !!meta && boldTitles >= 2 && boldMulti <= boldTitles * 0.2;
  const isTitle = (b, i) => {
    if (b.lines.length !== 1 || isStar(b)) return false;
    const next = blocks[i + 1];
    if (!next || isStar(next)) return false;
    if (useBold) return !!(b.bold || b.heading);
    if (b.heading) return true;
    const t = b.lines[0].trim();
    if (t.length > 80 || isDateLike(t)) return false;
    if (next.lines.length < 2) return false;
    if (/^(\d+|[IVXLC]+)\.?$/.test(t)) return true;
    if (/[,;:]$/.test(t) || /\.$/.test(t)) return false;
    if (/^[a-zа-яё]/.test(t)) return false;
    return true;
  };
  const poems = [];
  let p = null;
  const startP = (title, auto) => {
    p = {
      author: "",
      title,
      autoTitle: !!auto,
      classes: [],
      tropes: [],
      lines: []
    };
    poems.push(p);
  };
  blocks.forEach((b, i) => {
    if (isStar(b)) {
      p = null;
      return;
    }
    if (isTitle(b, i)) {
      startP(poemCleanTitle(b.lines[0]), false);
      return;
    }
    if (!p) startP("", true);
    if (p.lines.length) p.lines.push("");
    b.lines.forEach(l => p.lines.push(l));
  });
  poems.forEach(x => {
    if (!x.title) {
      x.title = poemAutoTitle(x.lines[0]);
      x.autoTitle = true;
    }
  });
  // обрывки из одной строки (оглавление, шапка, подписи) стихотворениями не считаем
  const all = poems.filter(x => x.lines.length);
  all.forEach(x => {
    if (/^(\d+|[IVXLC]+)\.?$/.test(x.title)) {
      x.title = poemAutoTitle(x.lines.find(l => l.trim()));
      x.autoTitle = true;
    }
  });
  const norm = t => (t || "").toLowerCase().replace(/[^a-zа-яё0-9]+/g, " ").trim();
  const titles = new Set(all.filter(x => !x.autoTitle).map(x => norm(x.title)));
  const tocLike = x => {
    const ls = x.lines.filter(l => l.trim());
    if (ls.length > 4) return false;
    const hit = ls.filter(l => {
      const t = l.trim();
      return t.length <= 60 && t === t.toLocaleUpperCase("ru") && /[А-ЯЁA-Z]/.test(t) || titles.has(norm(t)) || /^(стихи|стихотворения|содержание|оглавление|аннотация|annotation|предисловие)$/i.test(t);
    }).length;
    return hit >= Math.ceil(ls.length * 2 / 3);
  };
  const res = all.filter(x => x.lines.filter(l => l.trim()).length > 1 && !tocLike(x));
  res.dropped = all.length - res.length;
  return res;
}
function parsePoemsFromLines(rawLines) {
  const L = {
    author: /^(?:автор|поэт)\s*:\s*(.*)$/i,
    title: /^(?:название|заглавие|стихотворение)\s*:\s*(.*)$/i,
    cls: /^(?:классы?|кл\.?)\s*:\s*(.*)$/i,
    tropes: /^(?:тропы?|искомые\s+тропы?|искать)\s*:\s*(.*)$/i,
    text: /^текст\s*:\s*(.*)$/i,
    sep: /^\s*[*#=_~\-–—]{3,}\s*$/
  };
  const poems = [];
  let cur = null;
  const finish = () => {
    if (!cur) return;
    while (cur.lines.length && cur.lines[0].trim() === "") cur.lines.shift();
    while (cur.lines.length && cur.lines[cur.lines.length - 1].trim() === "") cur.lines.pop();
    if ((cur.title || cur.author) && cur.lines.length) poems.push(cur);
    cur = null;
  };
  const start = () => {
    finish();
    cur = {
      author: "",
      title: "",
      classes: [],
      tropes: [],
      lines: [],
      inText: false
    };
  };
  const unq = t => t.trim().replace(/^[«"“„'‘]+|[»"”“'’]+$/g, "").trim();
  const grades = v => {
    const set = new Set();
    v.replace(/(\d+)\s*[-–—]\s*(\d+)/g, (m, a, b) => {
      for (let i = +a; i <= +b && i <= 11; i++) set.add(String(i));
      return " ";
    }).replace(/\d+/g, m => {
      set.add(String(+m));
      return m;
    });
    return Array.from(set);
  };
  rawLines.forEach(raw => {
    const line = raw.replace(/\u00a0/g, " ").replace(/\s+$/g, "");
    const t = line.trim();
    if (L.sep.test(t)) {
      finish();
      return;
    }
    let m;
    if ((m = t.match(L.author)) || (m = t.match(L.title))) {
      const isAuthor = !!t.match(L.author);
      if (!cur || cur.inText || (isAuthor ? cur.author : cur.title)) start();
      if (isAuthor) cur.author = m[1].trim();else cur.title = unq(m[1]);
      return;
    }
    if (m = t.match(L.cls)) {
      if (!cur || cur.inText) start();
      cur.classes = grades(m[1]);
      return;
    }
    if (m = t.match(L.tropes)) {
      if (!cur || cur.inText) start();
      cur.tropes = m[1].split(/[;,]/).map(x => x.trim()).filter(Boolean).map(x => x.charAt(0).toUpperCase() + x.slice(1));
      return;
    }
    if (!cur) return;
    if (m = t.match(L.text)) {
      cur.inText = true;
      if (m[1].trim()) cur.lines.push(m[1].trim());
      return;
    }
    if (!cur.inText) {
      if (t === "") return;
      cur.inText = true;
    }
    cur.lines.push(line);
  });
  finish();
  return poems;
}
function poemHtml(lines) {
  return lines.map(l => l.trim() === "" ? "<div><br></div>" : "<div>" + escapeHtml(l.trim()) + "</div>").join("");
}
function classTagIdForGrade(d, n) {
  const found = d.tags.classTag.find(t => {
    const nm = (t.name || "").trim().toLowerCase();
    return nm === n || new RegExp("^" + n + "\\s*(класс|кл\\.?)$").test(nm);
  });
  if (found) return found.id;
  const id = uid();
  d.tags.classTag.push({
    id,
    name: n,
    color: SWATCHES[Math.floor(Math.random() * SWATCHES.length)]
  });
  return id;
}
// Создаёт карточки «Прогулки по тропам» из разобранных стихотворений (вызывать внутри update); дубликаты
// (тот же автор, название и текст) пропускаются. Возвращает { created, skipped }.
function createPoemExamples(d, poems) {
  let created = 0,
    skipped = 0;
  poems.forEach(p => {
    const authorId = p.author ? ensurePersonality(d, p.author) : null;
    const text = poemHtml(p.lines);
    if (d.examples.some(e => e.typeTagId === TROPE_WALK_TYPE_ID && (e.authorId || null) === authorId && (e.workTitle || "") === p.title && e.text === text)) {
      skipped++;
      return;
    }
    const tagIds = [...p.classes.map(n => classTagIdForGrade(d, n)), ...p.tropes.map(n => ensureTag(d, "tropeTag", n))].filter(Boolean);
    d.examples.push(migrateExample({
      typeTagId: TROPE_WALK_TYPE_ID,
      text,
      tagIds,
      authorId,
      workTitle: p.title,
      note: p.note || ""
    }));
    created++;
  });
  return {
    created,
    skipped
  };
}
// Шаблон-образец (.docx) для импорта: инструкция + два стихотворения.
function downloadPoemTemplateDocx() {
  const esc = t => escapeHtml(t);
  const para = (t, bold) => t === "" ? "<w:p/>" : `<w:p><w:r>${bold ? "<w:rPr><w:b/></w:rPr>" : ""}<w:t xml:space="preserve">${esc(t)}</w:t></w:r></w:p>`;
  const rows = [["ШАБЛОН ДЛЯ ИМПОРТА СТИХОТВОРЕНИЙ В «ПРОГУЛКУ ПО ТРОПАМ»", true], ["Как заполнять: перед каждым стихотворением пишутся строки с метками (метка, двоеточие, значение), затем сам текст — каждая строка стихотворения с новой строки, строфы разделяются пустой строкой.", false], ["Метки: «Автор», «Название» — обязательны; «Класс» (можно несколько: 5, 6 или 5-7) и «Тропы» (через запятую) — по желанию.", false], ["Новое стихотворение начинается с новой метки «Автор:» либо со строки из трёх звёздочек. Всё, что написано выше первой метки (в том числе эта инструкция), программа пропускает.", false], ["Если автора ещё нет в «Персоналиях», карточка создастся сама. Ниже — два образца, их можно удалить.", false], ["", false], ["***", false], ["Автор: Ф.И. Тютчев", false], ["Название: Фонтан", false], ["Класс: 5, 6", false], ["Тропы: метафора, олицетворение", false], ["Текст:", false], ["Смотри, как облаком живым", false], ["Фонтан сияющий клубится;", false], ["Как пламенеет, как дробится", false], ["Его на солнце влажный дым.", false], ["", false], ["Лучом поднявшись к небу, он", false], ["Коснулся высоты заветной —", false], ["И снова пылью огнецветной", false], ["Ниспасть на землю осуждён.", false], ["", false], ["***", false], ["Автор: А.С. Пушкин", false], ["Название: Зимнее утро", false], ["Класс: 5", false], ["Тропы: эпитет", false], ["Текст:", false], ["Мороз и солнце; день чудесный!", false], ["Ещё ты дремлешь, друг прелестный —", false], ["Пора, красавица, проснись:", false], ["Открой сомкнуты негой взоры", false], ["Навстречу северной Авроры,", false], ["Звездою севера явись!", false]];
  const body = rows.map(r => para(r[0], r[1])).join("");
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const files = [{
    name: "[Content_Types].xml",
    data: strToBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>')
  }, {
    name: "_rels/.rels",
    data: strToBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
  }, {
    name: "word/document.xml",
    data: strToBytes('<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:document xmlns:w="' + W + '"><w:body>' + body + "</w:body></w:document>")
  }];
  const blob = new Blob([makeZip(files)], {
    type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "Шаблон_стихотворения.docx";
  a.click();
  URL.revokeObjectURL(url);
}
function PoemImportModal({
  data,
  update,
  onClose
}) {
  const [author, setAuthor] = useState("");
  const [src, setSrc] = useState(null);
  const [poems, setPoems] = useState([]);
  const [busy, setBusy] = useState(false);
  const [pasted, setPasted] = useState("");
  const fileRef = useRef(null);
  const authorOf = p => (p.author || author || "").trim();
  const isDup = p => {
    const a = authorOf(p);
    const pers = a ? data.personalities.find(x => (x.title || "").trim().toLowerCase() === a.toLowerCase()) : null;
    const text = poemHtml(p.lines);
    return data.examples.some(e => e.typeTagId === TROPE_WALK_TYPE_ID && (e.authorId || null) === (pers ? pers.id : null) && (e.workTitle || "") === (p.title || "").trim() && e.text === text);
  };
  const onFile = async e => {
    const f = e.target.files && e.target.files[0];
    e.target.value = "";
    if (!f) return;
    setBusy(true);
    try {
      const r = await fileToPoems(f);
      showResult(r.poems, r.mode, f.name, r.dropped);
    } catch (err) {
      showToast("Не удалось прочитать файл: " + (err && err.message ? err.message : err));
    }
    setBusy(false);
  };
  const showResult = (ps, mode, name, dropped) => {
    setSrc({
      fileName: name,
      mode,
      dropped: dropped || 0
    });
    // длинные «стихи» со строками-абзацами (аннотация, предисловие) по умолчанию не отмечаем
    const prose = p => {
      const ls = p.lines.filter(l => l.trim());
      return ls.reduce((s, l) => s + l.length, 0) / Math.max(1, ls.length) > 80;
    };
    setPoems(ps.map(p => ({
      ...p,
      checked: mode !== "collection" || !prose(p),
      open: false
    })));
  };
  const parsePasted = () => {
    const t = pasted.trim();
    if (!t) return;
    try {
      if (/^[\[{]/.test(t)) {
        let obj = null;
        try {
          obj = JSON.parse(t);
        } catch (e) {/* не JSON — обычный текст */}
        if (obj) {
          showResult(poemsFromJson(obj), "cards", "вставленный текст");
          return;
        }
      }
      const r = importPoemsFromLines(t.split(/\r?\n/));
      showResult(r.poems, r.mode, "вставленный текст", r.dropped);
    } catch (err) {
      showToast("Не удалось разобрать текст: " + (err && err.message ? err.message : err));
    }
  };
  const setP = (i, patch) => setPoems(ps => ps.map((p, j) => j === i ? {
    ...p,
    ...patch
  } : p));
  const mergePrev = i => setPoems(ps => {
    if (i < 1) return ps;
    const prev = ps[i - 1],
      cur = ps[i];
    const lines = [...prev.lines, "", ...(cur.autoTitle || !cur.title ? [] : [cur.title, ""]), ...cur.lines];
    return [...ps.slice(0, i - 1), {
      ...prev,
      lines
    }, ...ps.slice(i + 1)];
  });
  const splitAt = (i, k) => setPoems(ps => {
    const p = ps[i],
      sts = splitStanzas(p.lines);
    if (k < 1 || k >= sts.length) return ps;
    const left = {
      ...p,
      lines: joinStanzas(sts.slice(0, k))
    };
    let rest = sts.slice(k),
      title,
      auto;
    if (rest[0].length === 1 && rest.length > 1) {
      title = poemCleanTitle(rest[0][0]);
      auto = false;
      rest = rest.slice(1);
    } else {
      title = poemAutoTitle(rest[0][0]);
      auto = true;
    }
    const right = {
      ...p,
      title,
      autoTitle: auto,
      lines: joinStanzas(rest),
      checked: true,
      open: false
    };
    return [...ps.slice(0, i), left, right, ...ps.slice(i + 1)];
  });
  const chosen = poems.filter(p => p.checked && !isDup(p));
  const needAuthor = !author.trim() && chosen.some(p => !p.author);
  const run = () => {
    const sel = chosen.map(p => ({
      author: authorOf(p),
      title: (p.title || "").trim(),
      classes: p.classes || [],
      tropes: p.tropes || [],
      lines: p.lines
    }));
    update(d => {
      createPoemExamples(d, sel);
      return d;
    });
    showToast("Создано карточек: " + sel.length + ". Класс и тропы выставьте в самих карточках.");
    onClose();
  };
  const authorInput = /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("label", {
    style: {
      fontWeight: 700,
      display: "block",
      marginBottom: 4
    }
  }, "\u0410\u0432\u0442\u043E\u0440 (\u0432\u0441\u0435 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u044F \u0438\u0437 \u0444\u0430\u0439\u043B\u0430 \u043F\u043E\u043B\u0443\u0447\u0430\u0442 \u0435\u0433\u043E)"), /*#__PURE__*/React.createElement("input", {
    className: "input",
    list: "poem-import-authors",
    value: author,
    onChange: e => setAuthor(e.target.value),
    placeholder: "\u041D\u0430\u043F\u0440\u0438\u043C\u0435\u0440: \u0424.\u0418. \u0422\u044E\u0442\u0447\u0435\u0432",
    style: {
      width: "100%"
    },
    autoFocus: true
  }), /*#__PURE__*/React.createElement("datalist", {
    id: "poem-import-authors"
  }, data.personalities.map(p => /*#__PURE__*/React.createElement("option", {
    key: p.id,
    value: p.title
  }))), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 3
    }
  }, "\u0415\u0441\u043B\u0438 \u0442\u0430\u043A\u043E\u0433\u043E \u0430\u0432\u0442\u043E\u0440\u0430 \u0435\u0449\u0451 \u043D\u0435\u0442 \u0432 \xAB\u041F\u0435\u0440\u0441\u043E\u043D\u0430\u043B\u0438\u044F\u0445\xBB, \u043E\u043D \u0441\u043E\u0437\u0434\u0430\u0441\u0442\u0441\u044F \u0441\u0430\u043C."));
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      width: "min(780px, 94vw)"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, "\u0418\u043C\u043F\u043E\u0440\u0442 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0439"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("input", {
    ref: fileRef,
    type: "file",
    accept: ".docx,.pdf,.txt,.html,.htm,.fb2,.epub,.json",
    style: {
      display: "none"
    },
    onChange: onFile
  }), authorInput, !src ? /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: {
      marginTop: 0
    }
  }, "\u0412\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u0444\u0430\u0439\u043B \u0441\u043E \u0441\u0431\u043E\u0440\u043D\u0438\u043A\u043E\u043C \u0438\u043B\u0438 \u0432\u0441\u0442\u0430\u0432\u044C\u0442\u0435 \u0442\u0435\u043A\u0441\u0442. \u041F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u0430\u043C\u0430 \u0440\u0430\u0437\u0434\u0435\u043B\u0438\u0442 \u0435\u0433\u043E \u043D\u0430 \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0435 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u044F \u0438 \u043F\u043E\u043A\u0430\u0436\u0435\u0442 \u0441\u043F\u0438\u0441\u043E\u043A \u2014 \u0442\u0430\u043C \u043C\u043E\u0436\u043D\u043E \u043F\u043E\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u043D\u0430\u0437\u0432\u0430\u043D\u0438\u044F, \u0441\u043A\u043B\u0435\u0438\u0442\u044C \u0438\u043B\u0438 \u0440\u0430\u0437\u0434\u0435\u043B\u0438\u0442\u044C \u0441\u0442\u0438\u0445\u0438. \u0422\u043E\u043B\u044C\u043A\u043E \u043F\u043E\u0441\u043B\u0435 \u044D\u0442\u043E\u0433\u043E \u0441\u043E\u0437\u0434\u0430\u044E\u0442\u0441\u044F \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438. ", /*#__PURE__*/React.createElement("b", null, "\u041A\u043B\u0430\u0441\u0441 \u0438 \u0442\u0440\u043E\u043F\u044B"), " \u043C\u043E\u0436\u043D\u043E \u0432\u044B\u0441\u0442\u0430\u0432\u0438\u0442\u044C \u043F\u043E\u0442\u043E\u043C, \u0432 \u0441\u0430\u043C\u0438\u0445 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0430\u0445."), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    disabled: busy,
    onClick: () => fileRef.current && fileRef.current.click()
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 13
  }), " ", busy ? "Читаю файл…" : "Выбрать файл…"), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "Word (.docx), \u044D\u043B\u0435\u043A\u0442\u0440\u043E\u043D\u043D\u044B\u0435 \u043A\u043D\u0438\u0433\u0438 (.fb2, .epub), PDF, .txt, \u0430 \u0442\u0430\u043A\u0436\u0435 \u0433\u043E\u0442\u043E\u0432\u044B\u0435 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 (.json)")), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 12,
      fontWeight: 700
    }
  }, "\u2026\u0438\u043B\u0438 \u0432\u0441\u0442\u0430\u0432\u044C\u0442\u0435 \u0442\u0435\u043A\u0441\u0442 \u0441\u044E\u0434\u0430"), /*#__PURE__*/React.createElement("textarea", {
    className: "input",
    value: pasted,
    onChange: e => setPasted(e.target.value),
    placeholder: "\u0412\u0441\u0442\u0430\u0432\u044C\u0442\u0435 \u0442\u0435\u043A\u0441\u0442 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0439 (Ctrl+V). \u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u2014 \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u043E\u0439 \u043A\u043E\u0440\u043E\u0442\u043A\u043E\u0439 \u0441\u0442\u0440\u043E\u043A\u043E\u0439 \u043D\u0430\u0434 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0435\u043C, \u0441\u0442\u0440\u043E\u0444\u044B \u2014 \u0447\u0435\u0440\u0435\u0437 \u043F\u0443\u0441\u0442\u0443\u044E \u0441\u0442\u0440\u043E\u043A\u0443.",
    style: {
      width: "100%",
      minHeight: 110,
      marginTop: 4,
      fontFamily: "inherit"
    }
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: !pasted.trim(),
    onClick: parsePasted
  }, "\u0420\u0430\u0437\u043E\u0431\u0440\u0430\u0442\u044C \u0432\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043D\u044B\u0439 \u0442\u0435\u043A\u0441\u0442")), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 12
    }
  }, "\u041B\u0443\u0447\u0448\u0435 \u0432\u0441\u0435\u0433\u043E \u043D\u0430\u0437\u0432\u0430\u043D\u0438\u044F \u043D\u0430\u0445\u043E\u0434\u044F\u0442\u0441\u044F \u0432 fb2/epub \u0438 \u0432 Word, \u0433\u0434\u0435 \u043E\u043D\u0438 \u043D\u0430\u0431\u0440\u0430\u043D\u044B \u0436\u0438\u0440\u043D\u044B\u043C \u0448\u0440\u0438\u0444\u0442\u043E\u043C \u0438\u043B\u0438 \u0441\u0442\u0438\u043B\u0435\u043C \xAB\u0417\u0430\u0433\u043E\u043B\u043E\u0432\u043E\u043A\xBB; \u0445\u0443\u0436\u0435 \u0432\u0441\u0435\u0433\u043E \u2014 \u0432 PDF (\u0442\u0430\u043C \u043D\u0435\u0442 \u043F\u043E\u043D\u044F\u0442\u0438\u0439 \xAB\u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0435\xBB \u0438 \xAB\u0441\u0442\u0440\u043E\u0444\u0430\xBB, \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0443\u0433\u0430\u0434\u044B\u0432\u0430\u0435\u0442). \u0415\u0441\u043B\u0438 \u0432 \u0444\u0430\u0439\u043B\u0435 \u0435\u0441\u0442\u044C \u043C\u0435\u0442\u043A\u0438 \xAB\u0410\u0432\u0442\u043E\u0440:\xBB, \xAB\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435:\xBB, \xAB\u041A\u043B\u0430\u0441\u0441:\xBB, \xAB\u0422\u0440\u043E\u043F\u044B:\xBB, \u043E\u043D\u0438 \u0442\u043E\u0436\u0435 \u043F\u043E\u043D\u0438\u043C\u0430\u044E\u0442\u0441\u044F (", /*#__PURE__*/React.createElement("a", {
    href: "#",
    onClick: e => {
      e.preventDefault();
      downloadPoemTemplateDocx();
    }
  }, "\u043E\u0431\u0440\u0430\u0437\u0435\u0446 \u0441 \u043C\u0435\u0442\u043A\u0430\u043C\u0438"), "). \u0413\u043E\u0442\u043E\u0432\u044B\u0435 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 \u0441 \xAB\u043A\u043B\u044E\u0447\u043E\u043C\xBB (\u043A\u0430\u043A\u0438\u0435 \u0442\u0440\u043E\u043F\u044B \u0432 \u043A\u0430\u043A\u0438\u0445 \u0441\u0442\u0440\u043E\u043A\u0430\u0445) \u0437\u0430\u0433\u0440\u0443\u0436\u0430\u044E\u0442\u0441\u044F \u0438\u0437 JSON (", /*#__PURE__*/React.createElement("a", {
    href: "#",
    onClick: e => {
      e.preventDefault();
      downloadPoemJsonSample();
    }
  }, "\u043E\u0431\u0440\u0430\u0437\u0435\u0446 JSON"), ") \u2014 \u043F\u0440\u0438 \u044D\u0442\u043E\u043C \u0442\u0440\u043E\u043F\u044B \u0438 \u043A\u043B\u0430\u0441\u0441 \u043F\u0440\u043E\u0441\u0442\u0430\u0432\u044F\u0442\u0441\u044F \u0441\u0430\u043C\u0438.")) : poems.length === 0 ? /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", null, "\u0412 \u0444\u0430\u0439\u043B\u0435 \xAB", src.fileName, "\xBB \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E \u0442\u0435\u043A\u0441\u0442\u0430 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0439."), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => {
      setSrc(null);
      setPoems([]);
    }
  }, "\u0412\u044B\u0431\u0440\u0430\u0442\u044C \u0434\u0440\u0443\u0433\u043E\u0439 \u0444\u0430\u0439\u043B")) : /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 6,
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("b", null, "\u041D\u0430\u0439\u0434\u0435\u043D\u043E \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0439: ", poems.length), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\xB7 \u0444\u0430\u0439\u043B: ", src.fileName), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setPoems(ps => ps.map(p => ({
      ...p,
      checked: true
    })))
  }, "\u0412\u044B\u0431\u0440\u0430\u0442\u044C \u0432\u0441\u0435"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setPoems(ps => ps.map(p => ({
      ...p,
      checked: false
    })))
  }, "\u0421\u043D\u044F\u0442\u044C \u0432\u0441\u0435"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => fileRef.current && fileRef.current.click()
  }, "\u0414\u0440\u0443\u0433\u043E\u0439 \u0444\u0430\u0439\u043B")), src.dropped > 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 6
    }
  }, "\u041D\u0435 \u0432\u043A\u043B\u044E\u0447\u0435\u043D\u043E \u043A\u043E\u0440\u043E\u0442\u043A\u0438\u0445 \u043E\u0431\u0440\u044B\u0432\u043A\u043E\u0432 \u0438\u0437 \u043E\u0434\u043D\u043E\u0439 \u0441\u0442\u0440\u043E\u043A\u0438 (\u043E\u0433\u043B\u0430\u0432\u043B\u0435\u043D\u0438\u0435, \u043F\u043E\u0434\u043F\u0438\u0441\u0438): ", src.dropped, "."), src.mode === "collection" && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 6
    }
  }, "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u044F \u043D\u0430\u0439\u0434\u0435\u043D\u044B \u0430\u0432\u0442\u043E\u043C\u0430\u0442\u0438\u0447\u0435\u0441\u043A\u0438. \u0415\u0441\u043B\u0438 \u043D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u043E\u043F\u0440\u0435\u0434\u0435\u043B\u0435\u043D\u043E \u043D\u0435\u0432\u0435\u0440\u043D\u043E \u2014 \u0438\u0441\u043F\u0440\u0430\u0432\u044C\u0442\u0435 \u0435\u0433\u043E \u0432 \u043F\u043E\u043B\u0435; \u0435\u0441\u043B\u0438 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0435 \u0440\u0430\u0437\u0434\u0440\u043E\u0431\u043B\u0435\u043D\u043E \u2014 \xAB\u0421\u043A\u043B\u0435\u0438\u0442\u044C \u0441 \u043F\u0440\u0435\u0434\u044B\u0434\u0443\u0449\u0438\u043C\xBB; \u0435\u0441\u043B\u0438 \u0434\u0432\u0430 \u0441\u0442\u0438\u0445\u0430 \u0441\u043B\u0438\u043F\u043B\u0438\u0441\u044C \u2014 \xAB\u0421\u0442\u0440\u043E\u0444\u044B\xBB \u2192 \xAB\u041D\u043E\u0432\u043E\u0435 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0435 \u043E\u0442\u0441\u044E\u0434\u0430\xBB."), /*#__PURE__*/React.createElement("div", {
    style: {
      maxHeight: "50vh",
      overflow: "auto",
      border: "1px solid var(--line)",
      borderRadius: 8
    }
  }, poems.map((p, i) => {
    const dup = isDup(p);
    const sts = p.open ? splitStanzas(p.lines) : null;
    return /*#__PURE__*/React.createElement("div", {
      key: i,
      style: {
        padding: "7px 10px",
        borderBottom: "1px solid var(--line)",
        opacity: dup ? 0.55 : 1
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row",
      style: {
        gap: 8,
        alignItems: "center"
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: !!p.checked,
      onChange: e => setP(i, {
        checked: e.target.checked
      })
    }), /*#__PURE__*/React.createElement("span", {
      className: "hint",
      style: {
        width: 24,
        textAlign: "right"
      }
    }, i + 1, "."), p.author && /*#__PURE__*/React.createElement("b", null, p.author, ","), /*#__PURE__*/React.createElement("input", {
      className: "input",
      value: p.title,
      onChange: e => setP(i, {
        title: e.target.value,
        autoTitle: false
      }),
      placeholder: "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u044F",
      style: {
        flex: 1,
        fontWeight: 700
      }
    })), /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        gap: 8,
        marginLeft: 56,
        marginTop: 3,
        alignItems: "center"
      }
    }, /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, p.lines.filter(l => l.trim()).length, " \u0441\u0442\u0440.", p.classes && p.classes.length ? " · класс: " + p.classes.join(", ") : "", p.tropes && p.tropes.length ? " · тропы: " + p.tropes.join(", ") : "", p.note ? " · есть заметка/ключ" : "", dup ? " · такая карточка уже есть" : ""), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => setP(i, {
        open: !p.open
      })
    }, p.open ? "Скрыть строфы" : "Строфы"), i > 0 && /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => mergePrev(i),
      title: "\u041F\u0440\u0438\u0441\u043E\u0435\u0434\u0438\u043D\u0438\u0442\u044C \u044D\u0442\u043E \u043A \u043F\u0440\u0435\u0434\u044B\u0434\u0443\u0449\u0435\u043C\u0443 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u044E"
    }, "\u0421\u043A\u043B\u0435\u0438\u0442\u044C \u0441 \u043F\u0440\u0435\u0434\u044B\u0434\u0443\u0449\u0438\u043C")), !p.open && /*#__PURE__*/React.createElement("div", {
      className: "hint",
      style: {
        marginLeft: 56,
        opacity: 0.8
      }
    }, p.lines.filter(l => l.trim()).slice(0, 2).join(" / "), "\u2026"), sts && /*#__PURE__*/React.createElement("div", {
      style: {
        marginLeft: 56,
        marginTop: 4
      }
    }, sts.map((st, k) => /*#__PURE__*/React.createElement("div", {
      key: k
    }, k > 0 && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => splitAt(i, k),
      title: "\u0412\u0441\u0451, \u0447\u0442\u043E \u043D\u0438\u0436\u0435, \u0441\u0442\u0430\u043D\u0435\u0442 \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u043C \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0435\u043C"
    }, "\u2702 \u041D\u043E\u0432\u043E\u0435 \u0441\u0442\u0438\u0445\u043E\u0442\u0432\u043E\u0440\u0435\u043D\u0438\u0435 \u043E\u0442\u0441\u044E\u0434\u0430", st.length === 1 && k < sts.length - 1 ? " (эта строка — название)" : "")), /*#__PURE__*/React.createElement("div", {
      style: {
        whiteSpace: "pre-wrap",
        fontSize: 13,
        margin: "2px 0 6px"
      }
    }, st.join("\n"))))));
  })))), src && poems.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, needAuthor && /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      color: "var(--danger, #b00020)"
    }
  }, "\u0412\u043F\u0438\u0448\u0438\u0442\u0435 \u0430\u0432\u0442\u043E\u0440\u0430 \u0432\u0432\u0435\u0440\u0445\u0443"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: chosen.length === 0 || needAuthor,
    onClick: run
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u0421\u043E\u0437\u0434\u0430\u0442\u044C \u043A\u0430\u0440\u0442\u043E\u0447\u0435\u043A: ", chosen.length))));
}

// Крупный просмотр «Прогулки по тропам» для вывода на экран/проектор: масштаб текста (кнопки A−/A+,
// клавиши + − 0, Ctrl+колесо), режим «Во весь экран» (F), экспорт Word/PDF/PowerPoint. Esc — закрыть.
function TropeWalkViewer({
  data,
  ex,
  onClose
}) {
  const readZoom = () => {
    try {
      const v = +localStorage.getItem("trope-zoom");
      return v >= 50 && v <= 600 ? v : 150;
    } catch (e) {
      return 150;
    }
  };
  const [zoom, setZoom] = useState(readZoom);
  const [full, setFull] = useState(false);
  const rootRef = useRef(null);
  const author = ex.authorId ? data.personalities.find(p => p.id === ex.authorId) : null;
  const classTags = nsTagIds(ex.tagIds, "classTag", data).map(id => data.tags.classTag.find(t => t.id === id)).filter(Boolean);
  const tropeTags = nsTagIds(ex.tagIds, "tropeTag", data).map(id => data.tags.tropeTag.find(t => t.id === id)).filter(Boolean);
  const heading = (author ? author.title : "Автор не указан") + (ex.workTitle ? `, «${ex.workTitle}»` : "");
  const setZ = fn => setZoom(z => {
    const v = Math.max(50, Math.min(600, Math.round(typeof fn === "function" ? fn(z) : fn)));
    try {
      localStorage.setItem("trope-zoom", String(v));
    } catch (e) {/* без сохранения */}
    return v;
  });
  const toggleFull = () => {
    if (document.fullscreenElement) document.exitFullscreen && document.exitFullscreen();else if (rootRef.current && rootRef.current.requestFullscreen) rootRef.current.requestFullscreen().catch(() => {});
  };
  useEffect(() => {
    const onFs = () => setFull(!!document.fullscreenElement);
    document.addEventListener("fullscreenchange", onFs);
    const onKey = e => {
      if (e.target && /^(INPUT|TEXTAREA|SELECT)$/.test(e.target.tagName)) return;
      if (e.key === "+" || e.key === "=") {
        e.preventDefault();
        setZ(z => z + 10);
      } else if (e.key === "-" || e.key === "_") {
        e.preventDefault();
        setZ(z => z - 10);
      } else if (e.key === "0") {
        e.preventDefault();
        setZ(150);
      } else if (e.key === "f" || e.key === "F" || e.key === "а" || e.key === "А") {
        e.preventDefault();
        toggleFull();
      } else if (e.key === "Escape" && !document.fullscreenElement) onClose();
    };
    window.addEventListener("keydown", onKey);
    return () => {
      document.removeEventListener("fullscreenchange", onFs);
      window.removeEventListener("keydown", onKey);
      if (document.fullscreenElement && document.exitFullscreen) document.exitFullscreen();
    };
  }, []);
  const base = 16 * (zoom / 100);
  return /*#__PURE__*/React.createElement("div", {
    ref: rootRef,
    style: {
      position: "fixed",
      inset: 0,
      zIndex: 1000,
      background: "#fff",
      color: "#000",
      display: "flex",
      flexDirection: "column"
    },
    onWheel: e => {
      if (e.ctrlKey) {
        e.preventDefault();
        setZ(z => z + (e.deltaY < 0 ? 10 : -10));
      }
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      padding: "8px 14px",
      borderBottom: "1px solid var(--line)",
      background: "#F4F6F9",
      flex: "none"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setZ(z => z - 10),
    title: "\u0423\u043C\u0435\u043D\u044C\u0448\u0438\u0442\u044C (\u2212)"
  }, "A\u2212"), /*#__PURE__*/React.createElement("input", {
    type: "range",
    min: "50",
    max: "600",
    step: "10",
    value: zoom,
    onChange: e => setZ(+e.target.value),
    style: {
      width: 160
    },
    title: "\u041C\u0430\u0441\u0448\u0442\u0430\u0431 \u0442\u0435\u043A\u0441\u0442\u0430"
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setZ(z => z + 10),
    title: "\u0423\u0432\u0435\u043B\u0438\u0447\u0438\u0442\u044C (+)"
  }, "A+"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontWeight: 700,
      minWidth: 48,
      textAlign: "center"
    }
  }, zoom, "%"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setZ(150),
    title: "\u0412\u0435\u0440\u043D\u0443\u0442\u044C \u0438\u0441\u0445\u043E\u0434\u043D\u044B\u0439 \u043C\u0430\u0441\u0448\u0442\u0430\u0431 (0)"
  }, "150%"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: toggleFull,
    title: "\u0412\u043E \u0432\u0435\u0441\u044C \u044D\u043A\u0440\u0430\u043D / \u043E\u0431\u0440\u0430\u0442\u043D\u043E (F)"
  }, full ? "Выйти из полного экрана" : "Во весь экран"), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportTropeWalkWord(ex, data)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 12
  }), " Word"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportTropeWalkPdf(ex, data)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "note",
    size: 12
  }), " PDF"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportTropeWalkPptx(ex, data)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "image",
    size: 12
  }), " PowerPoint"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: onClose,
    title: "\u0417\u0430\u043A\u0440\u044B\u0442\u044C (Esc)"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 14
  }), " \u0417\u0430\u043A\u0440\u044B\u0442\u044C")), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      overflow: "auto",
      padding: "4vh 6vw",
      textAlign: "center"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "Arial, sans-serif",
      fontWeight: 800,
      fontSize: base * 2.2,
      lineHeight: 1.2,
      marginBottom: base * 1.2
    }
  }, heading), (classTags.length > 0 || tropeTags.length > 0) && /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: base,
      fontSize: base * 0.9
    }
  }, classTags.map(t => /*#__PURE__*/React.createElement("span", {
    key: t.id,
    style: {
      display: "inline-block",
      border: "2px solid #000",
      borderRadius: 999,
      padding: "2px 14px",
      margin: 4,
      fontWeight: 700
    }
  }, t.name, " \u043A\u043B.")), tropeTags.map(t => /*#__PURE__*/React.createElement("span", {
    key: t.id,
    style: {
      display: "inline-block",
      background: "#000",
      color: "#fff",
      borderRadius: 999,
      padding: "2px 14px",
      margin: 4,
      fontWeight: 700
    }
  }, t.name))), /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "Arial, sans-serif",
      fontWeight: 700,
      fontSize: base * 1.6,
      lineHeight: 1.55,
      maxWidth: "100%",
      margin: "0 auto"
    },
    dangerouslySetInnerHTML: {
      __html: ex.text || "<i>Текст не заполнен</i>"
    }
  }), ex.note && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: base,
      fontSize: base * 0.8,
      color: "#555"
    }
  }, ex.note)));
}
function ExampleCard({
  data,
  ex,
  usage,
  onOpen,
  selected,
  onToggle,
  dim
}) {
  const box = onToggle ? /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    className: "selbox",
    checked: !!selected,
    readOnly: true,
    title: "\u0412\u044B\u0431\u0440\u0430\u0442\u044C (Shift+\u043A\u043B\u0438\u043A \u2014 \u0432\u044B\u0431\u0440\u0430\u0442\u044C \u0434\u0438\u0430\u043F\u0430\u0437\u043E\u043D)",
    onClick: e => {
      e.stopPropagation();
      onToggle(e.shiftKey);
    }
  }) : null;
  const cls = "card" + (onToggle ? " hasbox" : "") + (selected ? " sel" : "");
  const dimStyle = dim ? {
    opacity: 0.5
  } : undefined;
  if (ex.typeTagId === TROPE_WALK_TYPE_ID) {
    const author = ex.authorId ? data.personalities.find(p => p.id === ex.authorId) : null;
    const classTags = nsTagIds(ex.tagIds, "classTag", data).map(id => data.tags.classTag.find(t => t.id === id)).filter(Boolean);
    const tropeTags = nsTagIds(ex.tagIds, "tropeTag", data).map(id => data.tags.tropeTag.find(t => t.id === id)).filter(Boolean);
    return /*#__PURE__*/React.createElement("div", {
      className: cls,
      style: dimStyle,
      onClick: onOpen
    }, box, /*#__PURE__*/React.createElement("div", {
      className: "ct-title"
    }, author ? author.title : "Автор не указан", ex.workTitle ? `, «${ex.workTitle}»` : ""), (classTags.length > 0 || tropeTags.length > 0) && /*#__PURE__*/React.createElement("div", {
      className: "ct-tags"
    }, classTags.map(t => /*#__PURE__*/React.createElement(Badge, {
      key: t.id,
      tag: t
    })), tropeTags.map(t => /*#__PURE__*/React.createElement(Badge, {
      key: t.id,
      tag: t
    }))), usage > 0 && /*#__PURE__*/React.createElement("div", {
      className: "ct-sub"
    }, "\u2713 \u0438\u0441\u043F\u043E\u043B\u044C\u0437\u0443\u0435\u0442\u0441\u044F \u0432 ", usage, " ", usage === 1 ? "уроке" : "уроках", (() => {
      const k = Array.from(exerciseUsedScopeKeys(data, ex.id, null)).map(x => scopeLabelOf(data, x));
      return k.length ? " (" + k.join(", ") + ")" : "";
    })()), (ex.hiddenFor || []).length > 0 && /*#__PURE__*/React.createElement("div", {
      className: "ct-sub"
    }, "\uD83D\uDEAB \u041D\u0435 \u043E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0435\u0442\u0441\u044F: ", (ex.hiddenFor || []).map(x => scopeLabelOf(data, x)).join(", ")), (ex.alwaysFor || []).length > 0 && /*#__PURE__*/React.createElement("div", {
      className: "ct-sub"
    }, "\u2605 \u0412\u0441\u0435\u0433\u0434\u0430: ", (ex.alwaysFor || []).map(x => scopeLabelOf(data, x)).join(", ")));
  }
  return /*#__PURE__*/React.createElement("div", {
    className: cls,
    style: dimStyle,
    onClick: onOpen
  }, box, /*#__PURE__*/React.createElement("div", {
    className: "ct-preview"
  }, stripHtml(ex.text).slice(0, 160) || "Пусто…"), usage > 0 && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\u2713 \u0438\u0441\u043F\u043E\u043B\u044C\u0437\u0443\u0435\u0442\u0441\u044F \u0432 ", usage, " ", usage === 1 ? "уроке" : "уроках", (() => {
    const k = Array.from(exerciseUsedScopeKeys(data, ex.id, null)).map(x => scopeLabelOf(data, x));
    return k.length ? " (" + k.join(", ") + ")" : "";
  })()), (ex.hiddenFor || []).length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\uD83D\uDEAB \u041D\u0435 \u043E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0435\u0442\u0441\u044F: ", (ex.hiddenFor || []).map(x => scopeLabelOf(data, x)).join(", ")), (ex.alwaysFor || []).length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\u2605 \u0412\u0441\u0435\u0433\u0434\u0430: ", (ex.alwaysFor || []).map(x => scopeLabelOf(data, x)).join(", ")), ((ex.forClasses || []).length > 0 || (ex.forCourseIds || []).length > 0) && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\u0414\u043B\u044F: ", [...(ex.forClasses || []).map(c => c + " кл."), ...(ex.forCourseIds || []).map(id => (data.courses.find(c => c.id === id) || {}).name).filter(Boolean)].join(", ")));
}
function ExamplePeekModal({
  data,
  update,
  projectPath,
  exampleId,
  onClose,
  onDelete,
  onSetSlot
}) {
  const ex = data.examples.find(e => e.id === exampleId);
  const [editing, setEditing] = useState(false);
  const [printOpen, setPrintOpen] = useState(false);
  if (!ex) return null;
  const type = data.tags.exerciseType.find(t => t.id === ex.typeTagId);
  const isTropeWalk = ex.typeTagId === TROPE_WALK_TYPE_ID;
  const author = isTropeWalk && ex.authorId ? data.personalities.find(p => p.id === ex.authorId) : null;
  const qTag = data.tags.questionType.find(t => t.id === ex.questionTypeTagId);
  const tags = (ex.tagIds || []).map(id => findTagAnywhere(data, id)).filter(Boolean);
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, type ? type.name : "Упражнение"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, editing ? /*#__PURE__*/React.createElement(ExampleForm, {
    data: data,
    update: update,
    projectPath: projectPath,
    typeId: ex.typeTagId,
    editId: ex.id,
    onDone: () => setEditing(false)
  }) : /*#__PURE__*/React.createElement(React.Fragment, null, isTropeWalk && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("h2", {
    style: {
      fontFamily: "var(--font-display)",
      marginBottom: 10
    }
  }, author ? author.title : "Автор не указан", ex.workTitle ? `, «${ex.workTitle}»` : ""), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportTropeWalkWord(ex, data)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 12
  }), " Word"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportTropeWalkPdf(ex, data)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "note",
    size: 12
  }), " PDF"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => exportTropeWalkPptx(ex, data)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "image",
    size: 12
  }), " PowerPoint"))), /*#__PURE__*/React.createElement("div", {
    dangerouslySetInnerHTML: {
      __html: ex.text || "<p><i>Пусто</i></p>"
    }
  }), ex.note && /*#__PURE__*/React.createElement("div", {
    className: "noterow",
    style: {
      marginTop: 10
    }
  }, ex.note), (ex.questions || []).length > 0 && /*#__PURE__*/React.createElement(Field, {
    label: "\u0412\u043E\u043F\u0440\u043E\u0441\u044B \u043A \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044E"
  }, (ex.questions || []).map(qq => {
    const qt = data.tags.questionType.find(t => t.id === qq.questionTypeTagId);
    return /*#__PURE__*/React.createElement("div", {
      key: qq.id,
      className: "row",
      style: {
        padding: "3px 0",
        alignItems: "flex-start"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        fontSize: 13
      }
    }, qq.text, qt && /*#__PURE__*/React.createElement("span", {
      style: {
        marginLeft: 6
      }
    }, /*#__PURE__*/React.createElement(Badge, {
      tag: qt
    }))), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => update(d => {
        const e2 = d.examples.find(e => e.id === ex.id);
        if (e2) e2.questions = (e2.questions || []).filter(x => x.id !== qq.id);
        return d;
      }),
      title: "\u0423\u0431\u0440\u0430\u0442\u044C \u0432\u043E\u043F\u0440\u043E\u0441 \u0438\u0437 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "trash",
      size: 11
    })));
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u041D\u0435 \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0442\u044C / \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0442\u044C \u0432\u0441\u0435\u0433\u0434\u0430 \u043F\u0440\u0438 \u0441\u043E\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u0438\u0438 \u0443\u0440\u043E\u043A\u043E\u0432"
  }, /*#__PURE__*/React.createElement(HiddenScopesRow, {
    data: data,
    hiddenFor: ex.hiddenFor,
    alwaysFor: ex.alwaysFor,
    onChange: v => update(d => {
      const e2 = d.examples.find(e => e.id === ex.id);
      if (e2) {
        e2.hiddenFor = v.hiddenFor;
        e2.alwaysFor = v.alwaysFor;
      }
      return d;
    })
  })), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginTop: 10
    }
  }, qTag && /*#__PURE__*/React.createElement(Badge, {
    tag: qTag
  }), tags.map(t => /*#__PURE__*/React.createElement(Badge, {
    key: t.id,
    tag: t
  }))), (ex.assignedStudentIds || []).length > 0 && /*#__PURE__*/React.createElement(Field, {
    label: "\u0423\u0447\u0435\u043D\u0438\u043A\u0438"
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap"
  }, ex.assignedStudentIds.map(id => {
    const s = data.calendar.students.find(x => x.id === id);
    return s ? /*#__PURE__*/React.createElement("span", {
      key: id,
      className: "chip"
    }, s.name, " \xB7 ", s.class, " \u043A\u043B.") : null;
  }))), type && type.slots.length > 0 && /*#__PURE__*/React.createElement(Field, {
    label: "\u0421\u043B\u043E\u0442 (\u0446\u0435\u043F\u043E\u0447\u043A\u0430)"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: ex.slotId || "",
    onChange: e => onSetSlot(e.target.value || null)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0411\u0435\u0437 \u0441\u043B\u043E\u0442\u0430"), type.slots.map(s => /*#__PURE__*/React.createElement("option", {
    key: s.id,
    value: s.id
  }, s.name)))), /*#__PURE__*/React.createElement(AddToLessonPicker, {
    data: data,
    update: update,
    itemType: "exercise",
    refId: ex.id
  }))), printOpen && /*#__PURE__*/React.createElement(PrintMaterialModal, {
    data: data,
    type: "exercise",
    entityId: ex.id,
    lesson: null,
    onClose: () => setPrintOpen(false)
  }), !editing && /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "danger",
    onClick: onDelete
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  }), " \u0423\u0434\u0430\u043B\u0438\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => setPrintOpen(true),
    title: "\u0420\u0430\u0441\u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C \u0442\u0435\u043A\u0441\u0442 \u0438 \u0432\u043E\u043F\u0440\u043E\u0441\u044B \u044D\u0442\u043E\u0433\u043E \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F"
  }, "\uD83D\uDDA8 \u041F\u0435\u0447\u0430\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
    onClick: () => setEditing(true)
  }, "\u0418\u0437\u043C\u0435\u043D\u0438\u0442\u044C"))));
}
function ExerciseTypePage({
  data,
  update,
  typeId,
  onBack,
  projectPath
}) {
  const type = data.tags.exerciseType.find(t => t.id === typeId);
  const isTropeWalk = typeId === TROPE_WALK_TYPE_ID;
  const [openExampleId, setOpenExampleId] = useState(null);
  const [creating, setCreating] = useState(false);
  const [viewMode, setViewMode] = useState("flat");
  const [poemImport, setPoemImport] = useState(false);
  const [sel, setSel] = useState([]);
  const [hideOpen, setHideOpen] = useState(false);
  const [ctxFilter, setCtxFilter] = useState("");
  const [showHiddenToo, setShowHiddenToo] = useState(false);
  const lastSelRef = useRef(null);
  if (!type) return /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u0412\u0438\u0434 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D. ", /*#__PURE__*/React.createElement(Btn, {
    onClick: onBack
  }, "\u041D\u0430\u0437\u0430\u0434"));
  const allExamples = data.examples.filter(e => e.typeTagId === typeId);
  const filterCtx = ctxFilter ? scopeFromKey(ctxFilter) : null;
  const hiddenReasonOf = e => filterCtx ? exerciseHiddenReason(data, e, filterCtx, null) : null;
  const examples = filterCtx && !showHiddenToo ? allExamples.filter(e => !hiddenReasonOf(e)) : allExamples;
  const usageCount = exId => data.lessons.reduce((n, l) => n + l.items.filter(it => it.type === "exercise" && it.refId === exId).length, 0);
  // ---- Выбор карточек чек-боксами (Shift+клик — диапазон) и удаление нескольких сразу ----
  let visibleOrder = examples.map(e => e.id);
  const selIds = sel.filter(id => examples.some(e => e.id === id));
  const toggleSel = (id, shift) => {
    const last = lastSelRef.current;
    const on = !sel.includes(id);
    let ids = [id];
    if (shift && last && last !== id) {
      const a = visibleOrder.indexOf(last),
        b = visibleOrder.indexOf(id);
      if (a >= 0 && b >= 0) ids = visibleOrder.slice(Math.min(a, b), Math.max(a, b) + 1);
    }
    lastSelRef.current = id;
    setSel(prev => on ? Array.from(new Set([...prev, ...ids])) : prev.filter(x => !ids.includes(x)));
  };
  const renderCard = ex => /*#__PURE__*/React.createElement(ExampleCard, {
    key: ex.id,
    data: data,
    ex: ex,
    usage: usageCount(ex.id),
    onOpen: () => setOpenExampleId(ex.id),
    selected: sel.includes(ex.id),
    onToggle: shift => toggleSel(ex.id, shift),
    dim: !!hiddenReasonOf(ex)
  });
  const groupHead = (name, ids) => {
    const all = ids.length > 0 && ids.every(id => sel.includes(id));
    return /*#__PURE__*/React.createElement("div", {
      className: "group-h"
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: all,
      title: "\u0412\u044B\u0431\u0440\u0430\u0442\u044C \u0432\u0441\u044E \u0433\u0440\u0443\u043F\u043F\u0443",
      style: {
        marginRight: 8,
        verticalAlign: "middle",
        cursor: "pointer"
      },
      onChange: e => setSel(prev => e.target.checked ? Array.from(new Set([...prev, ...ids])) : prev.filter(x => !ids.includes(x)))
    }), name);
  };
  const removeSelected = async () => {
    const ids = selIds.slice();
    if (!ids.length) return;
    const used = ids.filter(id => usageCount(id) > 0).length;
    if (!(await Confirm(`Удалить выбранные упражнения (${ids.length})?` + (used ? ` ${used} из них используются в уроках — они исчезнут и из уроков.` : "")))) return;
    update(d => {
      d.examples = d.examples.filter(e => !ids.includes(e.id));
      d.lessons.forEach(l => {
        l.items = l.items.filter(it => !(it.type === "exercise" && ids.includes(it.refId)) && !(it.type === "question" && it.entityType === "exercise" && ids.includes(it.entityId)));
      });
      d.reserves = (d.reserves || []).filter(r => !ids.includes(r.exerciseId));
      return d;
    });
    setSel([]);
    showToast("Удалено упражнений: " + ids.length);
  };

  // Для «Прогулки по тропам» — автоматическая группировка по автору/классу/тропу: каждый новый
  // тег (или автор) сам заводит себе «папку», в которую попадают все подходящие упражнения.
  let groupedBody = null;
  if (isTropeWalk && viewMode !== "flat") {
    const mode = TROPE_VIEW_MODES.find(m => m.id === viewMode);
    const groups = {};
    const pushTo = (key, e) => {
      (groups[key] = groups[key] || []).push(e);
    };
    if (mode.id === "authors") {
      examples.forEach(e => pushTo(e.authorId || "__none", e));
    } else {
      examples.forEach(e => {
        const ids = nsTagIds(e.tagIds, mode.groupBy, data);
        if (ids.length === 0) pushTo("__none", e);else ids.forEach(id => pushTo(id, e));
      });
    }
    const nameOf = key => {
      if (key === "__none") return mode.id === "authors" ? "Без автора" : "Без тега";
      if (mode.id === "authors") return (data.personalities.find(p => p.id === key) || {}).title || "—";
      return (data.tags[mode.groupBy].find(t => t.id === key) || {}).name || "—";
    };
    const keys = Object.keys(groups).filter(k => k !== "__none").sort((a, b) => nameOf(a).localeCompare(nameOf(b), "ru"));
    if (groups.__none) keys.push("__none");
    visibleOrder = keys.flatMap(k => groups[k].map(e => e.id));
    groupedBody = keys.map(key => /*#__PURE__*/React.createElement("div", {
      key: key
    }, groupHead(nameOf(key), groups[key].map(e => e.id)), /*#__PURE__*/React.createElement("div", {
      className: "grid"
    }, groups[key].map(ex => renderCard(ex)))));
  }
  const addSlot = async () => {
    const name = await PromptStr("Название слота (цепочки):");
    if (!name || !name.trim()) return;
    update(d => {
      const t = d.tags.exerciseType.find(x => x.id === typeId);
      t.slots.push({
        id: uid(),
        name: name.trim(),
        order: t.slots.length
      });
      return d;
    });
  };
  const removeSlot = async slotId => {
    if (!(await Confirm("Удалить слот? Упражнения из него не удалятся, просто станут «без слота»."))) return;
    update(d => {
      const t = d.tags.exerciseType.find(x => x.id === typeId);
      t.slots = t.slots.filter(s => s.id !== slotId);
      d.examples.forEach(e => {
        if (e.typeTagId === typeId && e.slotId === slotId) e.slotId = null;
      });
      return d;
    });
  };
  const removeExample = async id => {
    if (!(await Confirm("Удалить упражнение?"))) return;
    update(d => {
      d.examples = d.examples.filter(e => e.id !== id);
      d.lessons.forEach(l => {
        l.items = l.items.filter(it => !(it.type === "exercise" && it.refId === id) && !(it.type === "question" && it.entityType === "exercise" && it.entityId === id));
      });
      d.reserves = (d.reserves || []).filter(r => r.exerciseId !== id);
      return d;
    });
    setSel(prev => prev.filter(x => x !== id));
    setOpenExampleId(null);
  };
  const setSlot = (exId, slotId) => update(d => {
    const ex = d.examples.find(e => e.id === exId);
    if (ex) ex.slotId = slotId || null;
    return d;
  });
  const renameThisType = async () => {
    const name = await PromptStr("Новое название вида упражнения:", type.name);
    if (!name || !name.trim()) return;
    update(d => {
      const x = d.tags.exerciseType.find(y => y.id === typeId);
      if (x) x.name = name.trim();
      return d;
    });
  };
  const deleteThisType = async () => {
    if (type.locked) return;
    const warn = allExamples.length > 0 ? ` Все его упражнения (${allExamples.length}) будут удалены вместе с ним, включая ссылки на них в уроках и в запасе.` : "";
    if (!(await Confirm(`Удалить вид «${type.name}»?${warn}`))) return;
    const exIds = allExamples.map(e => e.id);
    update(d => {
      d.tags.exerciseType = d.tags.exerciseType.filter(x => x.id !== typeId);
      d.examples = d.examples.filter(e => !exIds.includes(e.id));
      d.lessons.forEach(l => {
        l.items = l.items.filter(it => !(it.type === "exercise" && exIds.includes(it.refId)) && !(it.type === "question" && it.entityType === "exercise" && exIds.includes(it.entityId)));
      });
      d.reserves = (d.reserves || []).filter(r => !exIds.includes(r.exerciseId));
      return d;
    });
    onBack();
  };
  const applyHide = (keys, mode) => {
    const ids = selIds.slice();
    update(d => {
      d.examples.forEach(e => {
        if (!ids.includes(e.id)) return;
        applyScopeMode(e, keys, mode);
      });
      return d;
    });
    setHideOpen(false);
    showToast(scopeModeToast(mode) + keys.map(k => scopeLabelOf(data, k)).join(", ") + " · карточек: " + ids.length);
  };
  const bySlot = {};
  examples.forEach(e => {
    const k = e.slotId || "__none";
    (bySlot[k] = bySlot[k] || []).push(e);
  });
  if (!isTropeWalk && type.slots.length > 0) visibleOrder = type.slots.flatMap(s => (bySlot[s.id] || []).map(e => e.id)).concat((bySlot.__none || []).map(e => e.id));
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onBack
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "left",
    size: 15
  }), " \u041D\u0430\u0437\u0430\u0434"), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), isTropeWalk && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: downloadPoemTemplateDocx,
    title: "\u0421\u043A\u0430\u0447\u0430\u0442\u044C \u043E\u0431\u0440\u0430\u0437\u0435\u0446 Word-\u0444\u0430\u0439\u043B\u0430, \u043F\u043E \u043A\u043E\u0442\u043E\u0440\u043E\u043C\u0443 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u043E\u0437\u0434\u0430\u0451\u0442 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "note",
    size: 13
  }), " \u0428\u0430\u0431\u043B\u043E\u043D \u0434\u043B\u044F \u0438\u043C\u043F\u043E\u0440\u0442\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    onClick: () => setPoemImport(true),
    title: "\u0421\u043E\u0437\u0434\u0430\u0442\u044C \u0441\u0440\u0430\u0437\u0443 \u043C\u043D\u043E\u0433\u043E \u043A\u0430\u0440\u0442\u043E\u0447\u0435\u043A \u0438\u0437 \u0444\u0430\u0439\u043B\u0430 \u0441\u043E \u0441\u0431\u043E\u0440\u043D\u0438\u043A\u043E\u043C (Word, PDF): \u0443\u043A\u0430\u0437\u044B\u0432\u0430\u0435\u0442\u0435 \u0430\u0432\u0442\u043E\u0440\u0430, \u043E\u0441\u0442\u0430\u043B\u044C\u043D\u043E\u0435 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0434\u0435\u043B\u0430\u0435\u0442 \u0441\u0430\u043C\u0430"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 13
  }), " \u0418\u043C\u043F\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0442\u0435\u043A\u0441\u0442\u044B")), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: () => setCreating(true)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u0423\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0435")), /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      alignItems: "center",
      gap: 8,
      marginBottom: 4
    }
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      fontFamily: "var(--font-display)",
      margin: 0
    }
  }, type.name), type.locked && /*#__PURE__*/React.createElement("span", {
    className: "badge-lock",
    title: "\u041F\u043E\u0441\u0442\u043E\u044F\u043D\u043D\u044B\u0439 \u0432\u0438\u0434 \u2014 \u043D\u0435\u043B\u044C\u0437\u044F \u0443\u0434\u0430\u043B\u0438\u0442\u044C"
  }, "\uD83D\uDD12"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: renameThisType,
    title: "\u041F\u0435\u0440\u0435\u0438\u043C\u0435\u043D\u043E\u0432\u0430\u0442\u044C \u0432\u0438\u0434"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "edit",
    size: 12
  })), !type.locked && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: deleteThisType,
    title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C \u0432\u0438\u0434"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 12
  }))), isTropeWalk ? /*#__PURE__*/React.createElement(ViewModeToggle, {
    value: viewMode,
    onChange: setViewMode,
    options: TROPE_VIEW_MODES
  }) : /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      margin: "10px 0"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: addSlot
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u0421\u043B\u043E\u0442 (\u0446\u0435\u043F\u043E\u0447\u043A\u0430)"), type.slots.map(s => /*#__PURE__*/React.createElement("span", {
    key: s.id,
    className: "chip"
  }, s.name, /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => removeSlot(s.id)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  }))))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      margin: "6px 0",
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u2605 \u0412\u0441\u0435 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F \u044D\u0442\u043E\u0433\u043E \u0432\u0438\u0434\u0430 \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u044E\u0442\u0441\u044F \u0432\u0441\u0435\u0433\u0434\u0430 \u0432:"), /*#__PURE__*/React.createElement(AlwaysScopesRow, {
    data: data,
    scopes: type.alwaysFor,
    onChange: arr => update(d => {
      const t = d.tags.exerciseType.find(x => x.id === typeId);
      if (t) t.alwaysFor = arr;
      return d;
    })
  })), creating && /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement(ExampleForm, {
    data: data,
    update: update,
    projectPath: projectPath,
    typeId: typeId,
    onDone: () => setCreating(false)
  }))), examples.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      margin: "10px 0 4px",
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setSel(selIds.length === examples.length ? [] : examples.map(e => e.id))
  }, selIds.length === examples.length ? "Снять выбор" : "Выбрать все (" + examples.length + ")"), selIds.length > 0 && selIds.length < examples.length && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setSel([])
  }, "\u0421\u043D\u044F\u0442\u044C \u0432\u044B\u0431\u043E\u0440"), selIds.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    sm: true,
    onClick: () => setHideOpen("hide"),
    title: "\u041D\u0435 \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0442\u044C \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u044B\u0435 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 \u0432 \u0443\u0440\u043E\u043A\u0430\u0445 \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u044B\u0445 \u043A\u043B\u0430\u0441\u0441\u043E\u0432/\u043B\u0438\u0442\u0435\u0440"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 12
  }), " \u041D\u0435 \u043E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0442\u044C \u0432\u2026 (", selIds.length, ")"), selIds.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setHideOpen("always"),
    title: "\u041F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0442\u044C \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u044B\u0435 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 \u0432 \u044D\u0442\u0438\u0445 \u043A\u043B\u0430\u0441\u0441\u0430\u0445/\u043B\u0438\u0442\u0435\u0440\u0430\u0445 \u0432\u0441\u0435\u0433\u0434\u0430 \u2014 \u0434\u0430\u0436\u0435 \u0435\u0441\u043B\u0438 \u0438\u0441\u043F\u043E\u043B\u044C\u0437\u043E\u0432\u0430\u043D\u044B \u0438\u043B\u0438 \u043E\u0442\u043A\u043B\u044E\u0447\u0435\u043D\u044B"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "check",
    size: 12
  }), " \u041E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0442\u044C \u0432\u0441\u0435\u0433\u0434\u0430\u2026 (", selIds.length, ")"), selIds.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "danger",
    sm: true,
    onClick: removeSelected
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 12
  }), " \u0423\u0434\u0430\u043B\u0438\u0442\u044C \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u044B\u0435 (", selIds.length, ")"), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041E\u0442\u043C\u0435\u0447\u0430\u0439\u0442\u0435 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 \u0433\u0430\u043B\u043E\u0447\u043A\u0430\u043C\u0438; Shift+\u043A\u043B\u0438\u043A \u043F\u043E \u0433\u0430\u043B\u043E\u0447\u043A\u0435 \u0432\u044B\u0431\u0438\u0440\u0430\u0435\u0442 \u0432\u0441\u0451 \u043C\u0435\u0436\u0434\u0443 \u0434\u0432\u0443\u043C\u044F \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0430\u043C\u0438.")), allExamples.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      margin: "4px 0 8px",
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u0442\u044C \u0434\u043B\u044F:"), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: ctxFilter,
    onChange: e => setCtxFilter(e.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0412\u0441\u0435 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F"), CLASSES.map(c => /*#__PURE__*/React.createElement(React.Fragment, {
    key: c
  }, /*#__PURE__*/React.createElement("option", {
    value: "c:" + c
  }, c, " \u043A\u043B\u0430\u0441\u0441 (\u0431\u0435\u0437 \u043B\u0438\u0442\u0435\u0440\u044B)"), lettersForGrade(data, c).map(cl => /*#__PURE__*/React.createElement("option", {
    key: cl.id,
    value: "l:" + c + ":" + cl.id
  }, c, cl.name)))), data.courses.map(c => /*#__PURE__*/React.createElement("option", {
    key: c.id,
    value: "k:" + c.id
  }, "\u041A\u0443\u0440\u0441 \xAB", c.name, "\xBB"))), filterCtx && /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 5,
      fontSize: 12.5
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: showHiddenToo,
    onChange: e => setShowHiddenToo(e.target.checked)
  }), " \u043F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u0442\u044C \u0438 \u0441\u043A\u0440\u044B\u0442\u044B\u0435 (\u0437\u0430\u0442\u0435\u043C\u043D\u0435\u043D\u044B)"), filterCtx && /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u0421\u043A\u0440\u044B\u0442\u043E \u0442\u0430\u043C: ", allExamples.filter(e => hiddenReasonOf(e)).length, " \u0438\u0437 ", allExamples.length)), isTropeWalk ? /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 10
    }
  }, examples.length === 0 && !creating && /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041F\u043E\u043A\u0430 \u043D\u0435\u0442 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439 \u044D\u0442\u043E\u0433\u043E \u0432\u0438\u0434\u0430."), groupedBody || /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, examples.map(ex => renderCard(ex)))) : type.slots.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, examples.map(ex => renderCard(ex)), examples.length === 0 && !creating && /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041F\u043E\u043A\u0430 \u043D\u0435\u0442 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439 \u044D\u0442\u043E\u0433\u043E \u0432\u0438\u0434\u0430.")) : /*#__PURE__*/React.createElement(React.Fragment, null, type.slots.map(s => /*#__PURE__*/React.createElement("div", {
    key: s.id
  }, groupHead(s.name, (bySlot[s.id] || []).map(e => e.id)), /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, (bySlot[s.id] || []).map(ex => renderCard(ex))))), /*#__PURE__*/React.createElement("div", null, groupHead("Без слота", (bySlot.__none || []).map(e => e.id)), /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, (bySlot.__none || []).map(ex => renderCard(ex))))), hideOpen && /*#__PURE__*/React.createElement(HidePoemsModal, {
    data: data,
    count: selIds.length,
    variant: hideOpen,
    onApply: applyHide,
    onClose: () => setHideOpen(false)
  }), poemImport && /*#__PURE__*/React.createElement(PoemImportModal, {
    data: data,
    update: update,
    onClose: () => setPoemImport(false)
  }), openExampleId && /*#__PURE__*/React.createElement(ExamplePeekModal, {
    data: data,
    update: update,
    projectPath: projectPath,
    exampleId: openExampleId,
    onClose: () => setOpenExampleId(null),
    onDelete: () => removeExample(openExampleId),
    onSetSlot: slotId => setSlot(openExampleId, slotId)
  }));
}
function ExerciseTile({
  type,
  count,
  onOpen,
  onRename,
  onDelete
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "card",
    style: {
      borderLeft: `4px solid ${type.color}`,
      cursor: "pointer",
      minHeight: "auto"
    },
    onClick: () => onOpen(type.id)
  }, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      alignItems: "flex-start",
      gap: 4
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "ct-title",
    style: {
      flex: 1
    }
  }, type.name, type.locked && /*#__PURE__*/React.createElement("span", {
    className: "badge-lock",
    title: "\u041F\u043E\u0441\u0442\u043E\u044F\u043D\u043D\u044B\u0439 \u0432\u0438\u0434 \u2014 \u043D\u0435\u043B\u044C\u0437\u044F \u0443\u0434\u0430\u043B\u0438\u0442\u044C"
  }, "\uD83D\uDD12")), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onRename(type.id);
    },
    title: "\u041F\u0435\u0440\u0435\u0438\u043C\u0435\u043D\u043E\u0432\u0430\u0442\u044C \u0432\u0438\u0434"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "edit",
    size: 11
  })), !type.locked && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onDelete(type.id);
    },
    title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C \u0432\u0438\u0434"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 11
  }))), /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, count, " ", count === 1 ? "упражнение" : count >= 2 && count <= 4 ? "упражнения" : "упражнений"));
}
function ExampleMiniCard({
  ex,
  onOpen
}) {
  return /*#__PURE__*/React.createElement("div", {
    className: "card",
    onClick: onOpen
  }, /*#__PURE__*/React.createElement("div", {
    className: "ct-preview"
  }, stripHtml(ex.text).slice(0, 160) || "Пусто"));
}
function ExerciseBankTab({
  data,
  update,
  projectPath,
  nav,
  setNav
}) {
  const openTypeId = nav.exercises.openTypeId;
  const q = nav.exercises.q;
  const setOpenTypeId = v => setNav(n => ({
    ...n,
    exercises: {
      ...n.exercises,
      openTypeId: v
    }
  }));
  const setQ = v => setNav(n => ({
    ...n,
    exercises: {
      ...n.exercises,
      q: v
    }
  }));
  if (openTypeId) return /*#__PURE__*/React.createElement(ExerciseTypePage, {
    data: data,
    update: update,
    typeId: openTypeId,
    onBack: () => setOpenTypeId(null),
    projectPath: projectPath
  });
  const addCategory = async () => {
    const name = await PromptStr("Название категории:");
    if (!name || !name.trim()) return;
    update(d => {
      d.categories.push({
        id: uid(),
        name: name.trim(),
        color: SWATCHES[Math.floor(Math.random() * SWATCHES.length)],
        order: d.categories.length
      });
      return d;
    });
  };
  const addType = async categoryId => {
    const name = await PromptStr("Название вида упражнения:");
    if (!name || !name.trim()) return;
    const cat = data.categories.find(c => c.id === categoryId);
    update(d => {
      d.tags.exerciseType.push({
        id: uid(),
        name: name.trim(),
        color: cat ? cat.color : SWATCHES[0],
        categoryId: categoryId || null,
        order: d.tags.exerciseType.length,
        slots: []
      });
      return d;
    });
  };
  const deleteCategory = async id => {
    if (!(await Confirm("Удалить категорию? Виды упражнений останутся, просто станут «без категории»."))) return;
    update(d => {
      d.categories = d.categories.filter(c => c.id !== id);
      d.tags.exerciseType.forEach(t => {
        if (t.categoryId === id) t.categoryId = null;
      });
      return d;
    });
  };
  const renameType = async typeId => {
    const t = data.tags.exerciseType.find(x => x.id === typeId);
    if (!t) return;
    const name = await PromptStr("Новое название вида упражнения:", t.name);
    if (!name || !name.trim()) return;
    update(d => {
      const x = d.tags.exerciseType.find(y => y.id === typeId);
      if (x) x.name = name.trim();
      return d;
    });
  };
  const deleteType = async typeId => {
    const t = data.tags.exerciseType.find(x => x.id === typeId);
    if (!t || t.locked) return;
    const exIds = data.examples.filter(e => e.typeTagId === typeId).map(e => e.id);
    const warn = exIds.length > 0 ? ` Все его упражнения (${exIds.length}) будут удалены вместе с ним, включая ссылки на них в уроках и в запасе.` : "";
    if (!(await Confirm(`Удалить вид «${t.name}»?${warn}`))) return;
    update(d => {
      d.tags.exerciseType = d.tags.exerciseType.filter(x => x.id !== typeId);
      d.examples = d.examples.filter(e => !exIds.includes(e.id));
      d.lessons.forEach(l => {
        l.items = l.items.filter(it => !(it.type === "exercise" && exIds.includes(it.refId)) && !(it.type === "question" && it.entityType === "exercise" && exIds.includes(it.entityId)));
      });
      d.reserves = (d.reserves || []).filter(r => !exIds.includes(r.exerciseId));
      return d;
    });
    if (openTypeId === typeId) setOpenTypeId(null);
  };
  const groups = {};
  data.tags.exerciseType.forEach(t => {
    const key = t.categoryId || "__none";
    (groups[key] = groups[key] || []).push(t);
  });
  const countFor = typeId => data.examples.filter(e => e.typeTagId === typeId).length;
  const searched = q.trim() ? data.examples.filter(e => stripHtml(e.text).toLowerCase().includes(q.toLowerCase()) || (e.note || "").toLowerCase().includes(q.toLowerCase())) : null;
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "toolbar"
  }, /*#__PURE__*/React.createElement(TagSearchBar, {
    value: q,
    onChange: setQ,
    placeholder: "\u041F\u043E\u0438\u0441\u043A \u043F\u043E \u0442\u0435\u043A\u0441\u0442\u0443 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439\u2026"
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: addCategory
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u041A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u044F")), searched ? /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, searched.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E."), searched.map(ex => /*#__PURE__*/React.createElement(ExampleMiniCard, {
    key: ex.id,
    ex: ex,
    onOpen: () => {
      const t = data.tags.exerciseType.find(x => x.id === ex.typeTagId);
      if (t) setOpenTypeId(t.id);
    }
  }))) : /*#__PURE__*/React.createElement(React.Fragment, null, data.categories.map(cat => /*#__PURE__*/React.createElement("div", {
    key: cat.id
  }, /*#__PURE__*/React.createElement("div", {
    className: "group-h",
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "tk-dot",
    style: {
      background: cat.color
    }
  }), cat.name, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => addType(cat.id),
    title: "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0432\u0438\u0434"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 11
  })), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => deleteCategory(cat.id),
    title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C \u043A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u044E"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 11
  }))), /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, (groups[cat.id] || []).map(t => /*#__PURE__*/React.createElement(ExerciseTile, {
    key: t.id,
    type: t,
    count: countFor(t.id),
    onOpen: setOpenTypeId,
    onRename: renameType,
    onDelete: deleteType
  })), (groups[cat.id] || []).length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041F\u043E\u043A\u0430 \u043D\u0435\u0442 \u0432\u0438\u0434\u043E\u0432 \u2014 \u043D\u0430\u0436\u043C\u0438\u0442\u0435 + \u0440\u044F\u0434\u043E\u043C \u0441 \u043D\u0430\u0437\u0432\u0430\u043D\u0438\u0435\u043C \u043A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u0438.")))), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "group-h",
    style: {
      display: "flex",
      alignItems: "center",
      gap: 8
    }
  }, "\u0411\u0435\u0437 \u043A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u0438", /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => addType(null),
    title: "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0432\u0438\u0434"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 11
  }))), /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, (groups.__none || []).map(t => /*#__PURE__*/React.createElement(ExerciseTile, {
    key: t.id,
    type: t,
    count: countFor(t.id),
    onOpen: setOpenTypeId,
    onRename: renameType,
    onDelete: deleteType
  })), (groups.__none || []).length === 0 && data.categories.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041F\u043E\u043A\u0430 \u043D\u0435\u0442 \u043D\u0438 \u043E\u0434\u043D\u043E\u0439 \u043A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u0438. \u041D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB\u041A\u0430\u0442\u0435\u0433\u043E\u0440\u0438\u044F\xBB, \u0447\u0442\u043E\u0431\u044B \u043D\u0430\u0447\u0430\u0442\u044C \u0431\u0430\u043D\u043A \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439.")))));
}
;
// ============================================================
// lessons.js — Уроки: классы/курсы, карточка урока, система уроков.
// ============================================================

const ITEM_TYPE_META = {
  exercise: {
    label: "Упражнение",
    dataKey: "examples",
    icon: "note",
    titleOf: (d, e) => {
      if (e.typeTagId === TROPE_WALK_TYPE_ID) {
        const a = e.authorId ? d.personalities.find(p => p.id === e.authorId) : null;
        const t = (a ? a.title : "") + (e.workTitle ? (a ? ", " : "") + "«" + e.workTitle + "»" : "");
        if (t) return t;
      }
      return stripHtml(e.text).slice(0, 70) || "Упражнение";
    }
  },
  theory: {
    label: "История/теория",
    dataKey: "theoryHistory",
    icon: "book",
    titleOf: (d, e) => e.title || "Без названия"
  },
  personality: {
    label: "Персоналия",
    dataKey: "personalities",
    icon: "user",
    titleOf: (d, e) => e.title || "Без имени"
  },
  work: {
    label: "Произведение",
    dataKey: "works",
    icon: "book",
    titleOf: (d, e) => e.title || "Без названия"
  }
};
const QUESTION_SOURCE_TYPES = ["theory", "personality", "work"];
function itemRefObj(data, item) {
  return (data[ITEM_TYPE_META[item.type].dataKey] || []).find(e => e.id === item.refId);
}
function questionSourceEntity(data, item) {
  return (data[ITEM_TYPE_META[item.entityType].dataKey] || []).find(e => e.id === item.entityId);
}
function itemGradeReasonText(data, item) {
  if (item.type === "exercise") {
    const ex = data.examples.find(e => e.id === item.refId);
    return ex ? stripHtml(ex.text).slice(0, 60) : "";
  }
  if (item.type === "question") {
    const q = questionOfItem(data, item);
    return q ? q.text.slice(0, 60) : "";
  }
  return "";
}

// ---- Скрытие стихотворений («Прогулка по тропам») в классах и литерах ----
// Область — «класс целиком» (c:6), «литера» (l:6:<id>) или курс (k:<id>). Стихотворение скрыто в области, если:
// (а) учитель отключил его вручную (ex.hiddenFor) или (б) оно уже стоит в одном из уроков этой области.
// Урок без литеры («6 класс») считается уроком всего класса — стихотворение прячется во всех литерах класса.
function scopeKeyOf(s) {
  return s.courseId ? "k:" + s.courseId : s.classLetterId ? "l:" + s.class + ":" + s.classLetterId : "c:" + s.class;
}
function scopeLabelOf(data, key) {
  const p = String(key).split(":");
  if (p[0] === "k") return (data.courses.find(c => c.id === p[1]) || {}).name || "курс";
  if (p[0] === "l") return p[1] + classLetterName(data, p[2]);
  return p[1] + " класс";
}
function scopeFromKey(key) {
  const p = String(key).split(":");
  if (p[0] === "k") return {
    class: null,
    courseId: p[1],
    classLetterId: null
  };
  if (p[0] === "l") return {
    class: p[1],
    courseId: null,
    classLetterId: p[2]
  };
  return {
    class: p[1],
    courseId: null,
    classLetterId: null
  };
}
// Где материал (упражнение/персоналия/теория/произведение) «использован»: в уроке стоит сам материал или вопрос из него.
function materialUsedScopeKeys(data, type, id, exceptLessonId) {
  const keys = new Set();
  data.lessons.forEach(l => {
    if (l.archived || l.id === exceptLessonId) return;
    const used = (l.items || []).some(it => it.type === type && it.refId === id || it.type === "question" && it.entityType === type && it.entityId === id);
    if (used) keys.add(scopeKeyOf(l));
  });
  return keys;
}
function exerciseUsedScopeKeys(data, exId, exceptLessonId) {
  return materialUsedScopeKeys(data, "exercise", exId, exceptLessonId);
}
// Классы/курсы, для которых назначен материал: поле «доступно для» + метки классов (у стихотворений и произведений).
function materialAudience(data, e) {
  const classes = new Set((e.forClasses || []).map(String));
  const letters = new Set();
  (e.tagIds || []).forEach(id => {
    const t = (data.tags.classTag || []).find(x => x.id === id);
    if (!t) return;
    const name = String(t.name || "");
    const r = name.match(/(\d{1,2})\s*[-–—]\s*(\d{1,2})/);
    if (r) {
      for (let i = +r[1]; i <= +r[2] && i <= 11; i++) classes.add(String(i));
      return;
    }
    const ml = name.match(/^\s*(\d{1,2})\s*[«"']?([А-Яа-яЁёA-Za-z])[»"']?\s*(?:класс\w*|кл\.?)?\s*$/);
    if (ml) {
      letters.add(String(+ml[1]) + ml[2].toUpperCase());
      return;
    }
    const m = name.match(/^\s*(\d{1,2})/);
    if (m) classes.add(String(+m[1]));
  });
  return {
    classes,
    letters,
    courses: new Set(e.forCourseIds || [])
  };
}
function materialAvailableFor(data, e, ctx) {
  const a = materialAudience(data, e);
  if (!a.classes.size && !a.letters.size && !a.courses.size) return true;
  if (ctx.courseId) return a.courses.has(ctx.courseId);
  if (a.classes.has(String(ctx.class))) return true;
  const ln = ctx.classLetterId ? classLetterName(data, ctx.classLetterId) : "";
  return !!ln && a.letters.has(String(ctx.class) + ln.toUpperCase());
}
// null — показывать; "audience" — назначен другим классам/курсам; "manual" — отключён вручную; "used" — уже использован в уроке этой области
function materialHiddenReason(data, type, e, ctx, exceptLessonId) {
  if (!e) return null;
  const key = scopeKeyOf(ctx);
  const classKey = ctx.courseId ? null : "c:" + ctx.class;
  const hits = keys => keys.has(key) || !!classKey && keys.has(classKey);
  // «Отображать всегда» (у карточки или у всего вида упражнений) — сильнее метки класса, ручного отключения и «уже использовано»
  const alwaysKeys = new Set(e.alwaysFor || []);
  if (type === "exercise") {
    const et = (data.tags.exerciseType || []).find(t => t.id === e.typeTagId);
    if (et) (et.alwaysFor || []).forEach(k => alwaysKeys.add(k));
  }
  if (hits(alwaysKeys)) return null;
  if (!materialAvailableFor(data, e, ctx)) return "audience";
  if (hits(new Set(e.hiddenFor || []))) return "manual";
  if (hits(materialUsedScopeKeys(data, type, e.id, exceptLessonId))) return "used";
  return null;
}
function exerciseHiddenReason(data, ex, ctx, exceptLessonId) {
  return materialHiddenReason(data, "exercise", ex, ctx, exceptLessonId);
}
function hiddenReasonText(reason, ctxLabel) {
  return reason === "manual" ? "отключено вручную" : reason === "audience" ? "назначено другим классам" : "уже использовано в " + ctxLabel;
}

// ---- Вопросы к элементам урока ----
function entityQuestionsOf(ent, type) {
  return type === "exercise" ? ent.questions || [] : ent.course?.questions || [];
}
function entityTextHtmlOf(ent, type) {
  return type === "exercise" ? ent.text || "" : ent.course?.lectureHtml || "";
}
// Вопрос пункта урока: из карточки (персоналия/теория/произведение/упражнение) или «только в уроке» (item.local).
function questionOfItem(data, it) {
  if (!it || it.type !== "question") return null;
  if (it.local) return {
    id: it.questionId,
    text: it.local.text || "",
    questionTypeTagId: it.local.questionTypeTagId ?? null,
    local: true
  };
  const ent = questionSourceEntity(data, it);
  if (!ent) return null;
  const q = entityQuestionsOf(ent, it.entityType).find(x => x.id === it.questionId);
  return q ? {
    ...q,
    local: false
  } : null;
}
// Убирает из текста карточки привязки (подсветку) конкретного вопроса — нужно, когда вопрос «только в уроке» удаляют.
function stripQaAnchors(d, entityType, entityId, qid) {
  const meta = ITEM_TYPE_META[entityType];
  const ent = meta ? (d[meta.dataKey] || []).find(e => e.id === entityId) : null;
  if (!ent) return;
  const html = entityTextHtmlOf(ent, entityType);
  if (!html || html.indexOf(qid) < 0) return;
  const doc = new DOMParser().parseFromString(html, "text/html");
  const els = Array.from(doc.querySelectorAll(`[data-qa-id="${CSS.escape(qid)}"]`));
  if (!els.length) return;
  els.forEach(el => {
    const p = el.parentNode;
    while (el.firstChild) p.insertBefore(el.firstChild, el);
    p.removeChild(el);
  });
  if (entityType === "exercise") ent.text = doc.body.innerHTML;else if (ent.course) ent.course.lectureHtml = doc.body.innerHTML;
}
// Список учеников, вызванных к доске за конкретный элемент урока (упражнение/вопрос) — можно
// сразу несколько; каждая запись хранит своё markId, чтобы правило ту же отметку, а не плодило новые.
function itemStudents(item) {
  return Array.isArray(item.students) ? item.students : [];
}

// Находит/заводит ученика в журнале (по имени + классу/литере/курсу этого урока) и синхронизирует
// отметку по одной записи ученика на элементе урока (упражнению или вопросу) с общим списком отметок.
function syncStudentEntryMark(d, lesson, item, entry) {
  const name = (entry.name || "").trim();
  const gradeRaw = (entry.grade || "").trim();
  if (!name || !gradeRaw) {
    if (entry.markId) {
      d.calendar.marks = d.calendar.marks.filter(m => m.id !== entry.markId);
      entry.markId = null;
    }
    return;
  }
  // Новых учеников здесь не заводим: список класса пополняется только в «Списках классов».
  // Если такого имени в списке нет — отметка не ставится (а уже поставленная снимается).
  const student = d.calendar.students.find(s => s.name.trim().toLowerCase() === name.toLowerCase() && (s.class || null) === (lesson.class || null) && (s.courseId || null) === (lesson.courseId || null) && (s.classLetterId || null) === (lesson.classLetterId || null));
  if (!student) {
    if (entry.markId) {
      d.calendar.marks = d.calendar.marks.filter(m => m.id !== entry.markId);
      entry.markId = null;
    }
    return;
  }
  const markDate = lesson.actualDate || lesson.plannedDate || todayIso();
  const reason = itemGradeReasonText(d, item) || lesson.topic || "";
  const value = Number(gradeRaw) || gradeRaw;
  const scale = entry.gradeScale || "5";
  const excused = value === "Н" ? entry.gradeExcused !== false : null;
  let mark = entry.markId ? d.calendar.marks.find(m => m.id === entry.markId) : null;
  const patch = {
    studentId: student.id,
    value,
    date: markDate,
    reason,
    scale,
    excused,
    lessonId: lesson.id,
    exerciseId: item.type === "exercise" ? item.refId : null,
    questionId: item.type === "question" ? item.questionId : null
  };
  if (mark) Object.assign(mark, patch);else {
    mark = {
      id: uid(),
      ...patch
    };
    d.calendar.marks.push(mark);
    entry.markId = mark.id;
  }
}
function GradeAssignRow({
  data,
  update,
  lesson,
  item
}) {
  const students = itemStudents(item);
  const [open, setOpen] = useState(students.length > 0);
  const [justAddedId, setJustAddedId] = useState(null);
  const settings = settingsForLesson(data, lesson);
  const scales = scalesInUse(settings.gradingSystem);
  const roster = data.calendar.students.filter(s => lesson.courseId ? s.courseId === lesson.courseId : s.class === lesson.class && !s.courseId && (s.classLetterId || null) === (lesson.classLetterId || null));
  const rosterScore = s => numericMarksOf(data, s.id).length - studentMarksOf(data, s.id).filter(m => m.value === "Н").length;
  const takenNames = new Set(students.map(s => (s.name || "").trim().toLowerCase()));
  const flags = studentCallFlags(data, roster, settings.gradingSystem === "100" ? "100" : "5", settings.requiredMarks);
  const recommended = [...roster].filter(s => !s.noCall && !takenNames.has(s.name.trim().toLowerCase())).sort((a, b) => rosterScore(a) - rosterScore(b)).slice(0, 3);
  const rosterNames = new Set(roster.map(r => r.name.trim().toLowerCase()));
  const patchEntry = (entryId, patch) => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    const it = l.items.find(x => x.id === item.id);
    const entry = itemStudents(it).find(s => s.id === entryId);
    if (!entry) return d;
    Object.assign(entry, patch);
    syncStudentEntryMark(d, l, it, entry);
    return d;
  });
  const addEntry = name => {
    const newId = uid();
    update(d => {
      const l = d.lessons.find(x => x.id === lesson.id);
      const it = l.items.find(x => x.id === item.id);
      it.students = [...itemStudents(it), {
        id: newId,
        name: name || "",
        grade: "",
        gradeScale: "5",
        gradeExcused: null,
        markId: null
      }];
      return d;
    });
    if (!name) setJustAddedId(newId);
  };
  const removeEntry = entryId => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    const it = l.items.find(x => x.id === item.id);
    const entry = itemStudents(it).find(s => s.id === entryId);
    if (entry && entry.markId) d.calendar.marks = d.calendar.marks.filter(m => m.id !== entry.markId);
    it.students = itemStudents(it).filter(s => s.id !== entryId);
    return d;
  });
  if (!open) return /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => {
      setOpen(true);
      if (students.length === 0) addEntry();
    },
    title: "\u041D\u0430\u0437\u043D\u0430\u0447\u0438\u0442\u044C \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432 \u0438 \u043E\u0446\u0435\u043D\u043A\u0438"
  }, "\uD83C\uDF93");
  return /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6
    }
  }, students.map(entry => /*#__PURE__*/React.createElement("div", {
    key: entry.id,
    className: "row wrap",
    style: {
      marginBottom: 6,
      gap: 6
    }
  }, /*#__PURE__*/React.createElement(StudentNameInput, {
    roster: roster,
    flags: flags,
    autoFocus: entry.id === justAddedId,
    onBlur: () => setJustAddedId(null),
    value: entry.name,
    placeholder: "\u0423\u0447\u0435\u043D\u0438\u043A\u2026",
    onChange: v => patchEntry(entry.id, {
      name: v
    }),
    style: {
      border: "1px solid " + ((entry.name || "").trim() && !rosterNames.has(entry.name.trim().toLowerCase()) ? "var(--neg)" : "var(--line)"),
      borderRadius: 8,
      padding: "5px 8px",
      minWidth: 140
    }
  }), (entry.name || "").trim() && !rosterNames.has(entry.name.trim().toLowerCase()) && /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      color: "var(--neg)"
    }
  }, "\u043D\u0435\u0442 \u0432 \u0441\u043F\u0438\u0441\u043A\u0435 \u043A\u043B\u0430\u0441\u0441\u0430 \u2014 \u043E\u0442\u043C\u0435\u0442\u043A\u0430 \u043D\u0435 \u0441\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u0441\u044F"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    value: entry.grade === "Н" ? "" : entry.grade,
    onChange: e => patchEntry(entry.id, {
      grade: e.target.value
    }),
    placeholder: "\u041E\u0446\u0435\u043D\u043A\u0430",
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "5px 8px",
      width: 76
    }
  }), scales.length > 1 && /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: entry.gradeScale,
    onChange: e => patchEntry(entry.id, {
      gradeScale: e.target.value
    })
  }, /*#__PURE__*/React.createElement("option", {
    value: "5"
  }, "5-\u0431"), /*#__PURE__*/React.createElement("option", {
    value: "100"
  }, "100-\u0431")), /*#__PURE__*/React.createElement(Btn, {
    kind: entry.grade === "Н" ? "gold" : "ghost",
    sm: true,
    onClick: () => patchEntry(entry.id, entry.grade === "Н" ? {
      grade: "",
      gradeExcused: null
    } : {
      grade: "Н",
      gradeExcused: true
    })
  }, "\u041D"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => removeEntry(entry.id)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  })))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => addEntry()
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 11
  }), " \u0415\u0449\u0451 \u0443\u0447\u0435\u043D\u0438\u043A"), recommended.map(s => {
    const f = flags.get(s.id);
    return /*#__PURE__*/React.createElement("span", {
      key: s.id,
      style: {
        display: "inline-flex",
        ...(callFlagStyle(f) || {})
      }
    }, /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => addEntry(s.name),
      title: "Рекомендуем — меньше отметок" + callFlagTitle(f)
    }, s.name));
  }), recommended.some(s => callFlagStyle(flags.get(s.id))) && /*#__PURE__*/React.createElement(CallFlagLegend, null), students.length === 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setOpen(false)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  }))));
}
function MaterialPeekModal({
  data,
  type,
  refId,
  onClose
}) {
  const obj = (data[ITEM_TYPE_META[type].dataKey] || []).find(e => e.id === refId);
  if (!obj) return null;
  if (type === "exercise" && obj.typeTagId === TROPE_WALK_TYPE_ID) return /*#__PURE__*/React.createElement(TropeWalkViewer, {
    data: data,
    ex: obj,
    onClose: onClose
  });
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, obj.title || ITEM_TYPE_META[type].label), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, type === "exercise" ? /*#__PURE__*/React.createElement("div", {
    dangerouslySetInnerHTML: {
      __html: obj.text || "<i>Пусто</i>"
    }
  }) : /*#__PURE__*/React.createElement("div", {
    dangerouslySetInnerHTML: {
      __html: obj.course?.lectureHtml || "<i>Пока пусто</i>"
    }
  }))));
}

// Кнопка с выпадающим списком. items: [{ key, label, hint?, sep? }]; onPick(key) вызывается при выборе пункта.
function PopMenu({
  label,
  kind = "ghost",
  sm,
  items,
  onPick,
  title,
  selectedKey
}) {
  const [open, setOpen] = useState(false);
  const ref = useRef(null);
  useEffect(() => {
    if (!open) return undefined;
    const h = e => {
      if (ref.current && !ref.current.contains(e.target)) setOpen(false);
    };
    const k = e => {
      if (e.key === "Escape") setOpen(false);
    };
    document.addEventListener("mousedown", h);
    document.addEventListener("keydown", k);
    return () => {
      document.removeEventListener("mousedown", h);
      document.removeEventListener("keydown", k);
    };
  }, [open]);
  return /*#__PURE__*/React.createElement("span", {
    ref: ref,
    style: {
      position: "relative",
      display: "inline-block"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: kind,
    sm: sm,
    title: title,
    onClick: () => setOpen(v => !v)
  }, label, " \u25BE"), open && /*#__PURE__*/React.createElement("div", {
    className: "pm-menu"
  }, items.map(it => it.sep ? /*#__PURE__*/React.createElement("div", {
    key: it.key,
    className: "pm-sep"
  }) : /*#__PURE__*/React.createElement("div", {
    key: it.key,
    className: "pm-item" + (selectedKey === it.key ? " sel" : ""),
    onClick: () => {
      setOpen(false);
      onPick(it.key);
    }
  }, /*#__PURE__*/React.createElement("span", null, it.label), it.hint ? /*#__PURE__*/React.createElement("span", {
    className: "pm-hint"
  }, it.hint) : null))));
}
const PICKER_SECTIONS = [{
  type: "exercise",
  label: "Упражнения",
  icon: "📝"
}, {
  type: "theory",
  label: "История и теория",
  icon: "📚"
}, {
  type: "personality",
  label: "Персоналии",
  icon: "👤"
}, {
  type: "work",
  label: "Произведения",
  icon: "📖"
}];
const NEW_MATERIAL_LABEL = {
  exercise: "Упражнение",
  theory: "История / теория",
  personality: "Персоналия",
  work: "Произведение"
};
// Текст, по которому ищем материал: для упражнений — весь текст, примечание, вид, название; для остальных — название, автор, метки.
function materialSearchText(data, type, e, tagNames) {
  const tags = (e.tagIds || []).map(id => tagNames.get(id) || "").join(" ");
  if (type === "exercise") {
    const t = (data.tags.exerciseType || []).find(x => x.id === e.typeTagId);
    const a = e.authorId ? (data.personalities.find(p => p.id === e.authorId) || {}).title || "" : "";
    return (stripHtml(e.text || "") + " " + (e.note || "") + " " + (t ? t.name : "") + " " + (e.workTitle || "") + " " + a + " " + tags).toLowerCase();
  }
  const a = type === "work" && e.authorId ? (data.personalities.find(p => p.id === e.authorId) || {}).title || "" : "";
  return ((e.title || "") + " " + a + " " + tags).toLowerCase();
}

// «Выбрать материалы»: выпадающий список разделов (упражнения, теория, персоналии, произведения) → список материалов раздела
// с поиском; либо поиск сразу по всем разделам. extra — дополнительные кнопки в той же строке (например, «Создать материал»).
function ItemPicker({
  data,
  lesson,
  onAddQuestion,
  onAddQuestions,
  onRemoveItems,
  onAddWhole,
  onAddExercise,
  onRemoveItem,
  extra
}) {
  const [type, setType] = useState(null); // null — свёрнуто; "all" — поиск по всем разделам; иначе тип материала
  const [q, setQ] = useState("");
  const [exType, setExType] = useState("");
  const [expandedKey, setExpandedKey] = useState(null);
  const [showHidden, setShowHidden] = useState(false);
  const findItem = pred => lesson.items.find(pred);
  const lessonCtx = {
    class: lesson.class,
    courseId: lesson.courseId,
    classLetterId: lesson.classLetterId || null
  };
  const ctxLabel = scopeLabelOf(data, scopeKeyOf(lessonCtx));
  const tagNames = new Map();
  Object.values(data.tags || {}).forEach(arr => Array.isArray(arr) && arr.forEach(t => t && t.id && tagNames.set(t.id, t.name || "")));
  const inLesson = (t, e) => !!findItem(it => it.type === t && it.refId === e.id || it.type === "question" && it.entityType === t && it.entityId === e.id);
  const hiddenReason = (t, e) => inLesson(t, e) ? null : materialHiddenReason(data, t, e, lessonCtx, lesson.id);
  const choose = k => {
    setType(k);
    setQ("");
    setExType("");
    setExpandedKey(null);
    setShowHidden(false);
  };
  const counts = Object.fromEntries(PICKER_SECTIONS.map(sec => [sec.type, (data[ITEM_TYPE_META[sec.type].dataKey] || []).length]));
  const menuItems = [{
    key: "all",
    label: "🔎 Поиск по всем материалам"
  }, {
    key: "sep",
    sep: true
  }, ...PICKER_SECTIONS.map(sec => ({
    key: sec.type,
    label: sec.icon + " " + sec.label,
    hint: String(counts[sec.type])
  }))];
  const buttons = /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement(PopMenu, {
    kind: "primary",
    label: "\uFF0B \u0412\u044B\u0431\u0440\u0430\u0442\u044C \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B",
    items: menuItems,
    onPick: choose,
    selectedKey: type,
    title: "\u0412\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u0440\u0430\u0437\u0434\u0435\u043B \u2014 \u043E\u0442\u043A\u0440\u043E\u0435\u0442\u0441\u044F \u0441\u043F\u0438\u0441\u043E\u043A \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u043E\u0432, \u0438\u043B\u0438 \u043D\u0430\u0439\u0434\u0438\u0442\u0435 \u043D\u0443\u0436\u043D\u043E\u0435 \u043F\u043E\u0438\u0441\u043A\u043E\u043C"
  }), extra);
  if (!type) return buttons;
  const query = q.trim().toLowerCase();
  const searching = query.length > 0;
  const sections = type === "all" ? PICKER_SECTIONS.map(sec => sec.type) : [type];
  let rows = [];
  let hiddenCount = 0;
  sections.forEach(t => {
    (data[ITEM_TYPE_META[t].dataKey] || []).forEach(e => {
      if (type === "all" && !searching) return;
      if (t === "exercise" && type !== "all" && exType && e.typeTagId !== exType) return;
      if (searching && !materialSearchText(data, t, e, tagNames).includes(query)) return;
      const hr = hiddenReason(t, e);
      if (hr && !searching && type !== "all") {
        hiddenCount++;
        if (!showHidden) return;
      }
      rows.push({
        t,
        e,
        hr
      });
    });
  });
  const total = rows.length;
  const LIMIT = 40;
  rows = rows.slice(0, LIMIT);
  const secLabel = type === "all" ? "Все материалы" : (PICKER_SECTIONS.find(x => x.type === type) || {}).label;
  const exTypes = (data.tags.exerciseType || []).filter(t => (data.examples || []).some(e => e.typeTagId === t.id));
  const renderRow = ({
    t,
    e,
    hr
  }) => {
    const meta = ITEM_TYPE_META[t];
    const key = t + ":" + e.id;
    const exerciseItem = t === "exercise" ? findItem(it => it.type === "exercise" && it.refId === e.id) : null;
    const wholeItem = t !== "exercise" ? findItem(it => it.type === t && it.refId === e.id) : null;
    const qs = t !== "exercise" ? e.course?.questions || [] : [];
    const typeName = t === "exercise" ? ((data.tags.exerciseType || []).find(x => x.id === e.typeTagId) || {}).name : "";
    return /*#__PURE__*/React.createElement("div", {
      key: key
    }, /*#__PURE__*/React.createElement("div", {
      className: "row",
      style: {
        padding: "5px 0",
        borderBottom: "1px solid var(--line)",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        minWidth: 0,
        fontSize: 13,
        cursor: t === "exercise" ? "default" : "pointer"
      },
      onClick: () => t !== "exercise" && setExpandedKey(expandedKey === key ? null : key)
    }, type === "all" && /*#__PURE__*/React.createElement("span", {
      className: "chip",
      style: {
        marginRight: 6,
        padding: "1px 7px",
        fontSize: 11
      }
    }, NEW_MATERIAL_LABEL[t]), typeName && /*#__PURE__*/React.createElement("span", {
      className: "chip",
      style: {
        marginRight: 6,
        padding: "1px 7px",
        fontSize: 11
      }
    }, typeName), meta.titleOf(data, e), hr && /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, " \xB7 ", hiddenReasonText(hr, ctxLabel)), t !== "exercise" && qs.length > 0 && /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, " \xB7 \u0432\u043E\u043F\u0440\u043E\u0441\u043E\u0432: ", qs.length)), t === "exercise" ? /*#__PURE__*/React.createElement(Btn, {
      kind: exerciseItem ? "primary" : "ghost",
      sm: true,
      title: exerciseItem ? "Убрать из урока" : "Добавить в урок",
      onClick: () => exerciseItem ? onRemoveItem(exerciseItem.id) : onAddExercise(e.id)
    }, exerciseItem ? "✓ в уроке" : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Icon, {
      name: "plus",
      size: 12
    }), " \u0434\u043E\u0431\u0430\u0432\u0438\u0442\u044C")) : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Btn, {
      kind: wholeItem ? "primary" : "ghost",
      sm: true,
      title: wholeItem ? "Убрать карточку из урока" : "Добавить карточку целиком (без отдельных вопросов)",
      onClick: () => wholeItem ? onRemoveItem(wholeItem.id) : onAddWhole(t, e.id)
    }, wholeItem ? "✓ целиком" : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Icon, {
      name: "plus",
      size: 12
    }), " \u0446\u0435\u043B\u0438\u043A\u043E\u043C")), /*#__PURE__*/React.createElement("span", {
      style: {
        cursor: "pointer",
        display: "inline-flex",
        transform: expandedKey === key ? "rotate(90deg)" : undefined
      },
      onClick: () => setExpandedKey(expandedKey === key ? null : key),
      title: "\u0412\u044B\u0431\u0440\u0430\u0442\u044C \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0435 \u0432\u043E\u043F\u0440\u043E\u0441\u044B \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "chev",
      size: 13
    })))), t !== "exercise" && expandedKey === key && /*#__PURE__*/React.createElement("div", {
      style: {
        padding: "4px 0 8px 10px",
        borderBottom: "1px solid var(--line)"
      }
    }, qs.length === 0 && /*#__PURE__*/React.createElement("div", {
      className: "hint"
    }, "\u0423 \u044D\u0442\u043E\u0439 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 \u043F\u043E\u043A\u0430 \u043D\u0435\u0442 \u0432\u043E\u043F\u0440\u043E\u0441\u043E\u0432 \u0432 \xAB\u0425\u043E\u0434\u0435 \u0443\u0440\u043E\u043A\u0430\xBB."), qs.length > 0 && (() => {
      const all = [...qs].sort((a, b) => a.order - b.order);
      const picked = all.map(qq => findItem(it => it.type === "question" && it.questionId === qq.id)).filter(Boolean);
      return /*#__PURE__*/React.createElement("div", {
        className: "row wrap",
        style: {
          gap: 8,
          margin: "2px 0 6px",
          alignItems: "center"
        }
      }, /*#__PURE__*/React.createElement(Btn, {
        kind: "primary",
        sm: true,
        disabled: picked.length === all.length,
        onClick: () => onAddQuestions(t, e.id, all.map(qq => qq.id))
      }, "\u2611 \u0412\u044B\u0431\u0440\u0430\u0442\u044C \u0432\u0441\u0435 (", all.length, ")"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        disabled: picked.length === 0,
        onClick: () => onRemoveItems(picked.map(it => it.id))
      }, "\u0421\u043D\u044F\u0442\u044C \u0432\u0441\u0435"), /*#__PURE__*/React.createElement("span", {
        className: "hint"
      }, "\u0412\u044B\u0431\u0440\u0430\u043D\u043E: ", picked.length, " \u0438\u0437 ", all.length, " \u2014 \u0438\u043B\u0438 \u043E\u0442\u043C\u0435\u0447\u0430\u0439\u0442\u0435 \u043D\u0443\u0436\u043D\u044B\u0435 \u043A\u043D\u043E\u043F\u043A\u043E\u0439 \u0441\u043F\u0440\u0430\u0432\u0430 \u043E\u0442 \u0432\u043E\u043F\u0440\u043E\u0441\u0430"));
    })(), [...qs].sort((a, b) => a.order - b.order).map(qq => {
      const qItem = findItem(it => it.type === "question" && it.questionId === qq.id);
      return /*#__PURE__*/React.createElement("div", {
        key: qq.id,
        className: "row",
        style: {
          padding: "3px 0"
        }
      }, /*#__PURE__*/React.createElement("div", {
        style: {
          flex: 1,
          fontSize: 12.5
        }
      }, qq.text), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        title: qItem ? "Убрать из урока" : "Добавить в урок",
        onClick: () => qItem ? onRemoveItem(qItem.id) : onAddQuestion(t, e.id, qq.id)
      }, /*#__PURE__*/React.createElement(Icon, {
        name: qItem ? "check" : "plus",
        size: 11
      })));
    })));
  };
  return /*#__PURE__*/React.createElement("div", null, buttons, /*#__PURE__*/React.createElement("div", {
    className: "addbox",
    style: {
      flexDirection: "column",
      alignItems: "stretch",
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement(PopMenu, {
    label: "Раздел: " + secLabel,
    sm: true,
    items: menuItems,
    onPick: choose,
    selectedKey: type
  }), /*#__PURE__*/React.createElement("input", {
    value: q,
    onChange: e => setQ(e.target.value),
    autoFocus: true,
    placeholder: type === "all" ? "Поиск по упражнениям, теории, персоналиям и произведениям…" : "Поиск в разделе «" + secLabel + "»…",
    style: {
      flex: 1,
      minWidth: 180
    }
  }), type === "exercise" && exTypes.length > 0 && /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: exType,
    onChange: e => setExType(e.target.value),
    title: "\u041F\u043E\u043A\u0430\u0437\u0430\u0442\u044C \u0442\u043E\u043B\u044C\u043A\u043E \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u043E\u0433\u043E \u0432\u0438\u0434\u0430"
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0412\u0441\u0435 \u0432\u0438\u0434\u044B"), exTypes.map(t => /*#__PURE__*/React.createElement("option", {
    key: t.id,
    value: t.id
  }, t.name))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setType(null),
    title: "\u0421\u0432\u0435\u0440\u043D\u0443\u0442\u044C \u0432\u044B\u0431\u043E\u0440"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 12
  }), " \u0421\u0432\u0435\u0440\u043D\u0443\u0442\u044C")), hiddenCount > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      marginTop: 6,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041D\u0435 \u043F\u043E\u043A\u0430\u0437\u0430\u043D\u044B (", ctxLabel, "): ", hiddenCount, " \u2014 \u0443\u0436\u0435 \u0438\u0441\u043F\u043E\u043B\u044C\u0437\u043E\u0432\u0430\u043D\u044B \u0432 \u0443\u0440\u043E\u043A\u0430\u0445, \u043E\u0442\u043A\u043B\u044E\u0447\u0435\u043D\u044B \u0432\u0440\u0443\u0447\u043D\u0443\u044E \u0438\u043B\u0438 \u043D\u0430\u0437\u043D\u0430\u0447\u0435\u043D\u044B \u0434\u0440\u0443\u0433\u0438\u043C \u043A\u043B\u0430\u0441\u0441\u0430\u043C. \u041F\u0440\u0438 \u043F\u043E\u0438\u0441\u043A\u0435 \u043E\u043D\u0438 \u043F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u044E\u0442\u0441\u044F."), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setShowHidden(v => !v)
  }, showHidden ? "Скрыть их" : "Показать их")), /*#__PURE__*/React.createElement("div", {
    style: {
      maxHeight: 320,
      overflow: "auto",
      marginTop: 8
    }
  }, type === "all" && !searching && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0412\u0432\u0435\u0434\u0438\u0442\u0435 \u0441\u043B\u043E\u0432\u043E \u2014 \u043F\u043E\u0438\u0441\u043A \u0438\u0434\u0451\u0442 \u0441\u0440\u0430\u0437\u0443 \u043F\u043E \u0432\u0441\u0435\u043C \u0440\u0430\u0437\u0434\u0435\u043B\u0430\u043C (\u0442\u0435\u043A\u0441\u0442 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439, \u043D\u0430\u0437\u0432\u0430\u043D\u0438\u044F, \u0430\u0432\u0442\u043E\u0440\u044B, \u043C\u0435\u0442\u043A\u0438)."), !(type === "all" && !searching) && total === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E."), rows.map(renderRow), total > LIMIT && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      padding: "6px 0"
    }
  }, "\u041F\u043E\u043A\u0430\u0437\u0430\u043D\u043E ", LIMIT, " \u0438\u0437 ", total, " \u2014 \u0443\u0442\u043E\u0447\u043D\u0438\u0442\u0435 \u043F\u043E\u0438\u0441\u043A."))));
}

// Быстрая карточка для персоналии / произведения / истории-теории: название, (вид / автор), классы, текст.
// Создаётся в своём разделе («Персоналии», «Произведения», «История и теория»), а в урок встаёт целиком.
function EntityQuickForm({
  data,
  update,
  projectPath,
  type,
  editId,
  onDone
}) {
  const meta = ITEM_TYPE_META[type];
  const key = meta.dataKey;
  const existing = editId ? (data[key] || []).find(e => e.id === editId) : null;
  const [title, setTitle] = useState(existing ? existing.title || "" : "");
  const [kind, setKind] = useState(existing && existing.kind || "theory");
  const [authorId, setAuthorId] = useState(existing && existing.authorId || "");
  const [tagIds, setTagIds] = useState(existing && existing.tagIds || []);
  const [html, setHtml] = useState(existing && existing.course && existing.course.lectureHtml || "");
  const placeholder = type === "personality" ? "Имя автора" : type === "work" ? "Название произведения" : "Название темы";
  const save = () => {
    if (!title.trim()) {
      showToast("Введите название");
      return;
    }
    let createdId = null;
    update(d => {
      const arr = d[key];
      if (existing) {
        const i = arr.findIndex(e => e.id === editId);
        if (i >= 0) arr[i] = {
          ...arr[i],
          title: title.trim(),
          ...(type === "theory" ? {
            kind
          } : {}),
          ...(type === "work" ? {
            authorId: authorId || null
          } : {}),
          tagIds,
          course: {
            ...arr[i].course,
            lectureHtml: html
          },
          updatedAt: nowISO()
        };
      } else {
        const e = migrateEntity({
          title: title.trim(),
          kind: type === "theory" ? kind : null,
          authorId: type === "work" ? authorId || null : null,
          tagIds,
          course: {
            lectureHtml: html
          }
        }, {});
        arr.push(e);
        createdId = e.id;
      }
      return d;
    });
    onDone(createdId);
  };
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("input", {
    className: "titleinput",
    value: title,
    placeholder: placeholder,
    autoFocus: !existing,
    onChange: e => setTitle(e.target.value)
  }), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      margin: "8px 0 12px",
      gap: 12,
      alignItems: "center"
    }
  }, type === "theory" && /*#__PURE__*/React.createElement("div", {
    className: "seg"
  }, (data.tags.historyTheoryKind || []).map(k => /*#__PURE__*/React.createElement("button", {
    key: k.id,
    "aria-pressed": kind === k.id ? "true" : undefined,
    onClick: () => setKind(k.id)
  }, k.name))), type === "work" && /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: authorId,
    onChange: e => setAuthorId(e.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0410\u0432\u0442\u043E\u0440 \u043D\u0435 \u0443\u043A\u0430\u0437\u0430\u043D"), data.personalities.map(p => /*#__PURE__*/React.createElement("option", {
    key: p.id,
    value: p.id
  }, p.title)))), /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.classTag.title + " (для каких классов предлагается при составлении уроков; пусто — всем)"
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "classTag",
    selectedIds: nsTagIds(tagIds, "classTag", data),
    onChange: ids => setTagIds(withNsTagIds(tagIds, "classTag", data, ids))
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u0422\u0435\u043A\u0441\u0442 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438"
  }, /*#__PURE__*/React.createElement(RichEditor, {
    html: html,
    onChange: setHtml,
    projectPath: projectPath,
    compact: true,
    update: update,
    sourceInfo: {
      section: meta.label,
      entityId: editId || null,
      entityTitle: title,
      authorId: type === "personality" ? editId || null : type === "work" ? authorId || null : null,
      workId: type === "work" ? editId || null : null,
      genreTagIds: []
    }
  })), /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041E\u0441\u0442\u0430\u043B\u044C\u043D\u044B\u0435 \u043F\u043E\u043B\u044F (\u0436\u0430\u043D\u0440\u044B, \u0446\u0432\u0435\u0442, \u0434\u043E\u0441\u0442\u0443\u043F\u043D\u043E\u0441\u0442\u044C) \u043C\u043E\u0436\u043D\u043E \u043D\u0430\u0441\u0442\u0440\u043E\u0438\u0442\u044C \u043F\u043E\u0437\u0436\u0435 \u0432 \u0440\u0430\u0437\u0434\u0435\u043B\u0435 \xAB", type === "personality" ? "Персоналии" : type === "work" ? "Произведения" : "История и теория", "\xBB. \u0412\u043E\u043F\u0440\u043E\u0441\u044B \u043A \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0435 \u2014 \u0442\u0430\u043C \u0436\u0435, \u0432 \xAB\u0425\u043E\u0434\u0435 \u0443\u0440\u043E\u043A\u0430\xBB, \u043B\u0438\u0431\u043E \u043F\u0440\u044F\u043C\u043E \u0432 \u0443\u0440\u043E\u043A\u0435."), /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      justifyContent: "flex-end",
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => onDone(null)
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: save
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " ", existing ? "Сохранить" : "Создать и добавить в урок")));
}

// Окно «Создать материал в уроке» / «Править материал»: сразу выбирается тип, затем вводится и редактируется текст.
// spec: { mode: "create" | "edit", type, id? }. Новый материал сохраняется в своём разделе и добавляется в урок.
function LessonMaterialModal({
  data,
  update,
  projectPath,
  lesson,
  spec,
  onCreated,
  onClose
}) {
  const isEdit = spec.mode === "edit";
  const type = spec.type;
  const types = data.tags.exerciseType || [];
  const existingEx = isEdit && type === "exercise" ? data.examples.find(e => e.id === spec.id) : null;
  const [typeId, setTypeId] = useState(() => {
    if (existingEx) return existingEx.typeTagId;
    const t = types.find(x => x.id !== TROPE_WALK_TYPE_ID) || types[0];
    return t ? t.id : "";
  });
  const pickType = async v => {
    if (v !== "__new") {
      setTypeId(v);
      return;
    }
    const name = await PromptStr("Название нового вида упражнения:");
    if (!name || !name.trim()) return;
    const old = types.find(t => t.name.toLowerCase() === name.trim().toLowerCase());
    if (old) {
      setTypeId(old.id);
      return;
    }
    const id = uid();
    update(d => {
      d.tags.exerciseType.push({
        id,
        name: name.trim(),
        color: SWATCHES[Math.floor(Math.random() * SWATCHES.length)],
        categoryId: null,
        order: d.tags.exerciseType.length,
        slots: []
      });
      return d;
    });
    setTypeId(id);
  };
  const done = newId => {
    if (newId && !isEdit) onCreated(type, newId);
    onClose();
  };
  const where = type === "exercise" ? "в банке упражнений" : "в разделе «" + (type === "personality" ? "Персоналии" : type === "work" ? "Произведения" : "История и теория") + "»";
  const lessonName = (lesson.topic || "").trim() ? "«" + lesson.topic.trim() + "»" : "этот урок";
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    style: {
      zIndex: 200
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 940
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, isEdit ? "Править: " : "Новый материал — ", NEW_MATERIAL_LABEL[type]), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose,
    title: "\u0417\u0430\u043A\u0440\u044B\u0442\u044C \u0431\u0435\u0437 \u0441\u043E\u0445\u0440\u0430\u043D\u0435\u043D\u0438\u044F"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, isEdit ? "Изменения сохраняются в самой карточке (" + where + ") и будут видны во всех уроках, где она используется." : "Материал сохранится " + where + " и сразу добавится в " + lessonName + "."), type === "exercise" ? /*#__PURE__*/React.createElement(React.Fragment, null, !isEdit && /*#__PURE__*/React.createElement(Field, {
    label: "\u0412\u0438\u0434 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: typeId,
    onChange: e => pickType(e.target.value)
  }, !typeId && /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u2014 \u0432\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u0432\u0438\u0434 \u2014"), types.map(t => /*#__PURE__*/React.createElement("option", {
    key: t.id,
    value: t.id
  }, t.name)), /*#__PURE__*/React.createElement("option", {
    value: "__new"
  }, "\uFF0B \u041D\u043E\u0432\u044B\u0439 \u0432\u0438\u0434\u2026"))), typeId ? /*#__PURE__*/React.createElement(ExampleForm, {
    key: typeId + (spec.id || ""),
    data: data,
    update: update,
    projectPath: projectPath,
    typeId: typeId,
    editId: isEdit ? spec.id : null,
    onDone: done
  }) : /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0412\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u0432\u0438\u0434 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F (\u0438\u043B\u0438 \u0441\u043E\u0437\u0434\u0430\u0439\u0442\u0435 \u043D\u043E\u0432\u044B\u0439), \u0447\u0442\u043E\u0431\u044B \u0432\u0432\u0435\u0441\u0442\u0438 \u0442\u0435\u043A\u0441\u0442.")) : /*#__PURE__*/React.createElement(EntityQuickForm, {
    data: data,
    update: update,
    projectPath: projectPath,
    type: type,
    editId: isEdit ? spec.id : null,
    onDone: done
  }))));
}

// Независимая копия существующего урока в выбранный класс/литеру (или курс) — материалы и ДЗ
// клонируются, отметки/статус «пройден»/дата фактически — нет, копия начинает свою историю.
// Уроки одной «группы» (класс+литера, либо курс) — литеры считаются независимо друг от друга.
function sameLessonGroup(a, b) {
  if (a.courseId || b.courseId) return a.courseId === b.courseId;
  return a.class === b.class && (a.classLetterId || null) === (b.classLetterId || null);
}
// Переставляет элементы с датой в порядке дат, но только между местами, которые такие элементы уже занимали;
// элементы без даты остаются на своих местах. Так порядок «по датам» не ломает ручную расстановку.
function sortDatedInSlots(items, dateOf) {
  const idxs = [],
    dated = [];
  items.forEach((x, i) => {
    if (dateOf(x)) {
      idxs.push(i);
      dated.push({
        x,
        i
      });
    }
  });
  dated.sort((a, b) => dateOf(a.x) < dateOf(b.x) ? -1 : dateOf(a.x) > dateOf(b.x) ? 1 : a.i - b.i);
  const out = items.slice();
  idxs.forEach((pos, k) => {
    out[pos] = dated[k].x;
  });
  return out;
}
// Порядок отдельных уроков по датам. Список «Уроки» — это отдельные уроки и системы вперемешку, расставленные
// учителем как угодно; уроки между двумя системами (подряд идущие отдельные уроки) образуют «серию», и ВНУТРИ
// серии они идут по датам. Поэтому урок можно поставить выше системы, между системами или ниже — а в каждой
// серии порядок определяют даты проведения. Значения order переиспользуются те же (занятые «места»).
function resyncOrderForGroup(d, sample) {
  const entries = topLevelEntries(d, {
    class: sample.class,
    courseId: sample.courseId,
    classLetterId: sample.classLetterId || null
  }).sort((a, b) => a.order - b.order);
  let run = [];
  const flush = () => {
    if (run.length > 1) {
      const lessons = run.map(e => d.lessons.find(x => x.id === e.id)).filter(Boolean);
      const slots = run.map(e => e.order);
      sortDatedInSlots(lessons, l => l.plannedDate).forEach((l, k) => {
        l.order = slots[k];
      });
    }
    run = [];
  };
  entries.forEach(e => {
    if (e.type === "lesson") run.push(e);else flush();
  });
  flush();
}
// Порядок уроков ВНУТРИ системы по датам: уроки с датой идут по возрастанию дат на тех же местах, без даты — как стояли.
function resyncSystemLessonsByDate(d, systemId) {
  const sys = d.lessonSystems.find(s => s.id === systemId);
  if (!sys) return;
  const dateOf = id => {
    const l = d.lessons.find(x => x.id === id);
    return l ? l.plannedDate : "";
  };
  sys.lessonIds = sortDatedInSlots(sys.lessonIds, dateOf);
}
// Вызывать после любого изменения даты урока — держит согласованными и его позицию в общем
// списке (если он отдельный), и порядок внутри системы (если он в неё входит).
function resyncAfterDateChange(d, lesson) {
  resyncOrderForGroup(d, lesson);
  if (lesson.systemId) resyncSystemLessonsByDate(d, lesson.systemId);
}
// День после последнего назначенного урока этой же литеры/курса (пусто, если таких ещё нет).
function nextLessonDateFor(d, sample) {
  const dates = d.lessons.filter(l => sameLessonGroup(l, sample)).map(l => l.plannedDate).filter(Boolean).sort();
  return dates.length ? addDaysIso(dates[dates.length - 1], 1) : "";
}
// Следующее свободное значение order в ОБЩЕМ пространстве позиций списка "Уроки" — отдельные
// уроки и системы уроков делят одну числовую последовательность, чтобы их можно было
// перемешивать (перетаскивать урок перед/после системы) — см. reorderTopLevel.
function nextTopLevelOrder(d, sample) {
  const lessonOrders = d.lessons.filter(l => sameLessonGroup(l, sample) && !l.systemId).map(l => l.order);
  const systemOrders = d.lessonSystems.filter(s => sample.courseId ? s.courseId === sample.courseId : s.class === sample.class).map(s => s.order || 0);
  const all = [...lessonOrders, ...systemOrders];
  return all.length ? Math.max(...all) + 1 : 0;
}
// Верхнеуровневые "блоки" списка "Уроки" для группы (класс/литера либо курс) — отдельные уроки
// и системы уроков вперемешку, отсортированные по общему order. Системы не привязаны к литере
// (лежат в d.lessonSystems с одним classSlot на весь класс/курс), поэтому при просмотре разных
// литер один и тот же набор систем комбинируется с уроками именно этой литеры.
function topLevelEntries(d, sample) {
  // sample.anyLetter — режим «Все литеры»: берём уроки и системы класса целиком.
  const inGroup = l => sample.anyLetter ? sample.courseId ? l.courseId === sample.courseId : l.class === sample.class && !l.courseId : sameLessonGroup(l, sample);
  const standalone = d.lessons.filter(l => inGroup(l) && !l.systemId);
  const systems = d.lessonSystems.filter(s => (sample.courseId ? s.courseId === sample.courseId : s.class === sample.class) && (sample.anyLetter || !s.classLetterId || s.classLetterId === (sample.classLetterId || null)));
  return [...standalone.map(l => ({
    type: "lesson",
    id: l.id,
    order: l.order
  })), ...systems.map(s => ({
    type: "system",
    id: s.id,
    order: s.order || 0
  }))];
}
// Переставляет один верхнеуровневый блок (урок или систему) на новое место среди ВСЕХ блоков
// группы — так урок можно перетащить не только среди других уроков, но и перед/после системы,
// а систему — среди уроков и других систем. targetEntry === null — переместить в конец списка.
function reorderTopLevel(d, sample, draggedEntry, targetEntry) {
  const entries = topLevelEntries(d, sample).sort((a, b) => a.order - b.order);
  const fromIdx = entries.findIndex(e => e.type === draggedEntry.type && e.id === draggedEntry.id);
  if (fromIdx === -1) return;
  const [moved] = entries.splice(fromIdx, 1);
  let toIdx = targetEntry ? entries.findIndex(e => e.type === targetEntry.type && e.id === targetEntry.id) : -1;
  if (toIdx === -1) toIdx = entries.length;
  entries.splice(toIdx, 0, moved);
  entries.forEach((e, idx) => {
    if (e.type === "lesson") {
      const l = d.lessons.find(x => x.id === e.id);
      if (l) l.order = idx;
    } else {
      const s = d.lessonSystems.find(x => x.id === e.id);
      if (s) s.order = idx;
    }
  });
}
// Создаёт один урок с автоматической датой (если plannedDate не передана) и сразу выстраивает порядок.
function createLessonAuto(d, fields) {
  const sample = {
    class: fields.class ?? null,
    courseId: fields.courseId ?? null,
    classLetterId: fields.classLetterId ?? null
  };
  const lesson = migrateLesson({
    ...fields,
    order: nextTopLevelOrder(d, sample),
    plannedDate: fields.plannedDate !== undefined ? fields.plannedDate : nextLessonDateFor(d, sample)
  });
  d.lessons.push(lesson);
  resyncOrderForGroup(d, lesson);
  return lesson;
}
// Независимая копия урока в другую литеру; дата — следующий день после последнего урока той литеры.
function duplicateLessonInto(d, srcId, {
  classLetterId,
  intoSystemId
}) {
  const src = d.lessons.find(x => x.id === srcId);
  if (!src) return null;
  if (!src.copyGroupId) src.copyGroupId = uid();
  const target = {
    class: src.class,
    courseId: null,
    classLetterId: classLetterId || null
  };
  const clone = structuredClone(src);
  const copy = {
    ...clone,
    id: uid(),
    classLetterId: classLetterId || null,
    order: nextTopLevelOrder(d, target),
    plannedDate: nextLessonDateFor(d, target),
    completed: false,
    archived: false,
    actualDate: "",
    needsWork: false,
    // Копия всегда полностью автономна — даже если исходный урок был частью системы, копия
    // становится отдельным уроком: иначе она получила бы systemId чужой (не своей) системы,
    // не попав в её lessonIds, и была бы невидима в списке "Уроки" ни как отдельный урок, ни
    // внутри какой-либо системы.
    systemId: intoSystemId || null,
    items: (clone.items || []).map(it => ({
      ...it,
      id: uid(),
      students: []
    }))
  };
  d.lessons.push(copy);
  resyncOrderForGroup(d, copy);
  return copy.id;
}

// Дубликат урока «на продолжение»: копия (тема, материалы, ДЗ — без отметок учеников) встаёт на дату СЛЕДУЮЩЕГО по календарю
// урока этой литеры/класса, а следующие уроки сдвигаются по цепочке на даты друг друга (последний — на 7 дней позже своей даты).
// Проведённые уроки и уроки с отметками журнала не двигаются. Без следующего урока — на следующий день после оригинала.
function duplicateLessonInSequence(d, srcId) {
  const src = d.lessons.find(x => x.id === srcId);
  if (!src) return null;
  const sample = {
    class: src.class,
    courseId: src.courseId,
    classLetterId: src.classLetterId || null
  };
  const hasMark = l => d.calendar.marks.some(m => m.lessonId === l.id);
  let dupDate = "",
    moved = 0;
  const movedLessons = [];
  if (src.plannedDate) {
    const following = d.lessons.filter(l => l.id !== src.id && !l.archived && !l.excludeFromSchedule && !l.completed && sameLessonGroup(l, src) && l.plannedDate && l.plannedDate > src.plannedDate && !hasMark(l)).sort((a, b) => a.plannedDate < b.plannedDate ? -1 : a.plannedDate > b.plannedDate ? 1 : a.order - b.order);
    if (following.length) {
      const dates = following.map(l => l.plannedDate);
      dupDate = dates[0];
      following.forEach((l, i) => {
        l.plannedDate = i + 1 < dates.length ? dates[i + 1] : addDaysIso(dates[i], 7);
        movedLessons.push(l);
      });
      moved = following.length;
    } else dupDate = addDaysIso(src.plannedDate, 1);
  }
  const clone = structuredClone(src);
  const dup = migrateLesson({
    ...clone,
    id: uid(),
    copyGroupId: null,
    duplicateOf: src.id,
    plannedDate: dupDate,
    actualDate: "",
    completed: false,
    archived: false,
    needsWork: false,
    items: (clone.items || []).map(it => ({
      ...it,
      id: uid(),
      students: []
    })),
    createdAt: nowISO(),
    updatedAt: nowISO()
  });
  d.lessons.push(dup);
  if (src.systemId) {
    const sys = d.lessonSystems.find(x => x.id === src.systemId);
    dup.systemId = src.systemId;
    if (sys) {
      const i = sys.lessonIds.indexOf(src.id);
      sys.lessonIds.splice(i < 0 ? sys.lessonIds.length : i + 1, 0, dup.id);
    }
  } else {
    dup.systemId = null;
    const entries = topLevelEntries(d, sample).filter(e => !(e.type === "lesson" && e.id === dup.id)).sort((a, b) => a.order - b.order);
    const idx = entries.findIndex(e => e.type === "lesson" && e.id === src.id);
    reorderTopLevel(d, sample, {
      type: "lesson",
      id: dup.id
    }, idx >= 0 ? entries[idx + 1] || null : null);
  }
  resyncOrderForGroup(d, src);
  new Set([src.systemId, ...movedLessons.map(l => l.systemId)].filter(Boolean)).forEach(sid => resyncSystemLessonsByDate(d, sid));
  return {
    dup,
    dupDate,
    moved
  };
}

// ---- Экспорт/импорт ОДНОГО урока в отдельный файл .lsnlit (JSON внутри) — не путать с общим
// файлом проекта. Урок уходит со всем содержимым (упражнения, карточки персоналий/произведений/
// теории и выбранные вопросы встраиваются целиком), а при импорте сопоставляется по НАЗВАНИЯМ:
// найденное в программе не перезаписывается, недостающее создаётся — данные не теряются.
// Не переносятся: даты, отметки/ученики, прикреплённые файлы и картинки (лежат рядом с проектом). ----
function tagRefsOf(data, ids) {
  const out = [];
  (ids || []).forEach(id => {
    for (const ns of Object.keys(data.tags)) {
      if (ns === "historyTheoryKind" || ns === "exerciseType") continue;
      const t = (data.tags[ns] || []).find(x => x.id === id);
      if (t) {
        out.push({
          ns,
          name: t.name
        });
        break;
      }
    }
  });
  return out;
}
function tagIdsFromRefs(d, refs) {
  return (refs || []).map(r => Array.isArray(d.tags[r.ns]) ? ensureTag(d, r.ns, r.name) : null).filter(Boolean);
}
function entityExportData(data, type, e) {
  const author = type === "work" && e.authorId ? data.personalities.find(p => p.id === e.authorId) : null;
  return {
    title: e.title || "",
    kind: e.kind || null,
    color: e.color || null,
    authorTitle: author ? author.title : "",
    tags: tagRefsOf(data, e.tagIds),
    lectureHtml: e.course?.lectureHtml || "",
    questions: (e.course?.questions || []).map(q => ({
      text: q.text,
      order: q.order ?? 0,
      questionTypeName: (data.tags.questionType.find(t => t.id === q.questionTypeTagId) || {}).name || ""
    }))
  };
}
function exampleExportData(data, ex) {
  const type = data.tags.exerciseType.find(t => t.id === ex.typeTagId);
  const author = ex.authorId ? data.personalities.find(p => p.id === ex.authorId) : null;
  return {
    typeId: type ? type.id : null,
    typeName: type ? type.name : "",
    text: ex.text || "",
    note: ex.note || "",
    questionTypeName: (data.tags.questionType.find(t => t.id === ex.questionTypeTagId) || {}).name || "",
    tags: tagRefsOf(data, ex.tagIds),
    authorTitle: author ? author.title : "",
    workTitle: ex.workTitle || "",
    questions: (ex.questions || []).map(q => ({
      text: q.text,
      order: q.order ?? 0,
      questionTypeName: (data.tags.questionType.find(t => t.id === q.questionTypeTagId) || {}).name || ""
    }))
  };
}
function lessonExportPayload(data, lesson) {
  const items = [];
  const entities = {};
  const addEntity = (type, ent) => {
    const key = type + ":" + ent.id;
    if (!entities[key]) entities[key] = entityExportData(data, type, ent);
    return key;
  };
  (lesson.items || []).forEach(it => {
    if (it.type === "exercise") {
      const ex = data.examples.find(e => e.id === it.refId);
      if (ex) items.push({
        type: "exercise",
        note: it.note || "",
        ref: ex.id,
        exercise: exampleExportData(data, ex)
      });
    } else if (it.type === "question") {
      const ent = questionSourceEntity(data, it);
      const q = questionOfItem(data, it);
      const typeName = q ? (data.tags.questionType.find(t => t.id === q.questionTypeTagId) || {}).name || "" : "";
      if (ent && q) items.push({
        type: "question",
        entityType: it.entityType,
        note: it.note || "",
        entityKey: it.entityType === "exercise" ? "exercise:" + ent.id : addEntity(it.entityType, ent),
        questionText: q.text,
        local: !!q.local,
        questionTypeName: typeName
      });
    } else {
      const ent = itemRefObj(data, it);
      if (ent) items.push({
        type: it.type,
        note: it.note || "",
        entityKey: addEntity(it.type, ent)
      });
    }
  });
  return {
    type: "literature-constructor-lesson",
    app: "literature",
    version: 2,
    exportedAt: new Date().toISOString(),
    topic: lesson.topic || "",
    lessonTypeNames: (lesson.lessonTypeTagIds || []).map(id => (data.tags.lessonType.find(t => t.id === id) || {}).name).filter(Boolean),
    homework: lesson.homework || "",
    entities,
    items,
    groupNotes: Object.keys(lesson.groupNotes || {}).filter(k => entities[k]).map(k => ({
      entityKey: k,
      text: lesson.groupNotes[k]
    }))
  };
}
function findOrCreateEntityFromExport(d, type, ex) {
  const key = ITEM_TYPE_META[type].dataKey;
  const want = (ex.title || "").trim().toLowerCase();
  let e = d[key].find(x => (x.title || "").trim().toLowerCase() === want);
  if (!e) {
    const authorId = type === "work" && ex.authorTitle ? ensurePersonality(d, ex.authorTitle) : null;
    e = migrateEntity({
      title: ex.title || "",
      kind: type === "theory" ? ex.kind === "history" ? "history" : "theory" : null,
      color: ex.color || null,
      authorId,
      tagIds: tagIdsFromRefs(d, ex.tags),
      course: {
        lectureHtml: ex.lectureHtml || "",
        attachments: [],
        notes: [],
        questions: (ex.questions || []).map(q => ({
          id: uid(),
          text: q.text,
          order: q.order ?? 0,
          questionTypeTagId: q.questionTypeName ? ensureTag(d, "questionType", q.questionTypeName) : null
        }))
      }
    }, {});
    d[key].push(e);
  }
  return e;
}
function findOrCreateExampleFromExport(d, ex) {
  let type = ex.typeId === TROPE_WALK_TYPE_ID ? d.tags.exerciseType.find(t => t.id === TROPE_WALK_TYPE_ID) : null;
  if (!type && ex.typeName) type = d.tags.exerciseType.find(t => (t.name || "").trim().toLowerCase() === ex.typeName.trim().toLowerCase());
  if (!type) {
    type = {
      id: uid(),
      name: ex.typeName || "Импортированные",
      color: SWATCHES[Math.floor(Math.random() * SWATCHES.length)],
      categoryId: null,
      order: d.tags.exerciseType.length,
      slots: [],
      locked: false
    };
    d.tags.exerciseType.push(type);
  }
  const authorId = ex.authorTitle ? ensurePersonality(d, ex.authorTitle) : null;
  const exQuestions = (ex.questions || []).map(q => ({
    id: uid(),
    text: q.text,
    order: q.order ?? 0,
    questionTypeTagId: q.questionTypeName ? ensureTag(d, "questionType", q.questionTypeName) : null
  }));
  const same = d.examples.find(e => e.typeTagId === type.id && e.text === (ex.text || "") && (e.workTitle || "") === (ex.workTitle || "") && (e.authorId || null) === authorId);
  if (same) return same;
  const created = migrateExample({
    questions: exQuestions,
    typeTagId: type.id,
    text: ex.text || "",
    note: ex.note || "",
    tagIds: tagIdsFromRefs(d, ex.tags),
    authorId,
    workTitle: ex.workTitle || "",
    questionTypeTagId: ex.questionTypeName ? ensureTag(d, "questionType", ex.questionTypeName) : null
  });
  d.examples.push(created);
  return created;
}
// Создаёт новый отдельный урок в указанной группе из данных файла .lsnlit (вызывать внутри update).
function importLessonPayload(d, payload, group) {
  const lesson = createLessonAuto(d, {
    class: group.kind === "class" ? group.value : null,
    courseId: group.kind === "course" ? group.value : null,
    classLetterId: group.classLetterId || null,
    topic: payload.topic || ""
  });
  lesson.homework = payload.homework || "";
  lesson.groupNotes = lesson.groupNotes || {};
  (payload.groupNotes || []).forEach(gn => {
    const m = String(gn.entityKey || "").split(":");
    if (!ITEM_TYPE_META[m[0]] || !(payload.entities || {})[gn.entityKey]) return;
    const ent = findOrCreateEntityFromExport(d, m[0], payload.entities[gn.entityKey]);
    if (gn.text) lesson.groupNotes[m[0] + ":" + ent.id] = gn.text;
  });
  lesson.lessonTypeTagIds = (payload.lessonTypeNames || []).filter((n, i, a) => a.indexOf(n) === i).map(n => ensureTag(d, "lessonType", n)).filter(Boolean);
  lesson.items = [];
  const exByRef = {};
  (payload.items || []).forEach(it => {
    if (it.type === "exercise" && it.exercise) {
      const ex = findOrCreateExampleFromExport(d, it.exercise);
      if (it.ref) exByRef[it.ref] = ex;
      lesson.items.push(migrateLessonItem({
        type: "exercise",
        refId: ex.id,
        note: it.note
      }));
    } else if (it.type === "question" && it.entityType === "exercise") {
      const ex = exByRef[String(it.entityKey || "").split(":")[1]];
      if (!ex) return;
      const typeId = it.questionTypeName ? ensureTag(d, "questionType", it.questionTypeName) : null;
      if (it.local) {
        lesson.items.push(migrateLessonItem({
          type: "question",
          entityType: "exercise",
          entityId: ex.id,
          questionId: uid(),
          note: it.note,
          local: {
            text: it.questionText,
            questionTypeTagId: typeId
          }
        }));
        return;
      }
      ex.questions = ex.questions || [];
      let q = ex.questions.find(x => x.text === it.questionText);
      if (!q) {
        q = {
          id: uid(),
          text: it.questionText,
          order: ex.questions.length,
          questionTypeTagId: typeId
        };
        ex.questions.push(q);
      }
      lesson.items.push(migrateLessonItem({
        type: "question",
        entityType: "exercise",
        entityId: ex.id,
        questionId: q.id,
        note: it.note
      }));
    } else if (it.type === "question" && ITEM_TYPE_META[it.entityType] && (payload.entities || {})[it.entityKey]) {
      const entData = payload.entities[it.entityKey];
      const ent = findOrCreateEntityFromExport(d, it.entityType, entData);
      if (it.local) {
        lesson.items.push(migrateLessonItem({
          type: "question",
          entityType: it.entityType,
          entityId: ent.id,
          questionId: uid(),
          note: it.note,
          local: {
            text: it.questionText,
            questionTypeTagId: it.questionTypeName ? ensureTag(d, "questionType", it.questionTypeName) : null
          }
        }));
        return;
      }
      ent.course.questions = ent.course.questions || [];
      let q = ent.course.questions.find(x => x.text === it.questionText);
      if (!q) {
        const src = (entData.questions || []).find(x => x.text === it.questionText);
        q = {
          id: uid(),
          text: it.questionText,
          order: ent.course.questions.length,
          questionTypeTagId: src && src.questionTypeName ? ensureTag(d, "questionType", src.questionTypeName) : null
        };
        ent.course.questions.push(q);
      }
      lesson.items.push(migrateLessonItem({
        type: "question",
        entityType: it.entityType,
        entityId: ent.id,
        questionId: q.id,
        note: it.note
      }));
    } else if (ITEM_TYPE_META[it.type] && it.type !== "exercise" && (payload.entities || {})[it.entityKey]) {
      const ent = findOrCreateEntityFromExport(d, it.type, payload.entities[it.entityKey]);
      lesson.items.push(migrateLessonItem({
        type: it.type,
        refId: ent.id,
        note: it.note
      }));
    }
  });
  return lesson;
}
// ---- Экспорт/импорт ЦЕЛОЙ системы уроков в файл .lsnsys: название, цвет и все её уроки по порядку
// (каждый — как в .lsnlit, карточки персоналий/произведений/теории записаны один раз на всю систему). ----
function systemExportPayload(data, system) {
  const entities = {};
  const lessons = system.lessonIds.map(id => data.lessons.find(l => l.id === id)).filter(Boolean).map(l => {
    const p = lessonExportPayload(data, l);
    Object.assign(entities, p.entities);
    delete p.entities;
    return p;
  });
  const link = system.colorLinkKind && system.colorLinkId ? (data[system.colorLinkKind] || []).find(e => e.id === system.colorLinkId) : null;
  return {
    type: "literature-constructor-system",
    app: "literature",
    version: 1,
    exportedAt: new Date().toISOString(),
    title: system.title || "",
    color: system.color || null,
    collapsed: !!system.collapsed,
    colorLink: link ? {
      kind: system.colorLinkKind,
      title: link.title
    } : null,
    entities,
    lessons
  };
}
// Создаёт систему и все её уроки в указанной группе (вызывать внутри update). Уроки — отдельные копии,
// сопоставление карточек/упражнений по названиям, как у .lsnlit.
function importSystemPayload(d, payload, group) {
  const sample = {
    class: group.kind === "class" ? group.value : null,
    courseId: group.kind === "course" ? group.value : null,
    classLetterId: group.classLetterId || null
  };
  const link = payload.colorLink && ["personalities", "theoryHistory"].includes(payload.colorLink.kind) ? d[payload.colorLink.kind].find(e => (e.title || "").trim().toLowerCase() === (payload.colorLink.title || "").trim().toLowerCase()) : null;
  const sys = migrateLessonSystem({
    title: payload.title || "",
    color: payload.color || undefined,
    collapsed: !!payload.collapsed,
    colorLinkKind: link ? payload.colorLink.kind : null,
    colorLinkId: link ? link.id : null,
    class: sample.class,
    courseId: sample.courseId,
    classLetterId: sample.classLetterId,
    order: nextTopLevelOrder(d, sample)
  });
  d.lessonSystems.push(sys);
  (payload.lessons || []).forEach(lp => {
    const l = importLessonPayload(d, {
      ...lp,
      entities: payload.entities || {}
    }, group);
    l.systemId = sys.id;
    sys.lessonIds.push(l.id);
  });
  return sys;
}
function exportSystemFile(data, system) {
  downloadTextFile(safeFileName(system.title || "sistema") + ".lsnsys", JSON.stringify(systemExportPayload(data, system), null, 1), "application/json");
}
function downloadTextFile(filename, text, mime) {
  const blob = new Blob([text], {
    type: mime || "application/octet-stream"
  });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = filename;
  a.click();
  URL.revokeObjectURL(url);
}
function exportLessonFile(data, lesson) {
  downloadTextFile(safeFileName(lesson.topic || "urok") + ".lsnlit", JSON.stringify(lessonExportPayload(data, lesson), null, 1), "application/json");
}

// Делит items урока на «блоки» в порядке их следования: каждое упражнение — отдельный блок, а карточка
// (персоналия/произведение/теория) вместе с выбранными из неё вопросами — один блок. Блоки можно двигать
// стрелками ↑↓; после перестановки items собираются обратно из блоков (ничего не теряется).
function lessonItemBlocks(items) {
  const blocks = [],
    byKey = {},
    pending = [];
  items.forEach(it => {
    if (it.type === "exercise") {
      blocks.push({
        kind: "exercise",
        key: "ex:" + it.id,
        refId: it.refId,
        items: [it]
      });
      return;
    }
    const isQ = it.type === "question";
    if (isQ && it.entityType === "exercise") {
      pending.push(it);
      return;
    }
    const entityType = isQ ? it.entityType : it.type,
      entityId = isQ ? it.entityId : it.refId;
    const key = entityType + ":" + entityId;
    let b = byKey[key];
    if (!b) {
      b = {
        kind: "group",
        key,
        entityType,
        entityId,
        whole: null,
        questions: [],
        items: []
      };
      byKey[key] = b;
      blocks.push(b);
    }
    if (isQ) b.questions.push(it);else b.whole = it;
  });
  blocks.forEach(b => {
    if (b.kind === "group") b.items = [b.whole, ...b.questions].filter(Boolean);
  });
  // вопросы к упражнению едут вместе с упражнением (если самого упражнения нет в уроке — «потерянный» блок, не показывается)
  pending.forEach(q => {
    const eb = blocks.find(b => b.kind === "exercise" && b.refId === q.entityId);
    if (eb) eb.items.push(q);else blocks.push({
      kind: "orphan",
      key: "orphan:" + q.id,
      items: [q]
    });
  });
  return blocks;
}
function MoveArrows({
  canUp,
  canDown,
  onUp,
  onDown
}) {
  return /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flexDirection: "column",
      gap: 2,
      flex: "none"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !canUp,
    onClick: onUp,
    title: "\u041F\u043E\u0434\u043D\u044F\u0442\u044C \u0432\u044B\u0448\u0435"
  }, "\u2191"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !canDown,
    onClick: onDown,
    title: "\u041E\u043F\u0443\u0441\u0442\u0438\u0442\u044C \u043D\u0438\u0436\u0435"
  }, "\u2193"));
}

// Выбор вида вопроса: избранные виды — кнопками-чипами, все виды — списком, можно создать новый вид и
// пометить вид избранным (★). compact — только список (для строки вопроса).
function QuestionTypeSelect({
  data,
  update,
  value,
  onChange,
  compact
}) {
  const types = data.tags.questionType || [];
  const favs = types.filter(t => t.favorite);
  const cur = types.find(t => t.id === value);
  const toggleFav = () => {
    if (cur) update(d => {
      const t = d.tags.questionType.find(x => x.id === cur.id);
      if (t) t.favorite = !t.favorite;
      return d;
    });
  };
  const createNew = async () => {
    const name = await PromptStr("Название нового вида вопроса:");
    if (!name || !name.trim()) return;
    const found = types.find(t => (t.name || "").trim().toLowerCase() === name.trim().toLowerCase());
    if (found) {
      onChange(found.id);
      return;
    }
    const id = uid();
    update(d => {
      d.tags.questionType.push({
        id,
        name: name.trim(),
        color: SWATCHES[Math.floor(Math.random() * SWATCHES.length)],
        favorite: true
      });
      return d;
    });
    onChange(id);
  };
  const select = /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: value || "",
    onChange: e => {
      const v = e.target.value;
      if (v === "__new__") createNew();else onChange(v || null);
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0412\u0438\u0434 \u0432\u043E\u043F\u0440\u043E\u0441\u0430\u2026"), favs.length > 0 && /*#__PURE__*/React.createElement("optgroup", {
    label: "\u2605 \u0418\u0437\u0431\u0440\u0430\u043D\u043D\u044B\u0435"
  }, favs.map(t => /*#__PURE__*/React.createElement("option", {
    key: "f" + t.id,
    value: t.id
  }, t.name))), /*#__PURE__*/React.createElement("optgroup", {
    label: "\u0412\u0441\u0435 \u0432\u0438\u0434\u044B"
  }, types.map(t => /*#__PURE__*/React.createElement("option", {
    key: t.id,
    value: t.id
  }, t.name))), /*#__PURE__*/React.createElement("option", {
    value: "__new__"
  }, "+ \u0421\u043E\u0437\u0434\u0430\u0442\u044C \u043D\u043E\u0432\u044B\u0439 \u0432\u0438\u0434\u2026"));
  if (compact) return select;
  return /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      alignItems: "center"
    }
  }, favs.map(t => /*#__PURE__*/React.createElement("button", {
    key: t.id,
    type: "button",
    className: "chip",
    onClick: () => onChange(value === t.id ? null : t.id),
    style: {
      cursor: "pointer",
      borderColor: value === t.id ? "var(--accent)" : undefined,
      background: value === t.id ? "var(--accent-soft)" : undefined,
      color: value === t.id ? "var(--accent)" : undefined
    }
  }, "\u2605 ", t.name)), select, cur && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: toggleFav,
    title: cur.favorite ? "Убрать из избранных видов" : "Сделать вид избранным (будет кнопкой)"
  }, cur.favorite ? "★" : "☆"));
}

// Форма «Свой вопрос» на странице урока: вопрос можно оставить только в уроке или сразу включить в карточку элемента.
function LessonAddQuestion({
  data,
  update,
  lesson,
  entityType,
  entityId
}) {
  const [text, setText] = useState("");
  const [typeId, setTypeId] = useState(null);
  const [include, setInclude] = useState(false);
  const label = entityType === "exercise" ? "упражнения" : entityType === "personality" ? "персоналии" : entityType === "work" ? "произведения" : "теории";
  const add = () => {
    const t = text.trim();
    if (!t) return;
    const qid = uid();
    update(d => {
      const meta = ITEM_TYPE_META[entityType];
      const ent = (d[meta.dataKey] || []).find(e => e.id === entityId);
      const l = d.lessons.find(x => x.id === lesson.id);
      if (!ent || !l) return d;
      if (include) {
        if (entityType === "exercise") {
          ent.questions = ent.questions || [];
          ent.questions.push({
            id: qid,
            text: t,
            order: ent.questions.length,
            questionTypeTagId: typeId
          });
        } else {
          ent.course = ent.course || {
            lectureHtml: "",
            attachments: [],
            notes: [],
            questions: []
          };
          ent.course.questions = ent.course.questions || [];
          ent.course.questions.push({
            id: qid,
            text: t,
            order: ent.course.questions.length,
            questionTypeTagId: typeId
          });
        }
        l.items.push(migrateLessonItem({
          type: "question",
          entityType,
          entityId,
          questionId: qid
        }));
      } else {
        l.items.push(migrateLessonItem({
          type: "question",
          entityType,
          entityId,
          questionId: qid,
          local: {
            text: t,
            questionTypeTagId: typeId
          }
        }));
      }
      return d;
    });
    showToast(include ? "Вопрос добавлен в урок и в карточку" : "Вопрос добавлен в урок");
    setText("");
  };
  return /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 8,
      padding: 8,
      border: "1px dashed var(--line)",
      borderRadius: 8,
      background: "var(--accent-soft)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700,
      fontSize: 12.5,
      marginBottom: 4
    }
  }, "\u0421\u0432\u043E\u0439 \u0432\u043E\u043F\u0440\u043E\u0441 \u043A \u044D\u0442\u043E\u043C\u0443 \u044D\u043B\u0435\u043C\u0435\u043D\u0442\u0443"), /*#__PURE__*/React.createElement("input", {
    value: text,
    onChange: e => setText(e.target.value),
    onKeyDown: e => {
      if (e.key === "Enter") add();
    },
    placeholder: "\u0422\u0435\u043A\u0441\u0442 \u0432\u043E\u043F\u0440\u043E\u0441\u0430\u2026",
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      marginBottom: 6
    }
  }), /*#__PURE__*/React.createElement(QuestionTypeSelect, {
    data: data,
    update: update,
    value: typeId,
    onChange: setTypeId
  }), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginTop: 6,
      gap: 10,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 6,
      fontSize: 12.5,
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: include,
    onChange: e => setInclude(e.target.checked)
  }), " \u0412\u043A\u043B\u044E\u0447\u0438\u0442\u044C \u0432\u043E\u043F\u0440\u043E\u0441 \u0432 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0443 ", label), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    disabled: !text.trim(),
    onClick: add
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0432\u043E\u043F\u0440\u043E\u0441")), !include && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 3
    }
  }, "\u0411\u0435\u0437 \u0433\u0430\u043B\u043E\u0447\u043A\u0438 \u0432\u043E\u043F\u0440\u043E\u0441 \u043E\u0441\u0442\u0430\u043D\u0435\u0442\u0441\u044F \u0442\u043E\u043B\u044C\u043A\u043E \u0432 \u044D\u0442\u043E\u043C \u0443\u0440\u043E\u043A\u0435 (\u043F\u043E\u0437\u0436\u0435 \u0435\u0433\u043E \u043C\u043E\u0436\u043D\u043E \u043F\u0435\u0440\u0435\u043D\u0435\u0441\u0442\u0438 \u0432 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0443 \u043A\u043D\u043E\u043F\u043A\u043E\u0439 \xAB\u0412 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0443\xBB)."));
}

// Поле, которое растёт по высоте вместе с текстом — длинный вопрос виден целиком.
function AutoTextarea({
  value,
  onChange,
  style,
  placeholder,
  onBlur,
  autoFocus
}) {
  const ref = useRef(null);
  const fit = () => {
    const el = ref.current;
    if (el) {
      el.style.height = "auto";
      el.style.height = el.scrollHeight + "px";
    }
  };
  useEffect(fit, [value]);
  useEffect(() => {
    window.addEventListener("resize", fit);
    return () => window.removeEventListener("resize", fit);
  }, []);
  return /*#__PURE__*/React.createElement("textarea", {
    ref: ref,
    rows: 1,
    value: value,
    placeholder: placeholder,
    autoFocus: autoFocus,
    onBlur: onBlur,
    onChange: e => onChange(e.target.value),
    style: {
      resize: "none",
      overflow: "hidden",
      ...style
    }
  });
}

// Примечание-«стикер» учителю к элементу урока: виден только в уроке (в карточке самого материала не показывается).
function StickyNote({
  value,
  onChange
}) {
  const [editing, setEditing] = useState(false);
  const has = !!(value && value.trim());
  if (!has && !editing) return /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 4
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setEditing(true),
    title: "\u0417\u0430\u043F\u0438\u0441\u043A\u0430-\u043F\u043E\u0434\u0441\u043A\u0430\u0437\u043A\u0430 \u0434\u043B\u044F \u0432\u0430\u0441: \u0432 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0435 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u0430 \u043D\u0435 \u043F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u0435\u0442\u0441\u044F"
  }, "\uD83D\uDCDD \u041F\u0440\u0438\u043C\u0435\u0447\u0430\u043D\u0438\u0435"));
  return /*#__PURE__*/React.createElement("div", {
    className: "sticky",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement(AutoTextarea, {
    autoFocus: !has,
    value: value || "",
    placeholder: "\u041F\u043E\u0434\u0441\u043A\u0430\u0437\u043A\u0430 \u0441\u0435\u0431\u0435 (\u0432\u0438\u0434\u0435\u043D \u0442\u043E\u043B\u044C\u043A\u043E \u0432 \u0443\u0440\u043E\u043A\u0435)\u2026",
    onChange: onChange,
    onBlur: () => {
      if (!(value && value.trim())) setEditing(false);
    }
  }), /*#__PURE__*/React.createElement("button", {
    type: "button",
    className: "sticky-x",
    title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C \u043F\u0440\u0438\u043C\u0435\u0447\u0430\u043D\u0438\u0435",
    onClick: () => {
      onChange("");
      setEditing(false);
    }
  }, "\u2715"));
}

// Один «свёрнутый» пункт материалов урока на карточку-источник (персоналия/произведение/история-теория/упражнение) —
// объединяет и «карточку целиком», и вопросы к ней. Разворачивается по клику: слева текст, справа вопросы — каждая часть
// листается сама; щелчок по вопросу прокручивает текст к привязанному фрагменту. «На весь экран» — разворот на всё окно.
// embedded — без заголовка (внутри строки упражнения).
function EntityMaterialRow({
  data,
  update,
  lesson,
  entityType,
  entityId,
  questionItemsList,
  hasWhole,
  onRemoveQuestion,
  onRemoveGroup,
  onMoveQuestion,
  embedded,
  onPrint,
  onEdit
}) {
  const [open, setOpen] = useState(!!embedded);
  const [full, setFull] = useState(false);
  const [activeQ, setActiveQ] = useState(null);
  const entity = questionSourceEntity(data, {
    entityType,
    entityId
  });
  const containerRef = useRef(null);
  const savedRangeRef = useRef(null);
  const textHtml = entity ? entityTextHtmlOf(entity, entityType) : "";
  useEffect(() => {
    const c = containerRef.current;
    if (!c) return;
    c.querySelectorAll(".qa-active").forEach(el => el.classList.remove("qa-active"));
    if (activeQ) c.querySelectorAll(`[data-qa-id="${CSS.escape(activeQ)}"]`).forEach(el => el.classList.add("qa-active"));
  });
  useEffect(() => {
    if (!full) return;
    const onKey = e => {
      if (e.key === "Escape") setFull(false);
    };
    window.addEventListener("keydown", onKey);
    return () => window.removeEventListener("keydown", onKey);
  }, [full]);
  if (!entity) return null;
  const dataKey = ITEM_TYPE_META[entityType].dataKey;
  const highlight = (qId, on) => {
    const c = containerRef.current;
    if (!c) return;
    c.querySelectorAll(`[data-qa-id="${CSS.escape(qId)}"]`).forEach(el => el.classList.toggle("qa-highlight", on));
  };
  const focusQuestion = qId => {
    setActiveQ(qId);
    const c = containerRef.current;
    if (!c) return;
    const el = c.querySelector(`[data-qa-id="${CSS.escape(qId)}"]`);
    if (!el) return;
    const cr = c.getBoundingClientRect(),
      er = el.getBoundingClientRect();
    c.scrollTo({
      top: c.scrollTop + (er.top - cr.top) - c.clientHeight / 3,
      behavior: "smooth"
    });
  };
  const onSelectionCapture = () => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount && !sel.isCollapsed) {
      const r = sel.getRangeAt(0);
      if (containerRef.current && containerRef.current.contains(r.commonAncestorContainer)) savedRangeRef.current = r.cloneRange();
    }
  };
  const writeBackText = () => {
    if (!containerRef.current) return;
    const html = containerRef.current.innerHTML;
    update(d => {
      const arr = d[dataKey];
      const idx = arr.findIndex(e => e.id === entityId);
      if (idx < 0) return d;
      if (entityType === "exercise") arr[idx] = {
        ...arr[idx],
        text: html,
        updatedAt: nowISO()
      };else arr[idx] = {
        ...arr[idx],
        course: {
          ...arr[idx].course,
          lectureHtml: html
        },
        updatedAt: nowISO()
      };
      return d;
    });
  };
  const linkFragment = qId => {
    if (!savedRangeRef.current) {
      showToast("Сначала выделите фрагмент текста слева");
      return;
    }
    let span = null;
    try {
      span = wrapRangeInSpan(savedRangeRef.current, "lc-qa-anchor", {
        "data-qa-id": qId,
        "data-anchor-id": uid()
      });
    } catch (e) {
      span = null;
    }
    savedRangeRef.current = null;
    if (!span) return;
    writeBackText();
    setActiveQ(qId);
  };
  const unlinkFragment = (qId, anchorId) => {
    if (!containerRef.current) return;
    const removed = anchorId ? unwrapSpanById(containerRef.current, "data-anchor-id", anchorId) : unwrapSpanById(containerRef.current, "data-qa-id", qId);
    if (removed) writeBackText();
  };
  const setQuestionType = (it, typeId) => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    const item = l && l.items.find(x => x.id === it.id);
    if (!item) return d;
    if (item.local) {
      item.local.questionTypeTagId = typeId;
      return d;
    }
    const ent = (d[dataKey] || []).find(e => e.id === entityId);
    const q = ent && entityQuestionsOf(ent, entityType).find(x => x.id === item.questionId);
    if (q) q.questionTypeTagId = typeId;
    return d;
  });
  const promoteToCard = it => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    const item = l && l.items.find(x => x.id === it.id);
    const ent = (d[dataKey] || []).find(e => e.id === entityId);
    if (!item || !item.local || !ent) return d;
    const q = {
      id: item.questionId,
      text: item.local.text,
      order: 0,
      questionTypeTagId: item.local.questionTypeTagId ?? null
    };
    if (entityType === "exercise") {
      ent.questions = ent.questions || [];
      q.order = ent.questions.length;
      ent.questions.push(q);
    } else {
      ent.course = ent.course || {
        lectureHtml: "",
        attachments: [],
        notes: [],
        questions: []
      };
      ent.course.questions = ent.course.questions || [];
      q.order = ent.course.questions.length;
      ent.course.questions.push(q);
    }
    item.local = null;
    return d;
  });
  const setItemNote = (itemId, v) => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    const item = l && l.items.find(x => x.id === itemId);
    if (item) item.note = v;
    return d;
  });
  const groupKey = entityType + ":" + entityId;
  const setGroupNote = v => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    if (!l) return d;
    l.groupNotes = l.groupNotes || {};
    if (v && v.trim()) l.groupNotes[groupKey] = v;else delete l.groupNotes[groupKey];
    return d;
  });
  const meta = ITEM_TYPE_META[entityType];
  const paneH = full ? "calc(100vh - 70px)" : embedded ? "min(64vh, 600px)" : "calc(100vh - 150px)";
  const body = /*#__PURE__*/React.createElement("div", {
    className: "split-pane",
    style: {
      height: paneH,
      marginTop: full ? 0 : 8
    }
  }, /*#__PURE__*/React.createElement("div", {
    ref: containerRef,
    className: "editor-area",
    style: {
      cursor: "text",
      height: "100%"
    },
    onMouseUp: onSelectionCapture,
    dangerouslySetInnerHTML: {
      __html: textHtml || "<p><i>Текста пока нет</i></p>"
    }
  }), /*#__PURE__*/React.createElement("div", {
    className: "split-right"
  }, questionItemsList.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0412\u043E\u043F\u0440\u043E\u0441\u043E\u0432 \u043F\u043E\u043A\u0430 \u043D\u0435\u0442 \u2014 \u0434\u043E\u0431\u0430\u0432\u044C\u0442\u0435 \u0441\u0432\u043E\u0439 \u043D\u0438\u0436\u0435", entityType === "exercise" ? "." : " или выберите из карточки через форму выбора («" + meta.label + "» → раскрыть карточку)."), questionItemsList.map((it, qi) => {
    const q = questionOfItem(data, it);
    const fragments = q ? qaFragmentsOf(textHtml, it.questionId) : [];
    const isActive = activeQ === it.questionId;
    return /*#__PURE__*/React.createElement("div", {
      key: it.id,
      className: "qrow",
      style: {
        cursor: "pointer",
        borderColor: isActive ? "var(--accent)" : undefined,
        boxShadow: isActive ? "0 0 0 2px var(--accent-soft2)" : undefined
      },
      onMouseEnter: () => highlight(it.questionId, true),
      onMouseLeave: () => highlight(it.questionId, false),
      onClick: e => {
        if (e.target.closest && e.target.closest("button,select,input,textarea,a,label,.x")) return;
        focusQuestion(it.questionId);
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row",
      style: {
        alignItems: "flex-start"
      }
    }, onMoveQuestion && questionItemsList.length > 1 && /*#__PURE__*/React.createElement(MoveArrows, {
      canUp: qi > 0,
      canDown: qi < questionItemsList.length - 1,
      onUp: () => onMoveQuestion(it.id, -1),
      onDown: () => onMoveQuestion(it.id, 1)
    }), /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        minWidth: 0,
        fontWeight: 600,
        fontSize: 13.5,
        whiteSpace: "pre-wrap",
        overflowWrap: "anywhere"
      }
    }, q ? q.text : "(вопрос удалён)", q && q.local && /*#__PURE__*/React.createElement("span", {
      className: "chip",
      style: {
        marginLeft: 6,
        padding: "1px 7px",
        fontSize: 11
      },
      title: "\u042D\u0442\u043E\u0433\u043E \u0432\u043E\u043F\u0440\u043E\u0441\u0430 \u043D\u0435\u0442 \u0432 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0435 \u2014 \u043E\u043D \u0442\u043E\u043B\u044C\u043A\u043E \u0432 \u044D\u0442\u043E\u043C \u0443\u0440\u043E\u043A\u0435"
    }, "\u0442\u043E\u043B\u044C\u043A\u043E \u0432 \u0443\u0440\u043E\u043A\u0435")), q && q.local && /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => promoteToCard(it),
      title: "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u044D\u0442\u043E\u0442 \u0432\u043E\u043F\u0440\u043E\u0441 \u0432 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0443 (\u043E\u043D \u043F\u043E\u044F\u0432\u0438\u0442\u0441\u044F \u0438 \u0432 \u0434\u0440\u0443\u0433\u0438\u0445 \u0443\u0440\u043E\u043A\u0430\u0445)"
    }, "\u0412 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0443"), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => onRemoveQuestion(it.id)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "trash",
      size: 12
    }))), /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        marginTop: 5,
        gap: 6
      }
    }, q && /*#__PURE__*/React.createElement(QuestionTypeSelect, {
      compact: true,
      data: data,
      update: update,
      value: q.questionTypeTagId,
      onChange: v => setQuestionType(it, v)
    }), /*#__PURE__*/React.createElement(Btn, {
      kind: fragments.length > 0 ? "gold" : "ghost",
      sm: true,
      onClick: () => linkFragment(it.questionId),
      title: "\u0412\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0444\u0440\u0430\u0433\u043C\u0435\u043D\u0442 \u0442\u0435\u043A\u0441\u0442\u0430 \u0441\u043B\u0435\u0432\u0430 \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435"
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "link",
      size: 11
    }), " ", fragments.length > 0 ? `Привязано: ${fragments.length}` : "Привязать выделенный текст"), fragments.map((f, fi) => /*#__PURE__*/React.createElement("span", {
      key: f.anchorId || fi,
      className: "chip",
      style: {
        fontSize: 11
      },
      title: f.text
    }, f.text.slice(0, 26), f.text.length > 26 ? "…" : "", /*#__PURE__*/React.createElement("span", {
      className: "x",
      onClick: () => unlinkFragment(it.questionId, f.anchorId)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "close",
      size: 10
    }))))), /*#__PURE__*/React.createElement(StickyNote, {
      value: it.note,
      onChange: v => setItemNote(it.id, v)
    }), /*#__PURE__*/React.createElement(GradeAssignRow, {
      data: data,
      update: update,
      lesson: lesson,
      item: it
    }));
  }), /*#__PURE__*/React.createElement(LessonAddQuestion, {
    data: data,
    update: update,
    lesson: lesson,
    entityType: entityType,
    entityId: entityId
  })));
  const fullBtn = /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      setFull(v => !v);
    },
    title: "\u0422\u0435\u043A\u0441\u0442 \u0441\u043B\u0435\u0432\u0430, \u0432\u043E\u043F\u0440\u043E\u0441\u044B \u0441\u043F\u0440\u0430\u0432\u0430 \u2014 \u043D\u0430 \u0432\u0435\u0441\u044C \u044D\u043A\u0440\u0430\u043D (Esc \u2014 \u0441\u0432\u0435\u0440\u043D\u0443\u0442\u044C)"
  }, full ? "Свернуть" : "⤢ На весь экран");
  if (full) {
    return /*#__PURE__*/React.createElement("div", {
      style: {
        position: "fixed",
        inset: 0,
        zIndex: 900,
        background: "var(--bg)",
        padding: 10,
        display: "flex",
        flexDirection: "column",
        gap: 8
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row",
      style: {
        flex: "none"
      }
    }, /*#__PURE__*/React.createElement("span", {
      className: "chip"
    }, meta.label), /*#__PURE__*/React.createElement("b", {
      style: {
        flex: 1,
        fontSize: 15
      }
    }, meta.titleOf(data, entity)), fullBtn), body);
  }
  if (embedded) return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      justifyContent: "flex-end",
      marginTop: 4
    }
  }, fullBtn), body);
  return /*#__PURE__*/React.createElement("div", {
    className: "qrow"
  }, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      cursor: "pointer"
    },
    onClick: () => setOpen(v => !v)
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      display: "inline-flex",
      transform: open ? "rotate(90deg)" : undefined,
      transition: "transform .12s",
      flex: "none"
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chev",
    size: 13
  })), /*#__PURE__*/React.createElement("span", {
    className: "chip",
    style: {
      flex: "none"
    }
  }, meta.label), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      fontSize: 13,
      fontWeight: 600
    }
  }, meta.titleOf(data, entity)), questionItemsList.length > 0 && /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, questionItemsList.length, " \u0432\u043E\u043F\u0440."), open && fullBtn, onEdit && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onEdit();
    },
    title: "\u041F\u0440\u0430\u0432\u0438\u0442\u044C \u043D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u0438 \u0442\u0435\u043A\u0441\u0442 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 (\u0438\u0437\u043C\u0435\u043D\u0438\u0442\u0441\u044F \u0432 \u0440\u0430\u0437\u0434\u0435\u043B\u0435 \u043A\u0430\u0440\u0442\u043E\u0447\u0435\u043A)"
  }, "\u270E"), onPrint && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onPrint();
    },
    title: "\u0420\u0430\u0441\u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C (\u0442\u0435\u043A\u0441\u0442, \u0432\u043E\u043F\u0440\u043E\u0441\u044B \u0443\u0440\u043E\u043A\u0430, \u043F\u0440\u0438\u043C\u0435\u0447\u0430\u043D\u0438\u044F)"
  }, "\uD83D\uDDA8"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onRemoveGroup();
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 12
  }))), /*#__PURE__*/React.createElement(StickyNote, {
    value: (lesson.groupNotes || {})[groupKey],
    onChange: setGroupNote
  }), open && body);
}

// Под строкой упражнения: раскрывающийся блок «Вопросы» (текст упражнения, вопросы, привязка к тексту, добавление вопроса).
function ExerciseQuestionsBlock({
  data,
  update,
  lesson,
  exItem,
  questionItems,
  onRemoveQuestion,
  onMoveQuestion
}) {
  const [open, setOpen] = useState(false);
  return /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: questionItems.length > 0 ? "gold" : "ghost",
    sm: true,
    onClick: () => setOpen(v => !v)
  }, open ? "▾" : "▸", " \u0412\u043E\u043F\u0440\u043E\u0441\u044B", questionItems.length > 0 ? ` (${questionItems.length})` : "", " \u0438 \u043F\u0440\u0438\u0432\u044F\u0437\u043A\u0430 \u043A \u0442\u0435\u043A\u0441\u0442\u0443"), open && /*#__PURE__*/React.createElement(EntityMaterialRow, {
    embedded: true,
    data: data,
    update: update,
    lesson: lesson,
    entityType: "exercise",
    entityId: exItem.refId,
    questionItemsList: questionItems,
    hasWhole: false,
    onRemoveQuestion: onRemoveQuestion,
    onRemoveGroup: () => {},
    onMoveQuestion: onMoveQuestion
  }));
}

// Окно «Не отображать в …»: выбор классов/литер/курсов, где выбранные стихотворения больше не предлагаются.
// ---- Печать материала (из карточки материала и из карточки урока) ----
// Текст карточки приводится к печатному виду: без служебных кнопок, привязок к вопросам и плееров (вместо аудио/видео — название файла).
function cleanPrintHtml(html) {
  if (!html) return "";
  const doc = new DOMParser().parseFromString(html, "text/html");
  doc.querySelectorAll(".blk-del").forEach(el => el.remove());
  doc.querySelectorAll("audio,video").forEach(el => {
    const blk = el.closest(".lc-block");
    const nm = blk && blk.querySelector(".blk-name") ? blk.querySelector(".blk-name").textContent : "";
    const div = doc.createElement("div");
    div.textContent = "▶ " + (nm || (el.tagName === "AUDIO" ? "аудиозапись" : "видео"));
    el.replaceWith(div);
  });
  doc.querySelectorAll("a").forEach(a => {
    const sp = doc.createElement("span");
    sp.textContent = a.textContent;
    a.replaceWith(sp);
  });
  doc.querySelectorAll(".lc-qa-anchor,.lc-note").forEach(el => {
    const p = el.parentNode;
    while (el.firstChild) p.insertBefore(el.firstChild, el);
    p.removeChild(el);
  });
  doc.querySelectorAll(".lc-block").forEach(el => {
    el.style.border = "none";
    el.style.padding = "0";
    el.style.background = "none";
  });
  doc.querySelectorAll("img").forEach(im => {
    im.style.maxWidth = "100%";
    im.style.height = "auto";
  });
  return doc.body.innerHTML;
}
// o: { text, questions, answers (ответы из текста для учителя), space (место для ответа), sticky (примечания-стикеры), size (px) }
function buildMaterialPrintHtml(data, type, ent, o, lesson) {
  const meta = ITEM_TYPE_META[type];
  const esc = escapeHtml;
  const title = meta.titleOf(data, ent);
  const textHtml = entityTextHtmlOf(ent, type);
  const poem = type === "exercise" && ent.typeTagId === TROPE_WALK_TYPE_ID;
  let sub = "";
  if (type === "work" && ent.authorId) {
    const a = data.personalities.find(p => p.id === ent.authorId);
    if (a) sub = a.title;
  }
  if (type === "theory") {
    const k = data.tags.historyTheoryKind.find(x => x.id === ent.kind);
    if (k) sub = k.name;
  }
  let h = `<div style="font-family:Arial,sans-serif;color:#000;font-size:${o.size}px;line-height:1.5">`;
  h += `<h1 style="font-size:${Math.round(o.size * 1.6)}px;margin:0 0 4px;${poem ? "text-align:center;" : ""}">${esc(title)}</h1>`;
  if (sub) h += `<div style="color:#444;margin-bottom:10px;${poem ? "text-align:center;" : ""}">${esc(sub)}</div>`;
  const stickyOf = txt => `<div style="background:#FFF1A8;border:1px solid #E6D470;padding:6px 10px;margin:6px 0;font-size:${Math.max(11, o.size - 2)}px;">📌 ${esc(txt)}</div>`;
  if (o.sticky && lesson) {
    const exItem = type === "exercise" ? lesson.items.find(it => it.type === "exercise" && it.refId === ent.id) : null;
    const gn = type === "exercise" ? exItem && exItem.note : (lesson.groupNotes || {})[type + ":" + ent.id];
    if (gn && gn.trim()) h += stickyOf(gn);
  }
  if (o.text) {
    const body = cleanPrintHtml(textHtml) || "<p><i>Текста нет</i></p>";
    h += poem ? `<div style="text-align:center;font-weight:bold;font-size:${Math.round(o.size * 1.3)}px;line-height:1.6;margin:10px 0 18px;">${body}</div>` : `<div style="margin:8px 0 16px;">${body}</div>`;
  }
  if (o.questions) {
    let qs;
    if (lesson) qs = lesson.items.filter(it => it.type === "question" && it.entityType === type && it.entityId === ent.id).map(it => ({
      q: questionOfItem(data, it),
      note: it.note
    })).filter(x => x.q);else qs = [...entityQuestionsOf(ent, type)].sort((a, b) => (a.order ?? 0) - (b.order ?? 0)).map(q => ({
      q,
      note: ""
    }));
    if (qs.length) {
      h += `<h2 style="font-size:${Math.round(o.size * 1.2)}px;margin:14px 0 6px;">Вопросы</h2>`;
      qs.forEach((x, i) => {
        h += `<div style="margin-bottom:14px;page-break-inside:avoid;"><b>${i + 1}. ${esc(x.q.text)}</b>`;
        if (o.answers) {
          const frag = qaFragmentsOf(textHtml, x.q.id).map(f => f.text).join(" … ");
          if (frag) h += `<div style="color:#132c53;margin-top:3px;">Ответ (из текста): ${esc(frag)}</div>`;
        }
        if (o.space) h += `<div style="border-bottom:1px solid #999;height:${Math.round(o.size * 2.6)}px;"></div><div style="border-bottom:1px solid #999;height:${Math.round(o.size * 2.2)}px;"></div>`;
        if (o.sticky && x.note && x.note.trim()) h += stickyOf(x.note);
        h += `</div>`;
      });
    }
  }
  return h + `</div>`;
}
function PrintMaterialModal({
  data,
  type,
  entityId,
  lesson,
  onClose
}) {
  const ent = (data[ITEM_TYPE_META[type].dataKey] || []).find(e => e.id === entityId);
  const [o, setO] = useState({
    text: true,
    questions: true,
    answers: false,
    space: false,
    sticky: false
  });
  const [copies, setCopies] = useState(1);
  const [fontPx, setFontPx] = useState(16);
  const [overflow, setOverflow] = useState(false);
  if (!ent) return null;
  const set = patch => setO(p => ({
    ...p,
    ...patch
  }));
  const title = ITEM_TYPE_META[type].titleOf(data, ent);
  const nQ = lesson ? lesson.items.filter(it => it.type === "question" && it.entityType === type && it.entityId === entityId).length : entityQuestionsOf(ent, type).length;
  const cb = (k, label, disabled) => /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 6,
      padding: "3px 0",
      opacity: disabled ? 0.5 : 1,
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!o[k],
    disabled: disabled,
    onChange: e => set({
      [k]: e.target.checked
    })
  }), " ", label);
  const content = buildMaterialPrintHtml(data, type, ent, {
    ...o,
    size: fontPx
  }, lesson);
  const doPrint = () => {
    const h = cardsPageHtml(content, copies, fontPx, false);
    onClose();
    setTimeout(() => printCardsPdf(h), 60);
  };
  const doWord = async () => {
    const h = await inlineLitfileImages(cardsPageHtml(content, copies, fontPx, false));
    downloadWord(safeFileName(title) + ".doc", title, h, {
      noTitle: true,
      pageCss: CARDS_WORD_PAGE_CSS
    });
    onClose();
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      width: "min(860px, 96vw)"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, "\u041F\u0435\u0447\u0430\u0442\u044C \u0438 \u0441\u043A\u0430\u0447\u0438\u0432\u0430\u043D\u0438\u0435: ", title), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      gap: 18,
      alignItems: "flex-start",
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: "1 1 280px",
      minWidth: 250
    }
  }, /*#__PURE__*/React.createElement(CardsOptions, {
    copies: copies,
    setCopies: setCopies,
    fontPx: fontPx,
    setFontPx: setFontPx
  }), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontWeight: 700,
      margin: "12px 0 4px"
    }
  }, "\u0427\u0442\u043E \u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C"), cb("text", "Текст материала"), cb("questions", (lesson ? "Вопросы этого урока" : "Вопросы карточки") + " (" + nQ + ")", nQ === 0), cb("answers", "Ответы из текста (привязанные фрагменты) — для учителя", !o.questions || nQ === 0), cb("space", "Оставить место для ответов — для учеников", !o.questions || nQ === 0), lesson && cb("sticky", "Примечания-стикеры"), overflow && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12.5,
      color: "var(--neg)",
      background: "var(--neg-soft)",
      borderRadius: 8,
      padding: "8px 10px",
      marginTop: 8
    }
  }, "\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B \u043D\u0435 \u043F\u043E\u043C\u0435\u0449\u0430\u0435\u0442\u0441\u044F \u0432 \u043E\u0434\u043D\u0443 \u044F\u0447\u0435\u0439\u043A\u0443: \u0447\u0430\u0441\u0442\u044C \u0442\u0435\u043A\u0441\u0442\u0430 \u043E\u0431\u0440\u0435\u0436\u0435\u0442\u0441\u044F. \u0412\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u043C\u0435\u043D\u044C\u0448\u0435 \u043A\u043E\u043F\u0438\u0439 \u043D\u0430 \u043B\u0438\u0441\u0442\u0435 \u0438\u043B\u0438 \u043C\u0435\u043D\u044C\u0448\u0438\u0439 \u0440\u0430\u0437\u043C\u0435\u0440 \u0442\u0435\u043A\u0441\u0442\u0430."), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 8
    }
  }, "\u041A\u043E\u043F\u0438\u0438 \u0440\u0430\u0437\u0434\u0435\u043B\u0435\u043D\u044B \u043F\u0443\u043D\u043A\u0442\u0438\u0440\u043D\u044B\u043C\u0438 \u043B\u0438\u043D\u0438\u044F\u043C\u0438 \u2014 \u043F\u043E \u043D\u0438\u043C \u0443\u0434\u043E\u0431\u043D\u043E \u0440\u0435\u0437\u0430\u0442\u044C. \u0412 \u043E\u043A\u043D\u0435 \u043F\u0435\u0447\u0430\u0442\u0438 \u0432\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u043F\u0440\u0438\u043D\u0442\u0435\u0440 \u0438\u043B\u0438 \xAB\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u043A\u0430\u043A PDF\xBB.")), /*#__PURE__*/React.createElement(CardsPreview, {
    html: content,
    copies: copies,
    fontPx: fontPx,
    onOverflow: setOverflow
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: doWord
  }, "\uD83D\uDCC4 \u0421\u043A\u0430\u0447\u0430\u0442\u044C Word"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: !o.text && !(o.questions && nQ > 0),
    onClick: doPrint
  }, "\uD83D\uDDA8 \u041F\u0435\u0447\u0430\u0442\u044C / PDF"))));
}

// Печать домашнего задания: текст и картинки ДЗ; можно печатать всё ДЗ или одну картинку — N раз на листе.
function PrintHomeworkModal({
  lesson,
  onClose
}) {
  const parsed = useMemo(() => {
    const doc = new DOMParser().parseFromString("<body>" + cleanPrintHtml(lesson.homework || "") + "</body>", "text/html");
    const images = Array.from(doc.querySelectorAll("img")).map((im, i) => {
      const blk = im.closest(".lc-block");
      const nm = blk && blk.querySelector(".blk-name") && blk.querySelector(".blk-name").textContent || im.getAttribute("alt") || "Изображение " + (i + 1);
      return {
        src: im.getAttribute("src"),
        name: nm
      };
    });
    doc.querySelectorAll(".lc-block").forEach(b => {
      if (b.querySelector("img")) b.remove();
    });
    doc.querySelectorAll("img").forEach(im => im.remove());
    return {
      images,
      textHtml: doc.body.innerHTML
    };
  }, [lesson.homework]);
  const [mode, setMode] = useState("all");
  const [copies, setCopies] = useState(2);
  const [fontPx, setFontPx] = useState(14);
  const [incTitle, setIncTitle] = useState(true);
  const [incText, setIncText] = useState(true);
  const [incImg, setIncImg] = useState(true);
  const [sizes, setSizes] = useState({});
  const [overflow, setOverflow] = useState(false);
  useEffect(() => {
    parsed.images.forEach(a => {
      if (sizes[a.src]) return;
      const img = new Image();
      img.onload = () => setSizes(prev => ({
        ...prev,
        [a.src]: {
          w: img.naturalWidth,
          h: img.naturalHeight
        }
      }));
      img.src = a.src;
    });
  }, [parsed]);
  const cell = cardCellSize(copies);
  const hasText = !!stripHtml(parsed.textHtml).trim();
  let content = "";
  let title = "Домашнее задание — " + (lesson.topic || "урок");
  if (mode === "all") {
    const parts = [];
    if (incTitle) parts.push(`<div style="font-weight:bold;font-size:1.15em;margin-bottom:6px;">Домашнее задание</div>`);
    if (incText && hasText) parts.push(`<div style="font-weight:bold;margin-bottom:8px;">${parsed.textHtml}</div>`);
    if (incImg) parsed.images.forEach(a => parts.push(`<div style="margin-top:8px;">${fitImgHtml(a.src, sizes[a.src], cell.w, (cell.h - 2) * (hasText && incText ? 0.55 : 0.9))}</div>`));
    content = parts.join("");
  } else {
    const a = parsed.images[+mode];
    if (a) {
      title = a.name.replace(/\.[^.]+$/, "") || "Изображение";
      content = fitImgHtml(a.src, sizes[a.src], cell.w, cell.h - 2);
    }
  }
  const doPrint = () => {
    const h = cardsPageHtml(content, copies, fontPx, false);
    onClose();
    setTimeout(() => printCardsPdf(h), 60);
  };
  const doWord = async () => {
    const h = await inlineLitfileImages(cardsPageHtml(content, copies, fontPx, false));
    downloadWord(safeFileName(title) + ".doc", title, h, {
      noTitle: true,
      pageCss: CARDS_WORD_PAGE_CSS
    });
    onClose();
  };
  const radio = (val, label) => /*#__PURE__*/React.createElement("label", {
    key: val,
    className: "row",
    style: {
      gap: 6,
      padding: "3px 0",
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "radio",
    name: "hw-mode",
    checked: mode === val,
    onChange: () => setMode(val)
  }), " ", label);
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      width: "min(860px, 96vw)"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, "\u041F\u0435\u0447\u0430\u0442\u044C \u0434\u043E\u043C\u0430\u0448\u043D\u0435\u0433\u043E \u0437\u0430\u0434\u0430\u043D\u0438\u044F"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      gap: 18,
      alignItems: "flex-start",
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: "1 1 280px",
      minWidth: 250
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontWeight: 700,
      marginBottom: 4
    }
  }, "\u0427\u0442\u043E \u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C"), radio("all", "Всё домашнее задание (текст + изображения)"), parsed.images.map((a, i) => radio(String(i), "Только изображение: " + a.name)), mode === "all" && /*#__PURE__*/React.createElement("div", {
    style: {
      margin: "6px 0 10px 22px"
    }
  }, /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: incTitle,
    onChange: e => setIncTitle(e.target.checked)
  }), " \u0417\u0430\u0433\u043E\u043B\u043E\u0432\u043E\u043A \xAB\u0414\u043E\u043C\u0430\u0448\u043D\u0435\u0435 \u0437\u0430\u0434\u0430\u043D\u0438\u0435\xBB"), /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: incText,
    disabled: !hasText,
    onChange: e => setIncText(e.target.checked)
  }), " \u0422\u0435\u043A\u0441\u0442"), /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: incImg,
    disabled: !parsed.images.length,
    onChange: e => setIncImg(e.target.checked)
  }), " \u0418\u0437\u043E\u0431\u0440\u0430\u0436\u0435\u043D\u0438\u044F (", parsed.images.length, ")")), /*#__PURE__*/React.createElement("div", {
    style: {
      height: 8
    }
  }), /*#__PURE__*/React.createElement(CardsOptions, {
    copies: copies,
    setCopies: setCopies,
    fontPx: fontPx,
    setFontPx: setFontPx,
    noFont: mode !== "all"
  }), overflow && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12.5,
      color: "var(--neg)",
      background: "var(--neg-soft)",
      borderRadius: 8,
      padding: "8px 10px",
      marginTop: 8
    }
  }, "\u041D\u0435 \u043F\u043E\u043C\u0435\u0449\u0430\u0435\u0442\u0441\u044F \u0432 \u043E\u0434\u043D\u0443 \u044F\u0447\u0435\u0439\u043A\u0443: \u0447\u0430\u0441\u0442\u044C \u043E\u0431\u0440\u0435\u0436\u0435\u0442\u0441\u044F. \u0412\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u043C\u0435\u043D\u044C\u0448\u0435 \u043A\u043E\u043F\u0438\u0439 \u0438\u043B\u0438 \u043C\u0435\u043D\u044C\u0448\u0438\u0439 \u0440\u0430\u0437\u043C\u0435\u0440 \u0442\u0435\u043A\u0441\u0442\u0430."), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 8
    }
  }, "\u041E\u0434\u043D\u043E \u0438 \u0442\u043E \u0436\u0435 \u0414\u0417 \u043F\u043E\u0432\u0442\u043E\u0440\u044F\u0435\u0442\u0441\u044F \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u043E\u0435 \u0447\u0438\u0441\u043B\u043E \u0440\u0430\u0437 \u043D\u0430 \u043B\u0438\u0441\u0442\u0435; \u043A\u043E\u043F\u0438\u0438 \u0440\u0430\u0437\u0434\u0435\u043B\u0435\u043D\u044B \u043F\u0443\u043D\u043A\u0442\u0438\u0440\u043E\u043C \u2014 \u043F\u043E \u043D\u0435\u043C\u0443 \u0443\u0434\u043E\u0431\u043D\u043E \u0440\u0435\u0437\u0430\u0442\u044C.")), /*#__PURE__*/React.createElement(CardsPreview, {
    html: content,
    copies: copies,
    fontPx: fontPx,
    onOverflow: setOverflow
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    disabled: !content,
    onClick: doWord
  }, "\uD83D\uDCC4 \u0421\u043A\u0430\u0447\u0430\u0442\u044C Word"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: !content,
    onClick: doPrint
  }, "\uD83D\uDDA8 \u041F\u0435\u0447\u0430\u0442\u044C / PDF"))));
}

// Выбор курсов, для которых материал доступен (классы задаются метками классов; пусто — доступен всем).
function AudiencePicker({
  data,
  classes,
  courseIds,
  onChange,
  showClasses = true
}) {
  const toggle = (arr, v) => arr.includes(v) ? arr.filter(x => x !== v) : [...arr, v];
  const chip = (key, label, on, onClick) => /*#__PURE__*/React.createElement("button", {
    key: key,
    type: "button",
    className: "chip",
    onClick: onClick,
    style: {
      cursor: "pointer",
      borderColor: on ? "var(--accent)" : undefined,
      background: on ? "var(--accent-soft)" : undefined,
      color: on ? "var(--accent)" : undefined,
      fontWeight: on ? 700 : undefined
    }
  }, on ? "✓ " : "", label);
  return /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      alignItems: "center"
    }
  }, showClasses && CLASSES.map(c => chip("c" + c, c + " класс", classes.includes(c), () => onChange({
    classes: toggle(classes, c),
    courseIds
  }))), data.courses.map(c => chip("k" + c.id, "Курс «" + c.name + "»", courseIds.includes(c.id), () => onChange({
    classes,
    courseIds: toggle(courseIds, c.id)
  }))), classes.length > 0 || courseIds.length > 0 ? /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => onChange({
      classes: [],
      courseIds: []
    })
  }, "\u0421\u0431\u0440\u043E\u0441\u0438\u0442\u044C (\u0434\u043E\u0441\u0442\u0443\u043F\u043D\u043E \u0432\u0441\u0435\u043C)") : /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, data.courses.length ? "ничего не выбрано — доступно всем" : "курсов пока нет"));
}

// Применяет к карточке (упражнению/персоналии/…) изменение «не отображать» / «отображать всегда» для набора областей.
// Одна область не может быть одновременно в обоих списках: последнее действие побеждает.
function applyScopeMode(e, keys, mode) {
  const hidden = new Set(e.hiddenFor || []),
    always = new Set(e.alwaysFor || []);
  keys.forEach(k => {
    if (mode === "hide") {
      hidden.add(k);
      always.delete(k);
    } else if (mode === "show") hidden.delete(k);else if (mode === "always") {
      always.add(k);
      hidden.delete(k);
    } else if (mode === "unalways") always.delete(k);
  });
  e.hiddenFor = Array.from(hidden);
  e.alwaysFor = Array.from(always);
}
function scopeModeToast(mode) {
  return mode === "hide" ? "Больше не отображается: " : mode === "show" ? "Снова показывается: " : mode === "always" ? "Отображается всегда: " : "Больше не «всегда»: ";
}

// Чипы областей с кнопкой изменения: «Не отображать в…» и «Отображать всегда в…» у одной карточки.
function HiddenScopesRow({
  data,
  hiddenFor,
  alwaysFor,
  onChange
}) {
  const [variant, setVariant] = useState(null);
  const hid = hiddenFor || [],
    alw = alwaysFor || [];
  const apply = (keys, mode) => {
    const tmp = {
      hiddenFor: hid.slice(),
      alwaysFor: alw.slice()
    };
    applyScopeMode(tmp, keys, mode);
    onChange({
      hiddenFor: tmp.hiddenFor,
      alwaysFor: tmp.alwaysFor
    });
    setVariant(null);
  };
  const dropHidden = k => onChange({
    hiddenFor: hid.filter(x => x !== k),
    alwaysFor: alw
  });
  const dropAlways = k => onChange({
    hiddenFor: hid,
    alwaysFor: alw.filter(x => x !== k)
  });
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      alignItems: "center"
    }
  }, hid.map(k => /*#__PURE__*/React.createElement("span", {
    key: k,
    className: "chip"
  }, "\uD83D\uDEAB ", scopeLabelOf(data, k), /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => dropHidden(k)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  })))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setVariant("hide")
  }, "\u041D\u0435 \u043E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0442\u044C \u0432\u2026"), hid.length === 0 && /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u043D\u0438\u0433\u0434\u0435 \u043D\u0435 \u043E\u0442\u043A\u043B\u044E\u0447\u0435\u043D\u043E")), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      alignItems: "center",
      marginTop: 6
    }
  }, alw.map(k => /*#__PURE__*/React.createElement("span", {
    key: k,
    className: "chip",
    style: {
      borderColor: "var(--gold)"
    }
  }, "\u2605 ", scopeLabelOf(data, k), /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => dropAlways(k)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  })))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setVariant("always")
  }, "\u041E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0442\u044C \u0432\u0441\u0435\u0433\u0434\u0430 \u0432\u2026"), alw.length === 0 && /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\xAB\u0432\u0441\u0435\u0433\u0434\u0430\xBB \u043D\u0438\u0433\u0434\u0435 \u043D\u0435 \u0432\u043A\u043B\u044E\u0447\u0435\u043D\u043E")), variant && /*#__PURE__*/React.createElement(HidePoemsModal, {
    data: data,
    count: 1,
    variant: variant,
    onClose: () => setVariant(null),
    onApply: apply
  }));
}

// Чипы «отображать всегда» для целого ВИДА упражнений (все его упражнения).
function AlwaysScopesRow({
  data,
  scopes,
  onChange
}) {
  const [open, setOpen] = useState(false);
  const cur = scopes || [];
  return /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      alignItems: "center"
    }
  }, cur.map(k => /*#__PURE__*/React.createElement("span", {
    key: k,
    className: "chip",
    style: {
      borderColor: "var(--gold)"
    }
  }, "\u2605 ", scopeLabelOf(data, k), /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => onChange(cur.filter(x => x !== k))
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 11
  })))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setOpen(true)
  }, "\u041E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0442\u044C \u0432\u0441\u0435\u0433\u0434\u0430 \u0432\u2026"), cur.length === 0 && /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u043D\u0435 \u0437\u0430\u0434\u0430\u043D\u043E"), open && /*#__PURE__*/React.createElement(HidePoemsModal, {
    data: data,
    count: 0,
    variant: "always",
    onClose: () => setOpen(false),
    onApply: (keys, mode) => {
      const s = new Set(cur);
      keys.forEach(k => mode === "always" ? s.add(k) : s.delete(k));
      onChange(Array.from(s));
      setOpen(false);
    }
  }));
}

// Окно выбора классов/литер/курсов: variant "hide" — где карточки НЕ предлагать, "always" — где предлагать ВСЕГДА (перебивает всё остальное).
function HidePoemsModal({
  data,
  count,
  onApply,
  onClose,
  variant = "hide"
}) {
  const [keys, setKeys] = useState([]);
  const always = variant === "always";
  const toggle = k => setKeys(prev => prev.includes(k) ? prev.filter(x => x !== k) : [...prev, k]);
  const box = (k, label) => /*#__PURE__*/React.createElement("label", {
    key: k,
    className: "row",
    style: {
      gap: 6,
      fontSize: 13,
      cursor: "pointer",
      padding: "2px 0"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: keys.includes(k),
    onChange: () => toggle(k)
  }), " ", label);
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      width: "min(560px, 94vw)"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, always ? "Отображать всегда в…" : "Не отображать в…"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      marginTop: 0
    }
  }, count > 0 ? /*#__PURE__*/React.createElement(React.Fragment, null, "\u0412\u044B\u0431\u0440\u0430\u043D\u043E \u043A\u0430\u0440\u0442\u043E\u0447\u0435\u043A: ", /*#__PURE__*/React.createElement("b", null, count), ". ") : null, always ? "Отметьте, где они предлагаются при составлении уроков ВСЕГДА — даже если уже использованы, отключены вручную или у них метка другого класса. «Класс целиком» действует во всех литерах этого класса." : "Отметьте, где их больше не предлагать при составлении уроков. «Класс целиком» скрывает во всех литерах этого класса."), CLASSES.map(c => /*#__PURE__*/React.createElement("div", {
    key: c,
    className: "row wrap",
    style: {
      gap: 14,
      alignItems: "center",
      borderBottom: "1px solid var(--line)",
      padding: "3px 0"
    }
  }, box("c:" + c, c + " класс целиком"), lettersForGrade(data, c).map(cl => box("l:" + c + ":" + cl.id, c + cl.name)))), data.courses.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 14,
      padding: "3px 0"
    }
  }, data.courses.map(c => box("k:" + c.id, "Курс «" + c.name + "»"))), !always && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 8
    }
  }, "\u0423\u0440\u043E\u043A\u0438, \u0432 \u043A\u043E\u0442\u043E\u0440\u044B\u0435 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B \u0443\u0436\u0435 \u0434\u043E\u0431\u0430\u0432\u043B\u0435\u043D, \u043D\u0435 \u043C\u0435\u043D\u044F\u044E\u0442\u0441\u044F. \u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B, \u0443\u0436\u0435 \u0438\u0441\u043F\u043E\u043B\u044C\u0437\u043E\u0432\u0430\u043D\u043D\u044B\u0435 \u0432 \u0443\u0440\u043E\u043A\u0430\u0445 \u043A\u043B\u0430\u0441\u0441\u0430/\u043B\u0438\u0442\u0435\u0440\u044B, \u0441\u043A\u0440\u044B\u0432\u0430\u044E\u0442\u0441\u044F \u0430\u0432\u0442\u043E\u043C\u0430\u0442\u0438\u0447\u0435\u0441\u043A\u0438.")), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    disabled: !keys.length,
    onClick: () => onApply(keys, always ? "unalways" : "show")
  }, always ? "Убрать «всегда» там" : "Снова показывать там"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: !keys.length,
    onClick: () => onApply(keys, always ? "always" : "hide")
  }, always ? "Отображать всегда" : "Не отображать"))));
}
function LessonPage({
  data,
  update,
  lessonId,
  onBack,
  projectPath
}) {
  const [unarchiveOpen, setUnarchiveOpen] = useState(false);
  const lesson = data.lessons.find(l => l.id === lessonId);
  const [peek, setPeek] = useState(null);
  const [printTarget, setPrintTarget] = useState(null);
  const [printHw, setPrintHw] = useState(false);
  const [copyLettersOpen, setCopyLettersOpen] = useState(false);
  const [copyLettersChecked, setCopyLettersChecked] = useState([]);
  const [matModal, setMatModal] = useState(null); // { mode: "create" | "edit", type, id? } — окно создания/правки материала
  if (!lesson) return /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u0423\u0440\u043E\u043A \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D. ", /*#__PURE__*/React.createElement(Btn, {
    onClick: onBack
  }, "\u041D\u0430\u0437\u0430\u0434"));
  const patch = p => update(d => {
    const idx = d.lessons.findIndex(l => l.id === lessonId);
    if (idx >= 0) {
      d.lessons[idx] = {
        ...d.lessons[idx],
        ...p,
        updatedAt: nowISO()
      };
      if (p.plannedDate !== undefined || p.classLetterId !== undefined) resyncAfterDateChange(d, d.lessons[idx]);
    }
    return d;
  });
  const addExerciseItem = refId => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (l && !l.items.some(it => it.type === "exercise" && it.refId === refId)) l.items.push(migrateLessonItem({
      type: "exercise",
      refId
    }));
    return d;
  });
  const addWholeItem = (type, refId) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (l && !l.items.some(it => it.type === type && it.refId === refId)) l.items.push(migrateLessonItem({
      type,
      refId
    }));
    return d;
  });
  const addQuestionItem = (entityType, entityId, questionId) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (l && !l.items.some(it => it.type === "question" && it.questionId === questionId)) l.items.push(migrateLessonItem({
      type: "question",
      entityType,
      entityId,
      questionId
    }));
    return d;
  });
  // Добавить сразу несколько вопросов карточки (уже добавленные не дублируются, порядок — как в карточке)
  const addQuestionsBulk = (entityType, entityId, questionIds) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (!l) return d;
    questionIds.forEach(qid => {
      if (!l.items.some(it => it.type === "question" && it.questionId === qid)) l.items.push(migrateLessonItem({
        type: "question",
        entityType,
        entityId,
        questionId: qid
      }));
    });
    return d;
  });
  const removeItemsBulk = itemIds => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (!l) return d;
    const gone = new Set(itemIds);
    l.items.filter(it => gone.has(it.id)).forEach(v => {
      const markIds = itemStudents(v).map(st => st.markId).filter(Boolean);
      d.calendar.marks = d.calendar.marks.filter(m => !markIds.includes(m.id));
    });
    l.items = l.items.filter(it => !gone.has(it.id));
    return d;
  });
  const removeItem = itemId => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    const it = l.items.find(x => x.id === itemId);
    // вместе с упражнением уходят и его вопросы в этом уроке
    const victims = it ? [it] : [];
    if (it && it.type === "exercise") l.items.filter(x => x.type === "question" && x.entityType === "exercise" && x.entityId === it.refId).forEach(x => victims.push(x));
    victims.forEach(v => {
      const markIds = itemStudents(v).map(st => st.markId).filter(Boolean);
      d.calendar.marks = d.calendar.marks.filter(m => !markIds.includes(m.id));
      if (v.type === "question" && v.local) stripQaAnchors(d, v.entityType, v.entityId, v.questionId);
    });
    const gone = new Set(victims.map(v => v.id));
    l.items = l.items.filter(x => !gone.has(x.id));
    return d;
  });
  // Перестановка блоков (упражнение / карточка с вопросами) стрелками; соседом считается ближайший ВИДИМЫЙ блок.
  const blockVisible = (d, b) => b.kind === "exercise" || b.kind === "group" && !!questionSourceEntity(d, {
    entityType: b.entityType,
    entityId: b.entityId
  });
  const moveBlock = (key, dir) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    const bl = lessonItemBlocks(l.items);
    const i = bl.findIndex(b => b.key === key);
    if (i < 0) return d;
    let j = i + dir;
    while (j >= 0 && j < bl.length && !blockVisible(d, bl[j])) j += dir;
    if (j < 0 || j >= bl.length) return d;
    [bl[i], bl[j]] = [bl[j], bl[i]];
    const next = bl.flatMap(b => b.items);
    if (next.length === l.items.length) l.items = next;
    return d;
  });
  const moveQuestion = (itemId, dir) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    const it = l.items.find(x => x.id === itemId);
    if (!it) return d;
    const same = l.items.filter(x => x.type === "question" && x.entityType === it.entityType && x.entityId === it.entityId);
    const i = same.findIndex(x => x.id === itemId),
      j = i + dir;
    if (i < 0 || j < 0 || j >= same.length) return d;
    const a = l.items.indexOf(same[i]),
      b = l.items.indexOf(same[j]);
    [l.items[a], l.items[b]] = [l.items[b], l.items[a]];
    return d;
  });
  const letters = !lesson.courseId ? lettersForGrade(data, lesson.class) : [];
  const otherLetters = letters.filter(cl => cl.id !== (lesson.classLetterId || null));
  // «Перенести в другие литеры»: учитель сам отмечает, в какие литеры скопировать урок
  // (тема, теги, материалы и ДЗ) — как в конструкторе русского языка.
  const copyToOtherLetters = () => {
    if (copyLettersChecked.length === 0) return;
    update(d => {
      copyLettersChecked.forEach(letterId => duplicateLessonInto(d, lessonId, {
        classLetterId: letterId
      }));
      return d;
    });
    showToast("Урок скопирован в: " + otherLetters.filter(cl => copyLettersChecked.includes(cl.id)).map(cl => lesson.class + cl.name).join(", "));
    setCopyLettersOpen(false);
    setCopyLettersChecked([]);
  };
  const moveToReserve = item => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    l.items = l.items.filter(it => it.id !== item.id && !(item.type === "exercise" && it.type === "question" && it.entityType === "exercise" && it.entityId === item.refId));
    const groupKind = l.class ? "class" : "course";
    const groupValue = l.class || l.courseId;
    if (item.type === "exercise" && !d.reserves.some(r => r.exerciseId === item.refId && r.groupKind === groupKind && r.groupValue === groupValue)) {
      d.reserves.push({
        id: uid(),
        groupKind,
        groupValue,
        exerciseId: item.refId
      });
    }
    return d;
  });
  const removeLesson = async () => {
    if (!(await Confirm("Удалить урок? Упражнения и материалы, использованные в нём, останутся в базе."))) return;
    update(d => {
      d.lessons = d.lessons.filter(l => l.id !== lessonId);
      d.lessonSystems.forEach(s => {
        s.lessonIds = s.lessonIds.filter(id => id !== lessonId);
      });
      return d;
    });
    onBack();
  };
  const duplicateThis = async () => {
    const preview = duplicateLessonInSequence(structuredClone(data), lessonId);
    if (!preview) return;
    const msg = lesson.plannedDate ? "Создать дубликат этого урока на " + formatDateRu(preview.dupDate) + "?" + (preview.moved ? " Следующие уроки (" + preview.moved + ") сдвинутся по цепочке на даты друг друга; проведённые уроки и уроки с отметками не двигаются." : " Следующих уроков по календарю нет.") : "У этого урока нет даты — дубликат создастся без даты рядом с ним. Продолжить?";
    if (!(await Confirm(msg, "Да, дублировать"))) return;
    update(d => {
      duplicateLessonInSequence(d, lessonId);
      return d;
    });
    showToast("Дубликат создан" + (preview.dupDate ? " на " + formatDateRu(preview.dupDate) : "") + (preview.moved ? " · сдвинуто уроков: " + preview.moved : ""));
  };
  const exerciseItems = lesson.items.filter(it => it.type === "exercise");
  const entityGroups = [];
  const findOrAddGroup = (entityType, entityId) => {
    const key = entityType + ":" + entityId;
    let g = entityGroups.find(x => x.key === key);
    if (!g) {
      g = {
        key,
        entityType,
        entityId,
        wholeItemId: null,
        questionItemsList: []
      };
      entityGroups.push(g);
    }
    return g;
  };
  lesson.items.forEach(it => {
    if (it.type === "theory" || it.type === "personality" || it.type === "work") findOrAddGroup(it.type, it.refId).wholeItemId = it.id;else if (it.type === "question" && it.entityType !== "exercise") findOrAddGroup(it.entityType, it.entityId).questionItemsList.push(it);
  });
  const removeEntityGroup = g => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    const idsToRemove = new Set(g.questionItemsList.map(it => it.id));
    if (g.wholeItemId) idsToRemove.add(g.wholeItemId);
    const markIdsToRemove = new Set();
    idsToRemove.forEach(id => {
      const it = l.items.find(x => x.id === id);
      if (it) itemStudents(it).forEach(s => {
        if (s.markId) markIdsToRemove.add(s.markId);
      });
    });
    d.calendar.marks = d.calendar.marks.filter(m => !markIdsToRemove.has(m.id));
    l.items = l.items.filter(it => !idsToRemove.has(it.id));
    return d;
  });
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onBack
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "left",
    size: 15
  }), " \u041D\u0430\u0437\u0430\u0434"), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), otherLetters.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => {
      setCopyLettersOpen(!copyLettersOpen);
      setCopyLettersChecked([]);
    },
    title: "\u0421\u043A\u043E\u043F\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0443\u0440\u043E\u043A \u0432 \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u044B\u0435 \u043B\u0438\u0442\u0435\u0440\u044B"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u041F\u0435\u0440\u0435\u043D\u0435\u0441\u0442\u0438 \u0432 \u0434\u0440\u0443\u0433\u0438\u0435 \u043B\u0438\u0442\u0435\u0440\u044B"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => {
      exportLessonFile(data, lesson);
      showToast("Урок сохранён в файл .lsnlit");
    },
    title: "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u044D\u0442\u043E\u0442 \u0443\u0440\u043E\u043A \u0432 \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0439 \u0444\u0430\u0439\u043B .lsnlit \u2014 \u0435\u0433\u043E \u043C\u043E\u0436\u043D\u043E \u0438\u043C\u043F\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0432 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0443"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u042D\u043A\u0441\u043F\u043E\u0440\u0442 \u0443\u0440\u043E\u043A\u0430 (.lsnlit)"), /*#__PURE__*/React.createElement(Btn, {
    kind: "danger",
    onClick: removeLesson
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  }), " \u0423\u0434\u0430\u043B\u0438\u0442\u044C \u0443\u0440\u043E\u043A")), copyLettersOpen && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginTop: -8,
      marginBottom: 16,
      padding: 10,
      background: "var(--accent-soft)",
      borderRadius: "var(--radius)",
      gap: 10,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      fontWeight: 700
    }
  }, "\u0421\u043A\u043E\u043F\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0442\u0435\u043C\u0443, \u0442\u0435\u0433\u0438, \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B \u0438 \u0414\u0417 \u044D\u0442\u043E\u0433\u043E \u0443\u0440\u043E\u043A\u0430 \u0432 \u043B\u0438\u0442\u0435\u0440\u044B (\u0434\u0430\u0442\u0430 \u0438 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u043A\u0430\u0436\u0434\u043E\u0439 \u043B\u0438\u0442\u0435\u0440\u044B \u043D\u0435 \u0438\u0437\u043C\u0435\u043D\u044F\u0442\u0441\u044F):"), otherLetters.map(cl => /*#__PURE__*/React.createElement("label", {
    key: cl.id,
    className: "row",
    style: {
      gap: 4,
      fontSize: 12.5,
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: copyLettersChecked.includes(cl.id),
    onChange: e => setCopyLettersChecked(prev => e.target.checked ? [...prev, cl.id] : prev.filter(x => x !== cl.id))
  }), lesson.class, cl.name)), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    disabled: copyLettersChecked.length === 0,
    onClick: copyToOtherLetters
  }, "\u041F\u0435\u0440\u0435\u043D\u0435\u0441\u0442\u0438")), /*#__PURE__*/React.createElement("input", {
    className: "titleinput",
    value: lesson.topic,
    placeholder: "\u0422\u0435\u043C\u0430 \u0443\u0440\u043E\u043A\u0430",
    onChange: e => patch({
      topic: e.target.value
    })
  }), lesson.duplicateOf && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: -4,
      marginBottom: 6
    }
  }, "(\u0434\u0443\u0431\u043B\u0438\u043A\u0430\u0442) \u2014 \u043F\u0440\u043E\u0434\u043E\u043B\u0436\u0435\u043D\u0438\u0435 \u0443\u0440\u043E\u043A\u0430", (() => {
    const o = data.lessons.find(x => x.id === lesson.duplicateOf);
    return o ? " «" + (o.topic || "без темы") + "»" : "";
  })()), /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.lessonType.title
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "lessonType",
    selectedIds: lesson.lessonTypeTagIds,
    onChange: ids => patch({
      lessonTypeTagIds: ids
    })
  })), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      margin: "10px 0"
    }
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u0414\u0430\u0442\u0430 \u043F\u043E \u043F\u043B\u0430\u043D\u0443"
  }, /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: lesson.plannedDate || "",
    onChange: e => patch({
      plannedDate: e.target.value
    }),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: 7
    }
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u0414\u0430\u0442\u0430 \u043F\u043E \u0444\u0430\u043A\u0442\u0443"
  }, /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: lesson.actualDate || "",
    onChange: e => patch({
      actualDate: e.target.value
    }),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: 7
    }
  })), letters.length > 0 && /*#__PURE__*/React.createElement(Field, {
    label: "\u041B\u0438\u0442\u0435\u0440\u0430"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: lesson.classLetterId || "",
    onChange: e => patch({
      classLetterId: e.target.value || null
    })
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0411\u0435\u0437 \u043B\u0438\u0442\u0435\u0440\u044B"), letters.map(cl => /*#__PURE__*/React.createElement("option", {
    key: cl.id,
    value: cl.id
  }, lesson.class, cl.name)))), /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      marginTop: 20,
      gap: 6,
      background: lesson.needsWork ? "#fdf8ef" : undefined,
      border: lesson.needsWork ? "1px solid #e7d3a8" : "1px solid transparent",
      borderRadius: 8,
      padding: "2px 8px",
      fontWeight: lesson.needsWork ? 700 : 400
    },
    title: "\u041F\u043E\u043C\u0435\u0442\u043A\u0430 \u0434\u043B\u044F \u0441\u0435\u0431\u044F: \u0443\u0440\u043E\u043A \u043D\u0443\u0436\u043D\u043E \u0434\u043E\u0440\u0430\u0431\u043E\u0442\u0430\u0442\u044C. \u0412 \u0441\u043F\u0438\u0441\u043A\u0435 \xAB\u0423\u0440\u043E\u043A\u0438\xBB \u0442\u0430\u043A\u043E\u0439 \u0443\u0440\u043E\u043A \u043F\u043E\u0434\u0441\u0432\u0435\u0447\u0435\u043D"
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!lesson.needsWork,
    onChange: e => patch({
      needsWork: e.target.checked
    })
  }), " \u0422\u0440\u0435\u0431\u0443\u0435\u0442 \u0434\u043E\u0440\u0430\u0431\u043E\u0442\u043A\u0438"), /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      marginTop: 20,
      gap: 6
    },
    title: "\u041F\u0440\u043E\u0439\u0434\u0435\u043D\u043D\u044B\u0439 \u0443\u0440\u043E\u043A \u0443\u0445\u043E\u0434\u0438\u0442 \u0432 \u0430\u0440\u0445\u0438\u0432 (\u0438\u0437 \u0441\u043F\u0438\u0441\u043A\u0430 \xAB\u0423\u0440\u043E\u043A\u0438\xBB), \u043D\u043E \u043E\u0441\u0442\u0430\u0451\u0442\u0441\u044F \u0432 \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u0435"
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!lesson.completed,
    onChange: e => update(d => {
      const l = d.lessons.find(x => x.id === lessonId);
      if (l) setLessonCompleted(l, e.target.checked);
      return d;
    })
  }), " \u041F\u0440\u043E\u0432\u0435\u0434\u0451\u043D"), lesson.archived ? /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      marginTop: 14,
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u0412 \u0430\u0440\u0445\u0438\u0432\u0435"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: () => setUnarchiveOpen(true)
  }, "\u21A9 \u0420\u0430\u0437\u0430\u0440\u0445\u0438\u0432\u0438\u0440\u043E\u0432\u0430\u0442\u044C")) : /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      marginTop: 20,
      gap: 6
    },
    title: "\u0410\u0440\u0445\u0438\u0432\u043D\u044B\u0439 \u0443\u0440\u043E\u043A \u043F\u0440\u043E\u043F\u0430\u0434\u0430\u0435\u0442 \u0438\u0437 \u0441\u043F\u0438\u0441\u043A\u0430 \xAB\u0423\u0440\u043E\u043A\u0438\xBB, \u043D\u043E \u043E\u0441\u0442\u0430\u0451\u0442\u0441\u044F \u0432 \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u0435"
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: false,
    onChange: e => patch({
      archived: e.target.checked
    })
  }), " \u0412 \u0430\u0440\u0445\u0438\u0432"), unarchiveOpen && /*#__PURE__*/React.createElement(UnarchiveModal, {
    data: data,
    update: update,
    lessonId: lesson.id,
    onClose: () => setUnarchiveOpen(false)
  })), /*#__PURE__*/React.createElement("div", {
    className: "panel panel-sticky-host"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-head"
  }, /*#__PURE__*/React.createElement("h2", null, "\u0423\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F \u0438 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B")), /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, exerciseItems.length === 0 && entityGroups.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "muted",
    style: {
      fontSize: 13,
      marginBottom: 8
    }
  }, "\u041F\u043E\u043A\u0430 \u043D\u0435 \u0434\u043E\u0431\u0430\u0432\u043B\u0435\u043D\u044B."), (() => {
    const blocks = lessonItemBlocks(lesson.items).filter(bl => blockVisible(data, bl));
    return blocks.map((bl, bi) => {
      const arrows = blocks.length > 1 ? /*#__PURE__*/React.createElement(MoveArrows, {
        canUp: bi > 0,
        canDown: bi < blocks.length - 1,
        onUp: () => moveBlock(bl.key, -1),
        onDown: () => moveBlock(bl.key, 1)
      }) : null;
      if (bl.kind === "group") {
        const g = entityGroups.find(x => x.key === bl.key);
        if (!g) return null;
        return /*#__PURE__*/React.createElement("div", {
          key: bl.key,
          className: "row",
          style: {
            alignItems: "flex-start",
            gap: 6
          }
        }, arrows, /*#__PURE__*/React.createElement("div", {
          style: {
            flex: 1,
            minWidth: 0
          }
        }, /*#__PURE__*/React.createElement(EntityMaterialRow, {
          data: data,
          update: update,
          lesson: lesson,
          entityType: g.entityType,
          entityId: g.entityId,
          questionItemsList: g.questionItemsList,
          hasWhole: !!g.wholeItemId,
          onRemoveQuestion: removeItem,
          onRemoveGroup: () => removeEntityGroup(g),
          onMoveQuestion: moveQuestion,
          onPrint: () => setPrintTarget({
            type: g.entityType,
            id: g.entityId
          }),
          onEdit: () => setMatModal({
            mode: "edit",
            type: g.entityType,
            id: g.entityId
          })
        })));
      }
      const it = bl.items[0];
      const obj = itemRefObj(data, it);
      return /*#__PURE__*/React.createElement("div", {
        key: it.id,
        className: "row",
        style: {
          alignItems: "flex-start",
          gap: 6
        }
      }, arrows, /*#__PURE__*/React.createElement("div", {
        className: "qrow",
        style: {
          flex: 1,
          minWidth: 0
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "row",
        style: {
          alignItems: "flex-start"
        }
      }, /*#__PURE__*/React.createElement("span", {
        className: "chip",
        style: {
          flex: "none"
        }
      }, ITEM_TYPE_META.exercise.label), /*#__PURE__*/React.createElement("div", {
        style: {
          flex: 1,
          fontSize: 13,
          cursor: "pointer"
        },
        onClick: () => setPeek({
          type: "exercise",
          refId: it.refId
        })
      }, obj ? ITEM_TYPE_META.exercise.titleOf(data, obj) : "(удалено)"), obj && obj.typeTagId === TROPE_WALK_TYPE_ID && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => setPeek({
          type: "exercise",
          refId: it.refId
        }),
        title: "\u041E\u0442\u043A\u0440\u044B\u0442\u044C \u043A\u0440\u0443\u043F\u043D\u043E \u0434\u043B\u044F \u043F\u043E\u043A\u0430\u0437\u0430 \u043D\u0430 \u044D\u043A\u0440\u0430\u043D\u0435 (\u0441 \u043C\u0430\u0441\u0448\u0442\u0430\u0431\u043E\u043C)"
      }, /*#__PURE__*/React.createElement(Icon, {
        name: "eye",
        size: 12
      }), " \u041A\u0440\u0443\u043F\u043D\u043E"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => exportTropeWalkWord(obj, data),
        title: "\u041A\u0430\u0440\u0442\u043E\u0447\u043A\u0430 \u0434\u043B\u044F \u043F\u0435\u0447\u0430\u0442\u0438 \u2014 Word"
      }, "Word"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => exportTropeWalkPdf(obj, data),
        title: "\u041A\u0430\u0440\u0442\u043E\u0447\u043A\u0430 \u0434\u043B\u044F \u043F\u0435\u0447\u0430\u0442\u0438 \u2014 PDF"
      }, "PDF"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => exportTropeWalkPptx(obj, data),
        title: "\u0421\u043B\u0430\u0439\u0434 \u2014 PowerPoint"
      }, "PPTX")), obj && /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => setMatModal({
          mode: "edit",
          type: "exercise",
          id: it.refId
        }),
        title: "\u041F\u0440\u0430\u0432\u0438\u0442\u044C \u0442\u0435\u043A\u0441\u0442 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F (\u0438\u0437\u043C\u0435\u043D\u0438\u0442\u0441\u044F \u0432 \u0431\u0430\u043D\u043A\u0435 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439)"
      }, "\u270E"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => setPrintTarget({
          type: "exercise",
          id: it.refId
        }),
        title: "\u0420\u0430\u0441\u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C \u044D\u0442\u043E \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0435 (\u0441 \u0432\u043E\u043F\u0440\u043E\u0441\u0430\u043C\u0438 \u0443\u0440\u043E\u043A\u0430 \u0438 \u043F\u0440\u0438\u043C\u0435\u0447\u0430\u043D\u0438\u044F\u043C\u0438)"
      }, "\uD83D\uDDA8"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => moveToReserve(it),
        title: "\u0412 \u0437\u0430\u043F\u0430\u0441"
      }, "\u21C4"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => removeItem(it.id)
      }, /*#__PURE__*/React.createElement(Icon, {
        name: "trash",
        size: 12
      }))), /*#__PURE__*/React.createElement(StickyNote, {
        value: it.note,
        onChange: v => update(d => {
          const l2 = d.lessons.find(x => x.id === lessonId);
          const i2 = l2 && l2.items.find(x => x.id === it.id);
          if (i2) i2.note = v;
          return d;
        })
      }), /*#__PURE__*/React.createElement(GradeAssignRow, {
        data: data,
        update: update,
        lesson: lesson,
        item: it
      }), obj && /*#__PURE__*/React.createElement(ExerciseQuestionsBlock, {
        data: data,
        update: update,
        lesson: lesson,
        exItem: it,
        questionItems: bl.items.slice(1),
        onRemoveQuestion: removeItem,
        onMoveQuestion: moveQuestion
      })));
    });
  })(), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement(ItemPicker, {
    data: data,
    lesson: lesson,
    onAddExercise: addExerciseItem,
    onAddWhole: addWholeItem,
    onAddQuestion: addQuestionItem,
    onAddQuestions: addQuestionsBulk,
    onRemoveItems: removeItemsBulk,
    onRemoveItem: removeItem,
    extra: /*#__PURE__*/React.createElement(PopMenu, {
      label: "\u270E \u0421\u043E\u0437\u0434\u0430\u0442\u044C \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B",
      title: "\u0421\u043E\u0437\u0434\u0430\u0442\u044C \u043D\u043E\u0432\u044B\u0439 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B \u043F\u0440\u044F\u043C\u043E \u0432 \u0443\u0440\u043E\u043A\u0435: \u043E\u043D \u0441\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u0441\u044F \u0432 \u043D\u0443\u0436\u043D\u043E\u043C \u0440\u0430\u0437\u0434\u0435\u043B\u0435 \u0438 \u0432\u0441\u0442\u0430\u043D\u0435\u0442 \u0432 \u0443\u0440\u043E\u043A",
      items: PICKER_SECTIONS.map(sec => ({
        key: sec.type,
        label: sec.icon + " " + NEW_MATERIAL_LABEL[sec.type]
      })),
      onPick: t => setMatModal({
        mode: "create",
        type: t
      })
    })
  })))), /*#__PURE__*/React.createElement(Field, {
    label: "\u0414\u043E\u043C\u0430\u0448\u043D\u0435\u0435 \u0437\u0430\u0434\u0430\u043D\u0438\u0435"
  }, /*#__PURE__*/React.createElement(RichEditor, {
    html: lesson.homework,
    onChange: html => patch({
      homework: html
    }),
    projectPath: projectPath,
    compact: true,
    update: update,
    sourceInfo: {
      section: "Уроки",
      entityId: lesson.id,
      entityTitle: lesson.topic
    }
  }), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginTop: 6,
      gap: 8,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !stripHtml(lesson.homework || "").trim() && !/<img/i.test(lesson.homework || ""),
    onClick: () => setPrintHw(true),
    title: "\u0420\u0430\u0441\u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C \u0438\u043B\u0438 \u0441\u043A\u0430\u0447\u0430\u0442\u044C \u0434\u043E\u043C\u0430\u0448\u043D\u0435\u0435 \u0437\u0430\u0434\u0430\u043D\u0438\u0435 (\u0442\u0435\u043A\u0441\u0442 \u0438 \u0438\u0437\u043E\u0431\u0440\u0430\u0436\u0435\u043D\u0438\u044F) \u2014 \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u043E \u043A\u043E\u043F\u0438\u0439 \u043D\u0430 \u043B\u0438\u0441\u0442\u0435"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 12
  }), " \uD83D\uDDA8 \u041F\u0435\u0447\u0430\u0442\u044C / \u0441\u043A\u0430\u0447\u0430\u0442\u044C \u0414\u0417"), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u0414\u0417 \u043C\u043E\u0436\u043D\u043E \u043D\u0430\u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C \u043F\u043E \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u043E \u043A\u043E\u043F\u0438\u0439 \u043D\u0430 \u043B\u0438\u0441\u0442\u0435; \u0438\u0437\u043E\u0431\u0440\u0430\u0436\u0435\u043D\u0438\u044F \u0414\u0417 \u2014 \u0442\u043E\u0436\u0435."))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginTop: 16,
      gap: 10,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    onClick: duplicateThis,
    title: "\u0421\u043E\u0437\u0434\u0430\u0442\u044C \u043F\u0440\u043E\u0434\u043E\u043B\u0436\u0435\u043D\u0438\u0435 \u044D\u0442\u043E\u0433\u043E \u0443\u0440\u043E\u043A\u0430"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u0414\u0443\u0431\u043B\u0438\u0440\u043E\u0432\u0430\u0442\u044C"), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u0411\u0435\u0441\u0435\u0434\u0430 \u0437\u0430\u0442\u044F\u043D\u0443\u043B\u0430\u0441\u044C? \u0414\u0443\u0431\u043B\u0438\u043A\u0430\u0442 \u0432\u0441\u0442\u0430\u043D\u0435\u0442 \u0432 \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u0435 \u043D\u0430 \u0434\u0430\u0442\u0443 \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0435\u0433\u043E \u0443\u0440\u043E\u043A\u0430, \u0430 \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0438\u0435 \u0443\u0440\u043E\u043A\u0438 \u0441\u0434\u0432\u0438\u043D\u0443\u0442\u0441\u044F \u043F\u043E \u0446\u0435\u043F\u043E\u0447\u043A\u0435.")), peek && /*#__PURE__*/React.createElement(MaterialPeekModal, {
    data: data,
    type: peek.type,
    refId: peek.refId,
    onClose: () => setPeek(null)
  }), matModal && /*#__PURE__*/React.createElement(LessonMaterialModal, {
    data: data,
    update: update,
    projectPath: projectPath,
    lesson: lesson,
    spec: matModal,
    onCreated: (t, id) => t === "exercise" ? addExerciseItem(id) : addWholeItem(t, id),
    onClose: () => setMatModal(null)
  }), printHw && /*#__PURE__*/React.createElement(PrintHomeworkModal, {
    lesson: lesson,
    onClose: () => setPrintHw(false)
  }), printTarget && /*#__PURE__*/React.createElement(PrintMaterialModal, {
    data: data,
    type: printTarget.type,
    entityId: printTarget.id,
    lesson: lesson,
    onClose: () => setPrintTarget(null)
  }));
}

// Тонкий зазор между верхнеуровневыми блоками списка "Уроки" (уроками и системами) — во время
// перетаскивания превращается в зону сброса "поставить именно сюда", в обычном состоянии не
// занимает места и никак не виден.
function TopLevelGap({
  dnd,
  target
}) {
  const [over, setOver] = useState(false);
  if (!dnd.dragLessonId && !dnd.dragSystemId) return null;
  return /*#__PURE__*/React.createElement("div", {
    onDragOver: e => {
      e.preventDefault();
      setOver(true);
    },
    onDragLeave: () => setOver(false),
    onDrop: e => {
      e.preventDefault();
      e.stopPropagation();
      setOver(false);
      dnd.insertDraggedBefore(target);
    },
    title: "\u041E\u0442\u043F\u0443\u0441\u0442\u0438\u0442\u0435 \u0437\u0434\u0435\u0441\u044C, \u0447\u0442\u043E\u0431\u044B \u043F\u043E\u0441\u0442\u0430\u0432\u0438\u0442\u044C \u0438\u043C\u0435\u043D\u043D\u043E \u0432 \u044D\u0442\u043E \u043C\u0435\u0441\u0442\u043E",
    style: {
      height: over ? 34 : 22,
      margin: "3px 0",
      borderRadius: 8,
      transition: "height .12s",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      fontSize: 12,
      fontWeight: 700,
      color: over ? "var(--accent)" : "var(--faint)",
      background: over ? "var(--accent-soft)" : "transparent",
      border: over ? "2px dashed var(--accent)" : "2px dashed var(--line)"
    }
  }, over ? "Поставить сюда" : "· · ·");
}
function LessonCard({
  data,
  lesson,
  index,
  onOpen,
  onMove,
  onRemoveFromSystem,
  onToggleDone,
  drag
}) {
  const typeTags = (lesson.lessonTypeTagIds || []).map(id => data.tags.lessonType.find(t => t.id === id)).filter(Boolean);
  const [dragOver, setDragOver] = useState(false);
  return /*#__PURE__*/React.createElement("div", {
    className: "qrow",
    style: {
      cursor: "pointer",
      outline: dragOver ? "2px solid var(--accent)" : "none",
      outlineOffset: -2,
      ...(lesson.needsWork ? {
        background: "#fdf8ef",
        borderLeft: "3px solid #d9a94e"
      } : null)
    },
    onClick: onOpen,
    onDragOver: drag ? e => {
      e.preventDefault();
      e.stopPropagation();
      setDragOver(true);
    } : undefined,
    onDragLeave: drag ? () => setDragOver(false) : undefined,
    onDrop: drag ? e => {
      e.preventDefault();
      e.stopPropagation();
      setDragOver(false);
      drag.onDrop();
    } : undefined
  }, /*#__PURE__*/React.createElement("div", {
    className: "row"
  }, drag && /*#__PURE__*/React.createElement("span", {
    draggable: true,
    className: "muted",
    title: "\u041F\u0435\u0440\u0435\u0442\u0430\u0449\u0438\u0442\u0435, \u0447\u0442\u043E\u0431\u044B \u0438\u0437\u043C\u0435\u043D\u0438\u0442\u044C \u043F\u043E\u0440\u044F\u0434\u043E\u043A \u0438\u043B\u0438 \u043F\u0435\u0440\u0435\u043D\u0435\u0441\u0442\u0438 \u0432 \u0434\u0440\u0443\u0433\u0443\u044E \u0441\u0438\u0441\u0442\u0435\u043C\u0443",
    style: {
      cursor: "grab",
      flex: "none",
      lineHeight: 0
    },
    onClick: e => e.stopPropagation(),
    onDragStart: e => {
      e.stopPropagation();
      try {
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/plain", "lesson");
      } catch (er) {/* не критично */}
      drag.onDragStart();
    },
    onDragEnd: e => {
      e.stopPropagation();
      drag.onDragEnd();
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "grip",
    size: 14
  })), onToggleDone && /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!lesson.completed,
    title: "\u0423\u0440\u043E\u043A \u043F\u0440\u043E\u0439\u0434\u0435\u043D \u2014 \u0443\u0439\u0434\u0451\u0442 \u0432 \u0430\u0440\u0445\u0438\u0432 (\u0438\u0437 \u0441\u043F\u0438\u0441\u043A\u0430), \u043D\u043E \u043E\u0441\u0442\u0430\u043D\u0435\u0442\u0441\u044F \u0432 \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u0435",
    onClick: e => e.stopPropagation(),
    onChange: () => onToggleDone(),
    style: {
      flex: "none",
      width: 17,
      height: 17,
      cursor: "pointer"
    }
  }), /*#__PURE__*/React.createElement("span", {
    className: "muted",
    style: {
      fontSize: 12
    }
  }, "\u2116", index + 1), lesson.plannedDate && /*#__PURE__*/React.createElement("span", {
    className: "chip",
    style: {
      flex: "none",
      padding: "2px 8px",
      fontSize: 11.5
    },
    title: "\u0414\u0430\u0442\u0430 \u0443\u0440\u043E\u043A\u0430 \u043F\u043E \u043F\u043B\u0430\u043D\u0443"
  }, lessonNoOf(lesson) ? lessonNoOf(lesson) + " ур. · " : "", weekdayShort(lesson.plannedDate), " ", formatDateRu(lesson.plannedDate).slice(0, 5)), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      fontWeight: 600,
      fontSize: 13.5
    }
  }, lesson.topic || "Без темы", lesson.duplicateOf && /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      fontWeight: 600
    }
  }, " (\u0434\u0443\u0431\u043B\u0438\u043A\u0430\u0442)"), lesson.needsWork && /*#__PURE__*/React.createElement("span", {
    style: {
      color: "#a8741a",
      fontWeight: 700,
      fontSize: 12
    }
  }, " \xB7 \u0442\u0440\u0435\u0431\u0443\u0435\u0442 \u0434\u043E\u0440\u0430\u0431\u043E\u0442\u043A\u0438")), onMove && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onMove(-1);
    }
  }, "\u2191"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onMove(1);
    }
  }, "\u2193")), onRemoveFromSystem && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: e => {
      e.stopPropagation();
      onRemoveFromSystem();
    },
    title: "\u0423\u0431\u0440\u0430\u0442\u044C \u0438\u0437 \u0441\u0438\u0441\u0442\u0435\u043C\u044B"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 12
  }))), /*#__PURE__*/React.createElement("div", {
    className: "qmeta"
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, lesson.items.length, " \u043C\u0430\u0442\u0435\u0440., ", lesson.completed ? "проведён" : "не проведён"), typeTags.map(t => /*#__PURE__*/React.createElement(Badge, {
    key: t.id,
    tag: t
  }))));
}
const WIDE_COLOR_PALETTE = [...SWATCHES, "#1E2530", "#697A80", "#C1554B", "#2A9D8F", "#B7791F", "#5C6BC0"];
function LessonSystemBlock({
  data,
  update,
  system,
  onOpenLesson,
  projectPath,
  dnd
}) {
  const patch = p => update(d => {
    const idx = d.lessonSystems.findIndex(s => s.id === system.id);
    if (idx >= 0) d.lessonSystems[idx] = {
      ...d.lessonSystems[idx],
      ...p
    };
    return d;
  });
  const setColor = color => update(d => {
    const idx = d.lessonSystems.findIndex(s => s.id === system.id);
    d.lessonSystems[idx] = {
      ...d.lessonSystems[idx],
      color
    };
    syncColorFromLessonSystem(d, d.lessonSystems[idx]);
    return d;
  });
  const setLink = (kind, id) => update(d => {
    const idx = d.lessonSystems.findIndex(s => s.id === system.id);
    const linkedEntity = kind ? (d[kind] || []).find(e => e.id === id) : null;
    d.lessonSystems[idx] = {
      ...d.lessonSystems[idx],
      colorLinkKind: kind || null,
      colorLinkId: id || null,
      color: linkedEntity && linkedEntity.color ? linkedEntity.color : d.lessonSystems[idx].color
    };
    return d;
  });
  const [copyOpen, setCopyOpen] = useState(false);
  const [copyChecked, setCopyChecked] = useState([]);
  const sysLetters = system.class && !system.courseId ? lettersForGrade(data, system.class) : [];
  const otherSysLetters = sysLetters.filter(cl => cl.id !== (system.classLetterId || null));
  // «Перенести в другие литеры»: создаёт в выбранных литерах такую же систему с независимыми
  // копиями всех её уроков (материалы и ДЗ — те же, даты и отметки — свои).
  const copySystemToLetters = () => {
    if (copyChecked.length === 0) return;
    update(d => {
      const src = d.lessonSystems.find(x => x.id === system.id);
      if (!src) return d;
      if (!src.copyGroupId) src.copyGroupId = uid();
      copyChecked.forEach(letterId => {
        const target = {
          class: src.class,
          courseId: src.courseId,
          classLetterId: letterId
        };
        const oldL = (d.classLetters.find(c => c.id === src.classLetterId) || {}).name || "";
        const newL = (d.classLetters.find(c => c.id === letterId) || {}).name || "";
        const title = oldL && newL ? src.title.replace(new RegExp(src.class + oldL, "i"), m => src.class + (m.slice(-1) === m.slice(-1).toUpperCase() ? newL : newL.toLowerCase())) : src.title;
        const sys = migrateLessonSystem({
          title,
          color: src.color,
          colorLinkKind: src.colorLinkKind,
          colorLinkId: src.colorLinkId,
          class: src.class,
          courseId: src.courseId,
          classLetterId: letterId,
          collapsed: src.collapsed,
          copyGroupId: src.copyGroupId,
          order: nextTopLevelOrder(d, target)
        });
        d.lessonSystems.push(sys);
        src.lessonIds.forEach(lid => {
          const newId = duplicateLessonInto(d, lid, {
            classLetterId: letterId,
            intoSystemId: sys.id
          });
          if (newId) sys.lessonIds.push(newId);
        });
      });
      return d;
    });
    showToast("Система скопирована в: " + otherSysLetters.filter(cl => copyChecked.includes(cl.id)).map(cl => system.class + cl.name).join(", "));
    setCopyOpen(false);
    setCopyChecked([]);
  };
  const addLessonLine = async () => {
    const topic = await PromptStr("Пункт плана (тема нового урока):");
    if (!topic || !topic.trim()) return;
    update(d => {
      const l = createLessonAuto(d, {
        class: system.class,
        courseId: system.courseId,
        classLetterId: system.classLetterId || null,
        systemId: system.id,
        topic: topic.trim()
      });
      const idx = d.lessonSystems.findIndex(s => s.id === system.id);
      d.lessonSystems[idx].lessonIds = [...d.lessonSystems[idx].lessonIds, l.id];
      return d;
    });
  };
  const toggleLessonCompleted = id => update(d => {
    const l = d.lessons.find(x => x.id === id);
    if (l) setLessonCompleted(l, !l.completed);
    return d;
  });
  const moveInSystem = (lessonId, dir) => update(d => {
    const idx = d.lessonSystems.findIndex(s => s.id === system.id);
    const ids = d.lessonSystems[idx].lessonIds.slice();
    const i = ids.indexOf(lessonId);
    let j = i + dir;
    // скрытые (архивные) уроки при перемещении перешагиваем
    while (i >= 0 && j >= 0 && j < ids.length) {
      const lj = d.lessons.find(x => x.id === ids[j]);
      if (lj && !lj.archived) break;
      j += dir;
    }
    if (i < 0 || j < 0 || j >= ids.length) return d;
    [ids[i], ids[j]] = [ids[j], ids[i]];
    d.lessonSystems[idx].lessonIds = ids;
    resyncSystemLessonsByDate(d, system.id);
    return d;
  });
  const removeFromSystem = lessonId => update(d => {
    const idx = d.lessonSystems.findIndex(s => s.id === system.id);
    d.lessonSystems[idx].lessonIds = d.lessonSystems[idx].lessonIds.filter(id => id !== lessonId);
    const l = d.lessons.find(x => x.id === lessonId);
    if (l) l.systemId = null;
    return d;
  });
  const deleteSystem = async () => {
    if (!(await Confirm("Удалить систему уроков? Сами уроки останутся, просто перестанут быть связаны."))) return;
    update(d => {
      d.lessons.forEach(l => {
        if (l.systemId === system.id) l.systemId = null;
      });
      d.lessonSystems = d.lessonSystems.filter(s => s.id !== system.id);
      return d;
    });
  };
  const linkedEntity = system.colorLinkKind ? (data[system.colorLinkKind] || []).find(e => e.id === system.colorLinkId) : null;
  const allSysLessons = system.lessonIds.map(id => data.lessons.find(l => l.id === id)).filter(Boolean);
  const doneCount = allSysLessons.filter(l => l.completed).length;
  const lessons = allSysLessons.filter(l => !l.archived);
  const [headerDragOver, setHeaderDragOver] = useState(false);
  return /*#__PURE__*/React.createElement("div", {
    style: {
      border: `2px solid ${system.color}`,
      borderRadius: "var(--radius)",
      marginBottom: 14,
      overflow: "hidden",
      background: "var(--surface)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      background: system.color,
      padding: "10px 14px",
      display: "flex",
      alignItems: "center",
      gap: 8,
      color: contrastTextColor(system.color),
      outline: headerDragOver ? "2px solid #fff" : "none",
      outlineOffset: -2
    },
    onDragOver: e => {
      if (dnd.dragLessonId || dnd.dragSystemId) {
        e.preventDefault();
        setHeaderDragOver(true);
      }
    },
    onDragLeave: () => setHeaderDragOver(false),
    onDrop: e => {
      e.preventDefault();
      setHeaderDragOver(false);
      if (dnd.dragLessonId) dnd.dropLessonAtEnd(system.id);else if (dnd.dragSystemId) dnd.dropSystemOn(system.id);
    }
  }, /*#__PURE__*/React.createElement("span", {
    draggable: true,
    title: "\u041F\u0435\u0440\u0435\u0442\u0430\u0449\u0438\u0442\u0435, \u0447\u0442\u043E\u0431\u044B \u0438\u0437\u043C\u0435\u043D\u0438\u0442\u044C \u043F\u043E\u0440\u044F\u0434\u043E\u043A \u0441\u0438\u0441\u0442\u0435\u043C \u0443\u0440\u043E\u043A\u043E\u0432",
    style: {
      cursor: "grab",
      lineHeight: 0
    },
    onDragStart: e => {
      e.stopPropagation();
      try {
        e.dataTransfer.effectAllowed = "move";
        e.dataTransfer.setData("text/plain", "system");
      } catch (er) {/* не критично */}
      dnd.startSystemDrag(system.id);
    },
    onDragEnd: e => {
      e.stopPropagation();
      dnd.endSystemDrag();
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "grip",
    size: 15
  })), /*#__PURE__*/React.createElement("button", {
    onClick: () => patch({
      collapsed: !system.collapsed
    }),
    style: {
      background: "transparent",
      border: "none",
      color: "inherit",
      cursor: "pointer"
    },
    title: system.collapsed ? "Показать уроки" : "Скрыть уроки"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "chev",
    size: 16
  })), /*#__PURE__*/React.createElement("input", {
    value: system.title,
    onChange: e => patch({
      title: e.target.value
    }),
    placeholder: "\u0421\u0438\u0441\u0442\u0435\u043C\u0430 \u0443\u0440\u043E\u043A\u043E\u0432\u2026",
    style: {
      flex: 1,
      background: "transparent",
      border: "none",
      fontFamily: "var(--font-display)",
      fontWeight: 700,
      fontSize: 15,
      color: "inherit"
    }
  }), allSysLessons.length > 0 && /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      opacity: 0.85,
      whiteSpace: "nowrap"
    },
    title: "\u041A\u043E\u0433\u0434\u0430 \u043F\u0440\u043E\u0439\u0434\u0435\u043D\u044B \u0432\u0441\u0435 \u0443\u0440\u043E\u043A\u0438 \u0441\u0438\u0441\u0442\u0435\u043C\u044B, \u043E\u043D\u0430 \u0443\u0445\u043E\u0434\u0438\u0442 \u0432 \u0430\u0440\u0445\u0438\u0432"
  }, "\u043F\u0440\u043E\u0439\u0434\u0435\u043D\u043E ", doneCount, " \u0438\u0437 ", allSysLessons.length), linkedEntity && /*#__PURE__*/React.createElement("span", {
    title: linkedEntity.title,
    style: {
      width: 22,
      height: 22,
      borderRadius: "50%",
      background: "rgba(255,255,255,.35)",
      display: "grid",
      placeItems: "center",
      fontSize: 11,
      flex: "none"
    }
  }, linkedEntity.title.slice(0, 1).toUpperCase()), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: deleteSystem,
    style: {
      color: "inherit"
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "10px 14px",
      borderBottom: "1px solid var(--line)",
      display: "flex",
      gap: 10,
      alignItems: "center",
      flexWrap: "wrap"
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 5
    }
  }, WIDE_COLOR_PALETTE.map(c => /*#__PURE__*/React.createElement("div", {
    key: c,
    className: "sw",
    style: {
      background: c,
      width: 18,
      height: 18
    },
    "aria-pressed": system.color === c ? "true" : undefined,
    onClick: () => setColor(c)
  })), /*#__PURE__*/React.createElement("input", {
    type: "color",
    value: system.color,
    onChange: e => setColor(e.target.value),
    title: "\u0421\u0432\u043E\u0439 \u0446\u0432\u0435\u0442",
    style: {
      width: 22,
      height: 22,
      border: "none",
      padding: 0,
      background: "none"
    }
  })), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: system.colorLinkKind ? `${system.colorLinkKind}:${system.colorLinkId}` : "",
    onChange: e => {
      const [kind, id] = e.target.value ? e.target.value.split(":") : [null, null];
      setLink(kind, id);
    }
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u041D\u0435 \u043F\u0440\u0438\u0432\u044F\u0437\u0430\u043D\u043E \u043A \u043F\u0435\u0440\u0441\u043E\u043D\u0430\u043B\u0438\u0438/\u0442\u0435\u043E\u0440\u0438\u0438"), /*#__PURE__*/React.createElement("optgroup", {
    label: "\u041F\u0435\u0440\u0441\u043E\u043D\u0430\u043B\u0438\u0438"
  }, data.personalities.map(p => /*#__PURE__*/React.createElement("option", {
    key: p.id,
    value: `personalities:${p.id}`
  }, p.title))), /*#__PURE__*/React.createElement("optgroup", {
    label: "\u0418\u0441\u0442\u043E\u0440\u0438\u044F \u0438 \u0442\u0435\u043E\u0440\u0438\u044F"
  }, data.theoryHistory.map(t => /*#__PURE__*/React.createElement("option", {
    key: t.id,
    value: `theoryHistory:${t.id}`
  }, t.title)))), sysLetters.length > 0 && /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: system.classLetterId || "",
    onChange: e => patch({
      classLetterId: e.target.value || null
    }),
    title: "\u0412 \u043A\u0430\u043A\u043E\u0439 \u043B\u0438\u0442\u0435\u0440\u0435 \u043F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u0435\u0442\u0441\u044F \u0441\u0438\u0441\u0442\u0435\u043C\u0430"
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u041E\u0431\u0449\u0430\u044F (\u0432\u0441\u0435 \u043B\u0438\u0442\u0435\u0440\u044B)"), sysLetters.map(cl => /*#__PURE__*/React.createElement("option", {
    key: cl.id,
    value: cl.id
  }, system.class, cl.name))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => {
      exportSystemFile(data, system);
      showToast("Система сохранена в файл .lsnsys");
    },
    title: "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u0441\u0438\u0441\u0442\u0435\u043C\u0443 \u0441\u043E \u0432\u0441\u0435\u043C\u0438 \u0443\u0440\u043E\u043A\u0430\u043C\u0438 \u0432 \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u044B\u0439 \u0444\u0430\u0439\u043B .lsnsys"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 12
  }), " \u042D\u043A\u0441\u043F\u043E\u0440\u0442 (.lsnsys)"), otherSysLetters.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => {
      setCopyOpen(!copyOpen);
      setCopyChecked([]);
    },
    title: "\u0421\u043A\u043E\u043F\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0441\u0438\u0441\u0442\u0435\u043C\u0443 \u0432\u043C\u0435\u0441\u0442\u0435 \u0441\u043E \u0432\u0441\u0435\u043C\u0438 \u0443\u0440\u043E\u043A\u0430\u043C\u0438 \u0432 \u0434\u0440\u0443\u0433\u0438\u0435 \u043B\u0438\u0442\u0435\u0440\u044B"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 12
  }), " \u041F\u0435\u0440\u0435\u043D\u0435\u0441\u0442\u0438 \u0432 \u0434\u0440\u0443\u0433\u0438\u0435 \u043B\u0438\u0442\u0435\u0440\u044B")), copyOpen && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      padding: "10px 14px",
      gap: 10,
      alignItems: "center",
      background: "var(--accent-soft)",
      borderBottom: "1px solid var(--line)"
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      fontWeight: 700
    }
  }, "\u0421\u043E\u0437\u0434\u0430\u0442\u044C \u0442\u0430\u043A\u0443\u044E \u0436\u0435 \u0441\u0438\u0441\u0442\u0435\u043C\u0443 \u0441 \u043A\u043E\u043F\u0438\u044F\u043C\u0438 \u0432\u0441\u0435\u0445 \u0443\u0440\u043E\u043A\u043E\u0432 \u0432 \u043B\u0438\u0442\u0435\u0440\u0430\u0445 (\u0434\u0430\u0442\u044B \u0438 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u043A\u0430\u0436\u0434\u043E\u0439 \u043B\u0438\u0442\u0435\u0440\u044B \u0441\u0432\u043E\u0438):"), otherSysLetters.map(cl => /*#__PURE__*/React.createElement("label", {
    key: cl.id,
    className: "row",
    style: {
      gap: 4,
      fontSize: 12.5,
      cursor: "pointer"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: copyChecked.includes(cl.id),
    onChange: e => setCopyChecked(prev => e.target.checked ? [...prev, cl.id] : prev.filter(x => x !== cl.id))
  }), system.class, cl.name)), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    disabled: copyChecked.length === 0,
    onClick: copySystemToLetters
  }, "\u041F\u0435\u0440\u0435\u043D\u0435\u0441\u0442\u0438")), !system.collapsed && /*#__PURE__*/React.createElement("div", {
    style: {
      padding: "10px 14px"
    },
    onDragOver: e => {
      if (dnd.dragLessonId) e.preventDefault();
    },
    onDrop: e => {
      e.preventDefault();
      if (dnd.dragLessonId) dnd.dropLessonAtEnd(system.id);
    }
  }, lessons.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 8
    }
  }, "\u041F\u0443\u0441\u0442\u043E \u2014 \u043F\u0435\u0440\u0435\u0442\u0430\u0449\u0438\u0442\u0435 \u0441\u044E\u0434\u0430 \u0443\u0440\u043E\u043A \u0438\u0437 \u0441\u043F\u0438\u0441\u043A\u0430 \u0438\u043B\u0438 \u0434\u043E\u0431\u0430\u0432\u044C\u0442\u0435 \u043F\u0443\u043D\u043A\u0442 \u043F\u043B\u0430\u043D\u0430 \u043D\u0438\u0436\u0435."), lessons.map((l, i) => /*#__PURE__*/React.createElement(LessonCard, {
    key: l.id,
    data: data,
    lesson: l,
    index: i,
    onOpen: () => onOpenLesson(l.id),
    onRemoveFromSystem: () => removeFromSystem(l.id),
    onToggleDone: () => toggleLessonCompleted(l.id),
    onMove: dir => moveInSystem(l.id, dir),
    drag: {
      onDragStart: () => dnd.startLessonDrag(l.id, system.id),
      onDragEnd: dnd.endLessonDrag,
      onDrop: () => dnd.dropLessonBefore(l.id, system.id)
    }
  })), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: addLessonLine
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u041F\u0443\u043D\u043A\u0442 \u043F\u043B\u0430\u043D\u0430 / \u0443\u0440\u043E\u043A")));
}

// Куда отправлять: на ноутбуке мессенджер открывается с готовым текстом, а файл сохраняется и показывается в папке —
// его остаётся перетащить в чат (мессенджеры не дают программам прикреплять файлы сами). На планшете — системное «Поделиться».
const MESSENGERS = [{
  id: "telegram",
  label: "Telegram",
  url: t => "https://t.me/share/url?url=%20&text=" + encodeURIComponent(t)
}, {
  id: "whatsapp",
  label: "WhatsApp",
  url: t => "https://wa.me/?text=" + encodeURIComponent(t)
}, {
  id: "viber",
  label: "Viber",
  url: t => "viber://forward?text=" + encodeURIComponent(t)
}, {
  id: "vk",
  label: "ВКонтакте",
  url: () => "https://vk.com/im"
}, {
  id: "max",
  label: "MAX",
  url: () => "https://web.max.ru/"
}];
// Окно «Отправить задание»: уроки → материалы (упражнения, карточки с вопросами, домашнее задание) → срок → файл Word → мессенджер.
// Каждый материал можно ПРАВИТЬ: правка касается только отправляемого файла (копия текста лежит в этом окне),
// банк упражнений, карточки и сами уроки не меняются.
function SendLessonsModal({
  data,
  lessons,
  groupLabel,
  onClose
}) {
  const [sel, setSel] = useState({}); // lessonId → { blocks: {blockKey: bool}, hw: bool, due }
  const [edits, setEdits] = useState({}); // lessonId + "|" + blockKey → правленый HTML (только для отправки)
  const [editing, setEditing] = useState(null); // ключ правки, открытой сейчас
  const [space, setSpace] = useState(true);
  const [common, setCommon] = useState("");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const SIZE = 15;
  const blocksOf = l => lessonItemBlocks(l.items || []).filter(b => b.kind === "exercise" || b.kind === "group" && questionSourceEntity(data, {
    entityType: b.entityType,
    entityId: b.entityId
  }));
  const entOf = b => {
    const type = b.kind === "exercise" ? "exercise" : b.entityType;
    const id = b.kind === "exercise" ? b.refId : b.entityId;
    return {
      type,
      ent: (data[ITEM_TYPE_META[type].dataKey] || []).find(e => e.id === id)
    };
  };
  const blockTitle = b => {
    const {
      type,
      ent
    } = entOf(b);
    return ent ? ITEM_TYPE_META[type].titleOf(data, ent) : "(удалено)";
  };
  const blockMeta = b => (b.kind === "exercise" ? "упражнение" : ITEM_TYPE_META[b.entityType].label.toLowerCase()) + (b.questions && b.questions.length ? ", вопросов: " + b.questions.length : b.kind === "exercise" && b.items.length > 1 ? ", вопросов: " + (b.items.length - 1) : "");
  const baseHtml = (l, b) => {
    const {
      type,
      ent
    } = entOf(b);
    if (!ent) return "";
    const withText = b.kind === "exercise" || !!b.whole;
    return buildMaterialPrintHtml(data, type, ent, {
      text: withText,
      questions: true,
      answers: false,
      space,
      sticky: false,
      size: SIZE
    }, l);
  };
  const hwBase = l => `<div style="font-family:Arial,sans-serif;font-size:${SIZE}px;line-height:1.5"><h2 style="font-size:${Math.round(SIZE * 1.2)}px">Домашнее задание</h2>${cleanPrintHtml(l.homework || "")}</div>`;
  const htmlOf = (l, key, b) => {
    const k = l.id + "|" + key;
    if (edits[k] !== undefined) return edits[k];
    return key === "hw" ? hwBase(l) : baseHtml(l, b);
  };
  const toggleLesson = l => setSel(p => {
    const n = {
      ...p
    };
    if (n[l.id]) delete n[l.id];else n[l.id] = {
      blocks: Object.fromEntries(blocksOf(l).map(b => [b.key, true])),
      hw: false,
      due: common
    };
    return n;
  });
  const patch = (id, ch) => setSel(p => ({
    ...p,
    [id]: {
      ...p[id],
      ...ch
    }
  }));
  const chosen = lessons.filter(l => sel[l.id]);
  const taskCount = chosen.reduce((n, l) => n + Object.values(sel[l.id].blocks).filter(Boolean).length + (sel[l.id].hw ? 1 : 0), 0);
  const dueText = iso => iso ? formatDateRu(iso) + " (" + ["воскресенье", "понедельник", "вторник", "среда", "четверг", "пятница", "суббота"][new Date(iso + "T00:00:00").getDay()] + ")" : "";
  const buildBody = () => {
    const summary = chosen.map(l => `<tr><td style="padding:3px 10px;border:1px solid #bbb;">${escapeHtml(l.topic || "Без темы")}</td><td style="padding:3px 10px;border:1px solid #bbb;">${sel[l.id].due ? escapeHtml(dueText(sel[l.id].due)) : "—"}</td></tr>`).join("");
    let body = `<table style="border-collapse:collapse;margin:8px 0 14px;"><tr><th style="padding:3px 10px;border:1px solid #bbb;text-align:left;">Тема</th><th style="padding:3px 10px;border:1px solid #bbb;text-align:left;">Выполнить до</th></tr>${summary}</table>`;
    chosen.forEach(l => {
      const st = sel[l.id];
      body += `<h2>${escapeHtml(l.topic || "Без темы")}</h2>`;
      if (st.due) body += `<div style="margin-bottom:6px;"><b>Выполнить до: ${escapeHtml(dueText(st.due))}</b></div>`;
      blocksOf(l).forEach(b => {
        if (st.blocks[b.key]) body += `<div style="margin:10px 0 18px;">${htmlOf(l, b.key, b)}</div>`;
      });
      if (st.hw && (l.homework || "").trim()) body += `<div style="margin:10px 0 18px;">${htmlOf(l, "hw", null)}</div>`;
    });
    return body;
  };
  const messageText = () => {
    const dues = chosen.filter(l => sel[l.id].due).map(l => sel[l.id].due).sort();
    return `${groupLabel}: задание. Уроков: ${chosen.length}${dues.length ? ". Выполнить до " + formatDateRu(dues[0]) : ""}. Файл Word — во вложении.`;
  };
  const send = async m => {
    if (!chosen.length) return;
    if (!taskCount) {
      showToast("Не выбрано ни одного задания");
      return;
    }
    setBusy(true);
    setMsg("");
    try {
      const title = "Задание — " + groupLabel;
      const html = wordDocFullHtml(title, await inlineLitfileImages(buildBody()), {});
      const fileName = safeFileName(title + " " + todayIso()) + ".doc";
      const text = messageText();
      if (!window.litAPI || !window.litAPI.shareDocument) {
        downloadWord(fileName, title, await inlineLitfileImages(buildBody()), {});
        setMsg("Файл сохранён.");
      } else {
        const res = await window.litAPI.shareDocument({
          fileName,
          html,
          text,
          url: m ? m.url(text) : null
        });
        if (res.ok) setMsg(window.litTabletMode ? "Готово." : "Файл сохранён: " + res.path + ". Папка открыта — перетащите файл в чат" + (m ? " " + m.label : "") + ".");else if (!res.cancelled) setMsg("Не получилось: " + (res.error || "неизвестная ошибка"));
      }
    } catch (e) {
      setMsg("Не получилось: " + (e && e.message || e));
    }
    setBusy(false);
  };
  const inp = {
    border: "1px solid var(--line)",
    borderRadius: 8,
    padding: "6px 9px"
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 860
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u041E\u0442\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u0437\u0430\u0434\u0430\u043D\u0438\u0435 \u2014 ", groupLabel), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      display: "grid",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041E\u0442\u043C\u0435\u0442\u044C\u0442\u0435 \u0443\u0440\u043E\u043A\u0438, \u0432 \u043A\u0430\u0436\u0434\u043E\u043C \u2014 \u043D\u0443\u0436\u043D\u044B\u0435 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B \u0438 \u0441\u0440\u043E\u043A. \u0418\u0437 \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u043E\u0433\u043E \u0444\u043E\u0440\u043C\u0438\u0440\u0443\u0435\u0442\u0441\u044F \u0444\u0430\u0439\u043B Word (\u0431\u0435\u0437 \u0432\u0430\u0448\u0438\u0445 \u0437\u0430\u043C\u0435\u0442\u043E\u043A), \u0435\u0433\u043E \u043C\u043E\u0436\u043D\u043E \u0441\u0440\u0430\u0437\u0443 \u043E\u0442\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u0432 \u043C\u0435\u0441\u0441\u0435\u043D\u0434\u0436\u0435\u0440. \u041B\u044E\u0431\u043E\u0439 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B \u043C\u043E\u0436\u043D\u043E \u043F\u043E\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u043A\u043D\u043E\u043F\u043A\u043E\u0439 \xAB\u270E \u041F\u0440\u0430\u0432\u0438\u0442\u044C\xBB \u2014 ", /*#__PURE__*/React.createElement("b", null, "\u043F\u0440\u0430\u0432\u043A\u0430 \u043A\u0430\u0441\u0430\u0435\u0442\u0441\u044F \u0442\u043E\u043B\u044C\u043A\u043E \u043E\u0442\u043F\u0440\u0430\u0432\u043B\u044F\u0435\u043C\u043E\u0433\u043E \u0444\u0430\u0439\u043B\u0430"), ", \u0431\u0430\u043D\u043A \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439 \u0438 \u0441\u0430\u043C\u0438 \u0443\u0440\u043E\u043A\u0438 \u043D\u0435 \u043C\u0435\u043D\u044F\u044E\u0442\u0441\u044F."), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 10,
      fontSize: 12.5
    }
  }, /*#__PURE__*/React.createElement("span", null, "\u041E\u0431\u0449\u0438\u0439 \u0441\u0440\u043E\u043A \u0434\u043B\u044F \u043D\u043E\u0432\u044B\u0445 \u043E\u0442\u043C\u0435\u0442\u043E\u043A:"), /*#__PURE__*/React.createElement("input", {
    type: "date",
    style: inp,
    value: common,
    onChange: e => setCommon(e.target.value)
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !chosen.length || !common,
    onClick: () => setSel(p => Object.fromEntries(Object.entries(p).map(([k, v]) => [k, {
      ...v,
      due: common
    }])))
  }, "\u041F\u043E\u0441\u0442\u0430\u0432\u0438\u0442\u044C \u0432\u0441\u0435\u043C \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u044B\u043C"), /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 6,
      cursor: "pointer"
    },
    title: "\u041F\u043E\u0441\u043B\u0435 \u043A\u0430\u0436\u0434\u043E\u0433\u043E \u0432\u043E\u043F\u0440\u043E\u0441\u0430 \u0434\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u043B\u0438\u043D\u0438\u0438 \u0434\u043B\u044F \u043F\u0438\u0441\u044C\u043C\u0435\u043D\u043D\u043E\u0433\u043E \u043E\u0442\u0432\u0435\u0442\u0430"
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: space,
    onChange: e => setSpace(e.target.checked)
  }), " \u041C\u0435\u0441\u0442\u043E \u0434\u043B\u044F \u043E\u0442\u0432\u0435\u0442\u043E\u0432 \u043F\u043E\u0434 \u0432\u043E\u043F\u0440\u043E\u0441\u0430\u043C\u0438")), lessons.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0412 \u044D\u0442\u043E\u043C \u043A\u043B\u0430\u0441\u0441\u0435 \u043D\u0435\u0442 \u0443\u0440\u043E\u043A\u043E\u0432 \u0434\u043B\u044F \u043E\u0442\u043F\u0440\u0430\u0432\u043A\u0438."), lessons.map(l => {
    const st = sel[l.id];
    const bl = blocksOf(l);
    return /*#__PURE__*/React.createElement("div", {
      key: l.id,
      className: "panel",
      style: st ? {
        borderColor: "var(--accent)"
      } : undefined
    }, /*#__PURE__*/React.createElement("div", {
      className: "panel-body",
      style: {
        display: "grid",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("label", {
      className: "row",
      style: {
        gap: 8,
        cursor: "pointer",
        fontWeight: 700
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: !!st,
      onChange: () => toggleLesson(l)
    }), /*#__PURE__*/React.createElement("span", {
      style: {
        flex: 1
      }
    }, l.topic || "Без темы", " ", /*#__PURE__*/React.createElement("span", {
      className: "hint",
      style: {
        fontWeight: 400
      }
    }, l.plannedDate ? "· план: " + formatDateRu(l.plannedDate) : "", " \xB7 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u043E\u0432: ", bl.length))), st && /*#__PURE__*/React.createElement("div", {
      style: {
        marginLeft: 26,
        display: "grid",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        gap: 8,
        fontSize: 12.5
      }
    }, "\u0412\u044B\u043F\u043E\u043B\u043D\u0438\u0442\u044C \u0434\u043E:", /*#__PURE__*/React.createElement("input", {
      type: "date",
      style: inp,
      value: st.due,
      onChange: e => patch(l.id, {
        due: e.target.value
      })
    }), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => patch(l.id, {
        blocks: Object.fromEntries(bl.map(b => [b.key, true]))
      })
    }, "\u0432\u0441\u0435"), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => patch(l.id, {
        blocks: Object.fromEntries(bl.map(b => [b.key, false]))
      })
    }, "\u0441\u043D\u044F\u0442\u044C \u0432\u0441\u0435")), bl.map(b => {
      const ek = l.id + "|" + b.key;
      return /*#__PURE__*/React.createElement("div", {
        key: b.key,
        style: {
          display: "grid",
          gap: 4
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "row",
        style: {
          gap: 8,
          fontSize: 12.5,
          alignItems: "flex-start"
        }
      }, /*#__PURE__*/React.createElement("input", {
        type: "checkbox",
        checked: !!st.blocks[b.key],
        onChange: e => patch(l.id, {
          blocks: {
            ...st.blocks,
            [b.key]: e.target.checked
          }
        }),
        style: {
          marginTop: 3
        }
      }), /*#__PURE__*/React.createElement("span", {
        style: {
          flex: 1
        }
      }, /*#__PURE__*/React.createElement("b", null, blockTitle(b)), " ", /*#__PURE__*/React.createElement("span", {
        className: "hint"
      }, "\xB7 ", blockMeta(b)), edits[ek] !== undefined && /*#__PURE__*/React.createElement("span", {
        style: {
          color: "var(--accent)",
          fontWeight: 700
        }
      }, " \xB7 \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u043E \u0434\u043B\u044F \u043E\u0442\u043F\u0440\u0430\u0432\u043A\u0438")), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => setEditing(editing === ek ? null : ek)
      }, editing === ek ? "Готово" : "✎ Править"), edits[ek] !== undefined && /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => {
          setEdits(p => {
            const n = {
              ...p
            };
            delete n[ek];
            return n;
          });
          setEditing(null);
        },
        title: "\u0412\u0435\u0440\u043D\u0443\u0442\u044C \u0442\u0435\u043A\u0441\u0442 \u043A\u0430\u043A \u0432 \u0431\u0430\u043D\u043A\u0435"
      }, "\u21BA \u043A\u0430\u043A \u0432 \u0431\u0430\u043D\u043A\u0435")), editing === ek && /*#__PURE__*/React.createElement("div", {
        style: {
          marginLeft: 22
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "hint",
        style: {
          marginBottom: 4
        }
      }, "\u041F\u0440\u0430\u0432\u044C\u0442\u0435 \u043A\u0430\u043A \u043E\u0431\u044B\u0447\u043D\u044B\u0439 \u0442\u0435\u043A\u0441\u0442. \u042D\u0442\u043E \u043A\u043E\u043F\u0438\u044F \u0442\u043E\u043B\u044C\u043A\u043E \u0434\u043B\u044F \u043E\u0442\u043F\u0440\u0430\u0432\u043A\u0438 \u2014 \u0432 \u0431\u0430\u043D\u043A\u0435 \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u0438\u0437\u043C\u0435\u043D\u0438\u0442\u0441\u044F."), /*#__PURE__*/React.createElement(RichEditor, {
        key: ek,
        html: htmlOf(l, b.key, b),
        onChange: h => setEdits(p => ({
          ...p,
          [ek]: h
        })),
        compact: true,
        projectPath: null
      })));
    }), (l.homework || "").trim() && (() => {
      const ek = l.id + "|hw";
      return /*#__PURE__*/React.createElement("div", {
        style: {
          display: "grid",
          gap: 4
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "row",
        style: {
          gap: 8,
          fontSize: 12.5
        }
      }, /*#__PURE__*/React.createElement("input", {
        type: "checkbox",
        checked: st.hw,
        onChange: e => patch(l.id, {
          hw: e.target.checked
        })
      }), /*#__PURE__*/React.createElement("span", {
        style: {
          flex: 1
        }
      }, /*#__PURE__*/React.createElement("b", null, "\u0414\u043E\u043C\u0430\u0448\u043D\u0435\u0435 \u0437\u0430\u0434\u0430\u043D\u0438\u0435 \u0443\u0440\u043E\u043A\u0430"), edits[ek] !== undefined && /*#__PURE__*/React.createElement("span", {
        style: {
          color: "var(--accent)",
          fontWeight: 700
        }
      }, " \xB7 \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u043E \u0434\u043B\u044F \u043E\u0442\u043F\u0440\u0430\u0432\u043A\u0438")), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => setEditing(editing === ek ? null : ek)
      }, editing === ek ? "Готово" : "✎ Править"), edits[ek] !== undefined && /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => {
          setEdits(p => {
            const n = {
              ...p
            };
            delete n[ek];
            return n;
          });
          setEditing(null);
        }
      }, "\u21BA \u043A\u0430\u043A \u0432 \u0443\u0440\u043E\u043A\u0435")), editing === ek && /*#__PURE__*/React.createElement("div", {
        style: {
          marginLeft: 22
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "hint",
        style: {
          marginBottom: 4
        }
      }, "\u041A\u043E\u043F\u0438\u044F \u0442\u043E\u043B\u044C\u043A\u043E \u0434\u043B\u044F \u043E\u0442\u043F\u0440\u0430\u0432\u043A\u0438 \u2014 \u0434\u043E\u043C\u0430\u0448\u043D\u0435\u0435 \u0437\u0430\u0434\u0430\u043D\u0438\u0435 \u0432 \u0443\u0440\u043E\u043A\u0435 \u043D\u0435 \u0438\u0437\u043C\u0435\u043D\u0438\u0442\u0441\u044F."), /*#__PURE__*/React.createElement(RichEditor, {
        key: ek,
        html: htmlOf(l, "hw", null),
        onChange: h => setEdits(p => ({
          ...p,
          [ek]: h
        })),
        compact: true,
        projectPath: null
      })));
    })())));
  })), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot",
    style: {
      display: "grid",
      gap: 8,
      justifyContent: "stretch"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13
    }
  }, "\u0412\u044B\u0431\u0440\u0430\u043D\u043E \u0443\u0440\u043E\u043A\u043E\u0432: ", /*#__PURE__*/React.createElement("b", null, chosen.length), ", \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u043E\u0432: ", /*#__PURE__*/React.createElement("b", null, taskCount)), window.litTabletMode ? /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: busy || !taskCount,
    onClick: () => send(null)
  }, "\uD83D\uDCE4 \u041E\u0442\u043F\u0440\u0430\u0432\u0438\u0442\u044C\u2026 (\u0432\u044B\u0431\u0440\u0430\u0442\u044C \u043C\u0435\u0441\u0441\u0435\u043D\u0434\u0436\u0435\u0440)"), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 4
    }
  }, "\u041E\u0442\u043A\u0440\u043E\u0435\u0442\u0441\u044F \u0441\u0438\u0441\u0442\u0435\u043C\u043D\u043E\u0435 \u043C\u0435\u043D\u044E \xAB\u041F\u043E\u0434\u0435\u043B\u0438\u0442\u044C\u0441\u044F\xBB \u2014 \u0432\u044B\u0431\u0435\u0440\u0438\u0442\u0435 Telegram, WhatsApp, Viber, \u0412\u041A\u043E\u043D\u0442\u0430\u043A\u0442\u0435 \u0438\u043B\u0438 MAX, \u0438 \u0444\u0430\u0439\u043B \u0443\u0439\u0434\u0451\u0442 \u0441\u0440\u0430\u0437\u0443.")) : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6
    }
  }, MESSENGERS.map(m => /*#__PURE__*/React.createElement(Btn, {
    key: m.id,
    kind: "primary",
    sm: true,
    disabled: busy || !taskCount,
    onClick: () => send(m)
  }, m.label)), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: busy || !taskCount,
    onClick: () => send(null)
  }, "\u0422\u043E\u043B\u044C\u043A\u043E \u0441\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u0444\u0430\u0439\u043B")), /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041D\u0430\u0436\u043C\u0438\u0442\u0435 \u043D\u0430 \u043C\u0435\u0441\u0441\u0435\u043D\u0434\u0436\u0435\u0440: \u0444\u0430\u0439\u043B Word \u0441\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u0441\u044F (\u0432\u044B\u0431\u0435\u0440\u0435\u0442\u0435 \u043F\u0430\u043F\u043A\u0443), \u044D\u0442\u0430 \u043F\u0430\u043F\u043A\u0430 \u043E\u0442\u043A\u0440\u043E\u0435\u0442\u0441\u044F \u0441 \u0432\u044B\u0434\u0435\u043B\u0435\u043D\u043D\u044B\u043C \u0444\u0430\u0439\u043B\u043E\u043C, \u0430 \u043C\u0435\u0441\u0441\u0435\u043D\u0434\u0436\u0435\u0440 \u043E\u0442\u043A\u0440\u043E\u0435\u0442\u0441\u044F \u0441 \u0433\u043E\u0442\u043E\u0432\u044B\u043C \u0442\u0435\u043A\u0441\u0442\u043E\u043C \u2014 \u043E\u0441\u0442\u0430\u043D\u0435\u0442\u0441\u044F \u043F\u0435\u0440\u0435\u0442\u0430\u0449\u0438\u0442\u044C \u0444\u0430\u0439\u043B \u0432 \u0447\u0430\u0442.")), msg && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12.5,
      color: msg.startsWith("Не получилось") ? "var(--neg)" : "var(--pos, #2e8b57)"
    }
  }, msg))));
}

// Разархивирование урока: предлагается дата, на которую он встанет (по расписанию или после последнего урока группы); можно изменить.
function UnarchiveModal({
  data,
  update,
  lessonId,
  onClose
}) {
  const lesson = data.lessons.find(l => l.id === lessonId);
  const [date, setDate] = useState(() => lesson ? suggestUnarchiveDate(data, lesson) : todayIso());
  if (!lesson) return null;
  const no = date ? freeScheduleSlot(data, lesson, date) : null;
  const vac = date ? vacationAt(data, date) : null;
  const confirm = () => {
    if (!date) {
      showToast("Укажите дату");
      return;
    }
    update(d => {
      unarchiveLesson(d, lessonId, date);
      return d;
    });
    showToast("Урок возвращён в список на " + formatDateRu(date));
    onClose(true);
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: () => onClose(false)
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 460
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u0420\u0430\u0437\u0430\u0440\u0445\u0438\u0432\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0443\u0440\u043E\u043A"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => onClose(false)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      display: "grid",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700
    }
  }, lesson.topic || "Без темы", " ", /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\xB7 ", lessonPlaceLabel(lesson, data))), /*#__PURE__*/React.createElement(Field, {
    label: "\u0423\u0440\u043E\u043A \u0432\u0441\u0442\u0430\u043D\u0435\u0442 \u043D\u0430 \u0434\u0430\u0442\u0443"
  }, /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: date,
    onChange: e => setDate(e.target.value),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px"
    }
  })), /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, date ? weekdayShort(date) + ", " + formatDateRu(date) : "", no ? " · " + no + " урок по вашему расписанию" : "", vac ? /*#__PURE__*/React.createElement("span", {
    style: {
      color: "var(--neg)"
    }
  }, " \xB7 \u043D\u0430 \u044D\u0442\u0438 \u0434\u0430\u0442\u044B \u043F\u0440\u0438\u0445\u043E\u0434\u044F\u0442\u0441\u044F \xAB", vac.name, "\xBB") : null), /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041F\u0440\u0435\u0434\u043B\u043E\u0436\u0435\u043D\u0430 \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0430\u044F \u0441\u0432\u043E\u0431\u043E\u0434\u043D\u0430\u044F \u0434\u0430\u0442\u0430", (data.schedule && data.schedule.slots || []).some(x => lessonMatchesSlot(lesson, x)) ? " по «Моему расписанию»" : " после последнего урока класса", ". \u041C\u043E\u0436\u043D\u043E \u0432\u044B\u0431\u0440\u0430\u0442\u044C \u043B\u044E\u0431\u0443\u044E \u0434\u0440\u0443\u0433\u0443\u044E.")), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => onClose(false)
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: confirm
  }, "\u21A9 \u0420\u0430\u0437\u0430\u0440\u0445\u0438\u0432\u0438\u0440\u043E\u0432\u0430\u0442\u044C"))));
}
function ReservePool({
  data,
  update,
  groupKind,
  groupValue
}) {
  const items = data.reserves.filter(r => r.groupKind === groupKind && r.groupValue === groupValue);
  if (items.length === 0) return null;
  const addToLesson = (reserveId, lessonId) => update(d => {
    const r = d.reserves.find(x => x.id === reserveId);
    if (!r) return d;
    const l = d.lessons.find(x => x.id === lessonId);
    if (l && !l.items.some(it => it.type === "exercise" && it.refId === r.exerciseId)) l.items.push({
      id: uid(),
      type: "exercise",
      refId: r.exerciseId,
      note: ""
    });
    d.reserves = d.reserves.filter(x => x.id !== reserveId);
    return d;
  });
  const removeReserve = id => update(d => {
    d.reserves = d.reserves.filter(r => r.id !== id);
    return d;
  });
  const lessonsHere = data.lessons.filter(l => (groupKind === "class" ? l.class === groupValue : l.courseId === groupValue) && !l.archived);
  return /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-head"
  }, /*#__PURE__*/React.createElement("h2", null, "\u0417\u0430\u043F\u0430\u0441\u043D\u044B\u0435 \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u044F")), /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, items.map(r => {
    const ex = data.examples.find(e => e.id === r.exerciseId);
    return /*#__PURE__*/React.createElement("div", {
      key: r.id,
      className: "row",
      style: {
        padding: "6px 0",
        borderBottom: "1px solid var(--line)"
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        fontSize: 13
      }
    }, ex ? stripHtml(ex.text).slice(0, 80) : "(удалено)"), /*#__PURE__*/React.createElement("select", {
      className: "toolselect",
      onChange: e => {
        if (e.target.value) addToLesson(r.id, e.target.value);
        e.target.value = "";
      }
    }, /*#__PURE__*/React.createElement("option", {
      value: ""
    }, "\u0412 \u0443\u0440\u043E\u043A\u2026"), lessonsHere.map(l => /*#__PURE__*/React.createElement("option", {
      key: l.id,
      value: l.id
    }, l.topic || "Без темы"))), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => removeReserve(r.id)
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "trash",
      size: 12
    })));
  })));
}
function LessonsTab({
  data,
  update,
  projectPath,
  initialLessonId,
  onConsumedInitial,
  nav,
  setNav,
  originBack,
  compact,
  journalOpen,
  onToggleJournal
}) {
  const importRef = useRef(null);
  const importSysRef = useRef(null);
  const target = nav.lessons.target;
  const letterFilter = nav.lessons.letterFilter;
  const setTarget = v => setNav(n => ({
    ...n,
    lessons: {
      ...n.lessons,
      target: v
    }
  }));
  const setLetterFilter = v => setNav(n => ({
    ...n,
    lessons: {
      ...n.lessons,
      letterFilter: v
    }
  }));
  const openLessonId = nav.lessons.openId || null;
  const setOpenLessonId = v => setNav(n => ({
    ...n,
    lessons: {
      ...n.lessons,
      openId: v
    }
  }));
  // Drag-and-drop: перетаскиваемый урок (и система, из которой он взят — null, если был отдельным),
  // и отдельно — перетаскиваемая система уроков (порядок систем реорганизуется независимо).
  const [dragLessonId, setDragLessonId] = useState(null);
  const [dragLessonFromSystemId, setDragLessonFromSystemId] = useState(null);
  const [dragSystemId, setDragSystemId] = useState(null);
  useEffect(() => {
    if (initialLessonId) {
      setOpenLessonId(initialLessonId);
      onConsumedInitial && onConsumedInitial();
    }
  }, [initialLessonId]);
  const ltFirstRender = useRef(true);
  useEffect(() => {
    if (ltFirstRender.current) {
      ltFirstRender.current = false;
      return;
    }
    if (!letterValidFor(data, target.kind, target.value, letterFilter)) setLetterFilter("all");
  }, [target.kind, target.value]);
  const addCourse = async () => {
    const name = await PromptStr("Название курса:");
    if (!name || !name.trim()) return;
    const c = {
      id: uid(),
      name: name.trim(),
      color: SWATCHES[Math.floor(Math.random() * SWATCHES.length)]
    };
    update(d => {
      d.courses.push(c);
      return d;
    });
    setTarget({
      kind: "course",
      value: c.id
    });
  };
  const groupKind = target.kind;
  const groupValue = target.value;
  const letters = groupKind === "class" ? lettersForGrade(data, groupValue) : [];
  const wantLetterId = groupKind === "class" && letterFilter !== "all" && letterFilter !== "none" ? letterFilter : null;
  const lessonsHere = data.lessons.filter(l => groupKind === "class" ? l.class === groupValue : l.courseId === groupValue).filter(l => groupKind !== "class" || letterFilter === "all" || (letterFilter === "none" ? !l.classLetterId : l.classLetterId === letterFilter));
  const standaloneLessons = lessonsHere.filter(l => !l.systemId && !l.archived).sort((a, b) => a.order - b.order);
  const archivedHere = lessonsHere.filter(l => l.archived).sort((a, b) => (a.plannedDate || "9999") < (b.plannedDate || "9999") ? -1 : (a.plannedDate || "9999") > (b.plannedDate || "9999") ? 1 : a.order - b.order);
  const [archOpen, setArchOpen] = useState(false);
  const [sendOpen, setSendOpen] = useState(false);
  const [unarchiveId, setUnarchiveId] = useState(null);
  const systemsHere = data.lessonSystems.filter(s => !s.archived).filter(s => groupKind === "class" ? s.class === groupValue : s.courseId === groupValue).filter(s => groupKind !== "class" || letterFilter === "all" || !s.classLetterId || s.classLetterId === letterFilter).sort((a, b) => (a.order || 0) - (b.order || 0));
  // Что именно сейчас показано в списке — по этому «срезу» и пересчитывается общий порядок блоков.
  const viewSample = {
    class: groupKind === "class" ? groupValue : null,
    courseId: groupKind === "course" ? groupValue : null,
    classLetterId: wantLetterId,
    anyLetter: groupKind === "class" && letterFilter === "all"
  };
  const standaloneIndexById = new Map(standaloneLessons.map((l, i) => [l.id, i]));
  // Отдельные уроки и системы вперемешку, в общем порядке — так систему можно "обойти"
  // уроком до неё, между её частями по порядку в списке или после.
  const topLevelCombined = [...standaloneLessons.map(l => ({
    type: "lesson",
    item: l,
    order: l.order
  })), ...systemsHere.map(s => ({
    type: "system",
    item: s,
    order: s.order || 0
  }))].sort((a, b) => a.order - b.order);
  const handleImportLessons = async files => {
    for (const f of Array.from(files)) {
      let payload = null;
      try {
        payload = JSON.parse((await f.text()).replace(/^\uFEFF/, ""));
      } catch (e) {/* не JSON */}
      if (!payload || payload.type !== "literature-constructor-lesson" || payload.app !== "literature") {
        showToast("«" + f.name + "» — не файл урока конструктора литературы (.lsnlit)");
        continue;
      }
      update(d => {
        importLessonPayload(d, payload, {
          kind: groupKind,
          value: groupValue,
          classLetterId: wantLetterId
        });
        return d;
      });
      showToast("Урок «" + (payload.topic || "без темы") + "» добавлен в список");
    }
  };
  const handleImportSystems = async files => {
    for (const f of Array.from(files)) {
      let payload = null;
      try {
        payload = JSON.parse((await f.text()).replace(/^\uFEFF/, ""));
      } catch (e) {/* не JSON */}
      if (!payload || payload.type !== "literature-constructor-system" || payload.app !== "literature") {
        showToast("«" + f.name + "» — не файл системы уроков конструктора литературы (.lsnsys)");
        continue;
      }
      update(d => {
        importSystemPayload(d, payload, {
          kind: groupKind,
          value: groupValue,
          classLetterId: wantLetterId
        });
        return d;
      });
      showToast("Система «" + (payload.title || "без названия") + "» (" + (payload.lessons || []).length + " ур.) добавлена в список");
    }
  };
  const doAddLesson = () => {
    let newId = null;
    update(d => {
      const l = createLessonAuto(d, {
        class: groupKind === "class" ? groupValue : null,
        courseId: groupKind === "course" ? groupValue : null,
        classLetterId: wantLetterId
      });
      newId = l.id;
      return d;
    });
    setTimeout(() => setOpenLessonId(newId), 0);
  };
  const addSystem = async () => {
    const title = await PromptStr("Название системы уроков (план):");
    if (!title || !title.trim()) return;
    update(d => {
      const sample = {
        class: groupKind === "class" ? groupValue : null,
        courseId: groupKind === "course" ? groupValue : null,
        classLetterId: wantLetterId
      };
      d.lessonSystems.push(migrateLessonSystem({
        title: title.trim(),
        class: groupKind === "class" ? groupValue : null,
        courseId: groupKind === "course" ? groupValue : null,
        classLetterId: wantLetterId,
        color: SWATCHES[Math.floor(Math.random() * SWATCHES.length)],
        order: nextTopLevelOrder(d, sample)
      }));
      return d;
    });
  };
  // Стрелки ↑/↓ у отдельного урока — двигают его среди ВСЕХ верхнеуровневых блоков (в том числе
  // мимо систем уроков), а не только среди других отдельных уроков.
  const moveLesson = (id, dir) => update(d => {
    const l = d.lessons.find(x => x.id === id);
    if (!l) return d;
    const entries = topLevelEntries(d, viewSample).filter(e => {
      if (e.type === "system") {
        const sy = d.lessonSystems.find(x => x.id === e.id);
        return !(sy && sy.archived);
      }
      const le = d.lessons.find(x => x.id === e.id);
      return !(le && le.archived);
    }).sort((a, b) => a.order - b.order);
    const idx = entries.findIndex(e => e.type === "lesson" && e.id === id);
    if (idx === -1) return d;
    if (dir < 0) {
      if (idx === 0) return d;
      reorderTopLevel(d, viewSample, {
        type: "lesson",
        id
      }, entries[idx - 1]);
    } else {
      if (idx >= entries.length - 1) return d;
      reorderTopLevel(d, viewSample, {
        type: "lesson",
        id
      }, entries[idx + 2] || null);
    }
    resyncOrderForGroup(d, l);
    return d;
  });

  // ---- Drag-and-drop: перестановка уроков/систем и перенос уроков между отдельным списком и системами.
  // Отдельные уроки и системы уроков делят общее пространство порядка (topLevelEntries/reorderTopLevel),
  // поэтому урок можно перетащить и поставить перед системой, после неё или между системами. ----
  const reorderLessonInSystem = (systemId, draggedId, targetId) => update(d => {
    if (draggedId === targetId) return d;
    const sys = d.lessonSystems.find(s => s.id === systemId);
    if (!sys) return d;
    const ids = sys.lessonIds.slice();
    const fromIdx = ids.indexOf(draggedId);
    const toIdx = ids.indexOf(targetId);
    if (fromIdx === -1 || toIdx === -1) return d;
    ids.splice(fromIdx, 1);
    ids.splice(toIdx, 0, draggedId);
    sys.lessonIds = ids;
    resyncSystemLessonsByDate(d, systemId);
    return d;
  });
  const moveLessonIntoSystem = (lessonId, systemId, beforeLessonId) => {
    const l0 = data.lessons.find(x => x.id === lessonId);
    const s0 = data.lessonSystems.find(x => x.id === systemId);
    if (l0 && s0 && l0.classLetterId && s0.classLetterId && l0.classLetterId !== s0.classLetterId) {
      showToast("Урок из другой литеры — в систему этой литеры его перенести нельзя (используйте «Перенести в другие литеры»)");
      return;
    }
    moveLessonIntoSystemDo(lessonId, systemId, beforeLessonId);
  };
  const moveLessonIntoSystemDo = (lessonId, systemId, beforeLessonId) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    const sys = d.lessonSystems.find(s => s.id === systemId);
    if (!l || !sys) return d;
    if (l.systemId) {
      const oldSys = d.lessonSystems.find(s => s.id === l.systemId);
      if (oldSys) oldSys.lessonIds = oldSys.lessonIds.filter(id => id !== lessonId);
    }
    l.systemId = systemId;
    const ids = sys.lessonIds.filter(id => id !== lessonId);
    let insertAt = ids.length;
    if (beforeLessonId) {
      const idx = ids.indexOf(beforeLessonId);
      if (idx !== -1) insertAt = idx;
    }
    ids.splice(insertAt, 0, lessonId);
    sys.lessonIds = ids;
    resyncSystemLessonsByDate(d, systemId);
    return d;
  });
  // targetEntry — верхнеуровневый блок {type,id}, ПЕРЕД которым встанет вышедший из системы урок
  // (null — в конец списка). Позиция считается в общем пространстве порядка вместе с системами.
  const moveLessonOutOfSystem = (lessonId, targetEntry) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (!l || !l.systemId) return d;
    const oldSys = d.lessonSystems.find(s => s.id === l.systemId);
    if (oldSys) oldSys.lessonIds = oldSys.lessonIds.filter(id => id !== lessonId);
    l.systemId = null;
    reorderTopLevel(d, viewSample, {
      type: "lesson",
      id: lessonId
    }, targetEntry);
    resyncOrderForGroup(d, l);
    return d;
  });
  // Переставляет один отдельный урок среди ВСЕХ верхнеуровневых блоков группы (уроки и системы).
  const reorderTopLevelLesson = (lessonId, targetEntry) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (!l) return d;
    reorderTopLevel(d, viewSample, {
      type: "lesson",
      id: lessonId
    }, targetEntry);
    resyncOrderForGroup(d, l);
    return d;
  });
  // Переставляет систему уроков среди ВСЕХ верхнеуровневых блоков её класса/курса.
  const reorderTopLevelSystem = (systemId, targetEntry) => update(d => {
    const s = d.lessonSystems.find(x => x.id === systemId);
    if (!s) return d;
    reorderTopLevel(d, viewSample, {
      type: "system",
      id: systemId
    }, targetEntry);
    // система сдвинулась — серии отдельных уроков вокруг неё изменились, упорядочиваем каждую по датам
    const seen = [];
    d.lessons.filter(l => (s.courseId ? l.courseId === s.courseId : l.class === s.class && !l.courseId) && !l.systemId).forEach(l => {
      if (!seen.some(x => sameLessonGroup(x, l))) {
        seen.push(l);
        resyncOrderForGroup(d, l);
      }
    });
    return d;
  });
  const dnd = {
    dragLessonId,
    dragSystemId,
    startLessonDrag: (lessonId, fromSystemId) => {
      setTimeout(() => {
        setDragLessonId(lessonId);
        setDragLessonFromSystemId(fromSystemId);
      }, 0);
    },
    endLessonDrag: () => {
      setDragLessonId(null);
      setDragLessonFromSystemId(null);
    },
    // Урок перетащен и брошен точно на другой урок targetLessonId — вставляется перед ним.
    // targetSystemId — система, которой принадлежит targetLessonId (null — отдельный список).
    dropLessonBefore: (targetLessonId, targetSystemId) => {
      if (!dragLessonId || dragLessonId === targetLessonId) return;
      if (dragLessonFromSystemId === targetSystemId) {
        if (targetSystemId) reorderLessonInSystem(targetSystemId, dragLessonId, targetLessonId);else reorderTopLevelLesson(dragLessonId, {
          type: "lesson",
          id: targetLessonId
        });
      } else if (targetSystemId) {
        moveLessonIntoSystem(dragLessonId, targetSystemId, targetLessonId);
      } else {
        moveLessonOutOfSystem(dragLessonId, {
          type: "lesson",
          id: targetLessonId
        });
      }
    },
    // Урок брошен в пустую область системы/списка (не на конкретный урок) — добавляется в конец.
    dropLessonAtEnd: targetSystemId => {
      if (!dragLessonId || dragLessonFromSystemId === targetSystemId) return;
      if (targetSystemId) moveLessonIntoSystem(dragLessonId, targetSystemId, null);else moveLessonOutOfSystem(dragLessonId, null);
    },
    startSystemDrag: systemId => {
      setTimeout(() => setDragSystemId(systemId), 0);
    },
    endSystemDrag: () => setDragSystemId(null),
    dropSystemOn: targetSystemId => {
      if (dragSystemId) reorderTopLevelSystem(dragSystemId, {
        type: "system",
        id: targetSystemId
      });
      setDragSystemId(null);
    },
    // Зазоры между блоками (перед/после системы, между уроками) — принимают то, что сейчас
    // тащат (урок или систему), и ставят его точно перед targetEntry (null — в самый конец).
    insertDraggedBefore: targetEntry => {
      if (dragLessonId) {
        if (dragLessonFromSystemId) moveLessonOutOfSystem(dragLessonId, targetEntry);else reorderTopLevelLesson(dragLessonId, targetEntry);
        setDragLessonId(null);
        setDragLessonFromSystemId(null);
      } else if (dragSystemId) {
        reorderTopLevelSystem(dragSystemId, targetEntry);
        setDragSystemId(null);
      }
    }
  };
  if (openLessonId) {
    return /*#__PURE__*/React.createElement(LessonPage, {
      data: data,
      update: update,
      lessonId: openLessonId,
      onBack: () => {
        setOpenLessonId(null);
        if (originBack) originBack();
      },
      projectPath: projectPath
    });
  }
  const favC = c => (data.favorites?.classes || []).includes(c);
  const favK = id => (data.favorites?.courses || []).includes(id);
  return /*#__PURE__*/React.createElement("div", {
    className: compact ? "" : "row",
    style: compact ? {
      minWidth: 0
    } : {
      alignItems: "flex-start",
      gap: 20
    }
  }, compact && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      marginBottom: 10
    }
  }, [...CLASSES.filter(favC), ...CLASSES.filter(c => !favC(c))].map(c => /*#__PURE__*/React.createElement(TrackChip, {
    key: c,
    label: (favC(c) ? "★ " : "") + c + " класс",
    color: CLASS_COLORS[c] || "#888",
    active: target.kind === "class" && target.value === c,
    onClick: () => setTarget({
      kind: "class",
      value: c
    })
  })), [...data.courses.filter(c => favK(c.id)), ...data.courses.filter(c => !favK(c.id))].map(c => /*#__PURE__*/React.createElement(TrackChip, {
    key: c.id,
    label: (favK(c.id) ? "★ " : "") + c.name,
    color: c.color || "#888",
    active: target.kind === "course" && target.value === c.id,
    onClick: () => setTarget({
      kind: "course",
      value: c.id
    })
  }))), !compact && /*#__PURE__*/React.createElement("div", {
    style: {
      width: 210,
      flex: "none"
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 6
    }
  }, "\u041A\u043B\u0430\u0441\u0441\u044B"), [...CLASSES.filter(c => (data.favorites?.classes || []).includes(c)), ...CLASSES.filter(c => !(data.favorites?.classes || []).includes(c))].map(c => /*#__PURE__*/React.createElement("div", {
    key: c,
    className: "tk-row",
    style: {
      cursor: "pointer",
      background: target.kind === "class" && target.value === c ? "var(--accent-soft)" : "transparent",
      borderRadius: 8
    },
    onClick: () => setTarget({
      kind: "class",
      value: c
    })
  }, /*#__PURE__*/React.createElement("span", {
    className: "tk-name"
  }, c, " \u043A\u043B\u0430\u0441\u0441"), /*#__PURE__*/React.createElement("span", {
    title: (data.favorites?.classes || []).includes(c) ? "Убрать из избранного" : "В избранное: класс будет первым и откроется при запуске",
    style: {
      cursor: "pointer",
      fontSize: 15,
      color: (data.favorites?.classes || []).includes(c) ? "#C98A2B" : "#c3ceda"
    },
    onClick: e => {
      e.stopPropagation();
      update(d => {
        const f = d.favorites.classes;
        d.favorites.classes = f.includes(c) ? f.filter(x => x !== c) : [...f, c];
        return d;
      });
    }
  }, (data.favorites?.classes || []).includes(c) ? "★" : "☆"))), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      margin: "10px 0 6px"
    }
  }, "\u041A\u0443\u0440\u0441\u044B"), [...data.courses.filter(c => (data.favorites?.courses || []).includes(c.id)), ...data.courses.filter(c => !(data.favorites?.courses || []).includes(c.id))].map(c => /*#__PURE__*/React.createElement("div", {
    key: c.id,
    className: "tk-row",
    style: {
      cursor: "pointer",
      background: target.kind === "course" && target.value === c.id ? "var(--accent-soft)" : "transparent",
      borderRadius: 8
    },
    onClick: () => setTarget({
      kind: "course",
      value: c.id
    })
  }, /*#__PURE__*/React.createElement("span", {
    className: "tk-name"
  }, c.name), /*#__PURE__*/React.createElement("span", {
    title: (data.favorites?.courses || []).includes(c.id) ? "Убрать из избранного" : "В избранное",
    style: {
      cursor: "pointer",
      fontSize: 15,
      color: (data.favorites?.courses || []).includes(c.id) ? "#C98A2B" : "#c3ceda"
    },
    onClick: e => {
      e.stopPropagation();
      update(d => {
        const f = d.favorites.courses;
        d.favorites.courses = f.includes(c.id) ? f.filter(x => x !== c.id) : [...f, c.id];
        return d;
      });
    }
  }, (data.favorites?.courses || []).includes(c.id) ? "★" : "☆"))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: addCourse
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u041A\u0443\u0440\u0441")))), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      minWidth: 0
    }
  }, letters.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10,
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041B\u0438\u0442\u0435\u0440\u0430 (\u0435\u0441\u043B\u0438 \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u043E ", groupValue, " \u043A\u043B\u0430\u0441\u0441\u043E\u0432 \u0432 \u043F\u0430\u0440\u0430\u043B\u043B\u0435\u043B\u0438):"), letters.map(cl => /*#__PURE__*/React.createElement(TrackChip, {
    key: cl.id,
    label: `${groupValue}${cl.name}`,
    color: "#8593a3",
    active: letterFilter === cl.id,
    onClick: () => setLetterFilter(cl.id)
  })), (letterFilter === "none" || hasNoLetterItems(data, groupValue)) && /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0411\u0435\u0437 \u043B\u0438\u0442\u0435\u0440\u044B",
    color: "#8593a3",
    active: letterFilter === "none",
    onClick: () => setLetterFilter("none")
  }), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0412\u0441\u0435",
    color: "#8593a3",
    active: letterFilter === "all",
    onClick: () => setLetterFilter("all")
  })), /*#__PURE__*/React.createElement("div", {
    className: "toolbar"
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), onToggleJournal && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onToggleJournal,
    title: "\u0416\u0443\u0440\u043D\u0430\u043B \u044D\u0442\u043E\u0433\u043E \u043A\u043B\u0430\u0441\u0441\u0430 \u0440\u044F\u0434\u043E\u043C \u0441\u043E \u0441\u043F\u0438\u0441\u043A\u043E\u043C \u2014 \u043C\u043E\u0436\u043D\u043E \u0441\u0442\u0430\u0432\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0438, \u043D\u0435 \u0443\u0445\u043E\u0434\u044F \u0441 \u0443\u0440\u043E\u043A\u0430"
  }, journalOpen ? "✕ Закрыть журнал" : "📒 Открыть журнал"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => setSendOpen(true),
    title: "\u0412\u044B\u0431\u0440\u0430\u0442\u044C \u0443\u0440\u043E\u043A\u0438 \u0438 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B, \u043F\u0440\u0438 \u0436\u0435\u043B\u0430\u043D\u0438\u0438 \u043F\u043E\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u0438 \u043E\u0442\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u0443\u0447\u0430\u0449\u0438\u043C\u0441\u044F \u0444\u0430\u0439\u043B\u043E\u043C Word"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u041E\u0442\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u0437\u0430\u0434\u0430\u043D\u0438\u0435"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => importRef.current && importRef.current.click(),
    title: "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0432 \u0441\u043F\u0438\u0441\u043E\u043A \u0443\u0440\u043E\u043A \u0438\u0437 \u0444\u0430\u0439\u043B\u0430 .lsnlit"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 13
  }), " \u0418\u043C\u043F\u043E\u0440\u0442 \u0443\u0440\u043E\u043A\u0430 (.lsnlit)"), /*#__PURE__*/React.createElement("input", {
    ref: importRef,
    type: "file",
    accept: ".lsnlit,.json",
    multiple: true,
    style: {
      display: "none"
    },
    onChange: e => {
      if (e.target.files && e.target.files.length) handleImportLessons(e.target.files);
      e.target.value = "";
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => importSysRef.current && importSysRef.current.click(),
    title: "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0432 \u0441\u043F\u0438\u0441\u043E\u043A \u0441\u0438\u0441\u0442\u0435\u043C\u0443 \u0443\u0440\u043E\u043A\u043E\u0432 \u0438\u0437 \u0444\u0430\u0439\u043B\u0430 .lsnsys"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 13
  }), " \u0418\u043C\u043F\u043E\u0440\u0442 \u0441\u0438\u0441\u0442\u0435\u043C\u044B (.lsnsys)"), /*#__PURE__*/React.createElement("input", {
    ref: importSysRef,
    type: "file",
    accept: ".lsnsys,.json",
    multiple: true,
    style: {
      display: "none"
    },
    onChange: e => {
      if (e.target.files && e.target.files.length) handleImportSystems(e.target.files);
      e.target.value = "";
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    onClick: addSystem
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u0421\u0438\u0441\u0442\u0435\u043C\u0443 \u0443\u0440\u043E\u043A\u043E\u0432"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: doAddLesson
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u0423\u0440\u043E\u043A")), topLevelCombined.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041F\u043E\u043A\u0430 \u043D\u0435\u0442 \u0443\u0440\u043E\u043A\u043E\u0432. \u041D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB\u0423\u0440\u043E\u043A\xBB \u0438\u043B\u0438 \xAB\u0421\u0438\u0441\u0442\u0435\u043C\u0443 \u0443\u0440\u043E\u043A\u043E\u0432\xBB."), /*#__PURE__*/React.createElement(TopLevelGap, {
    dnd: dnd,
    target: topLevelCombined[0] ? {
      type: topLevelCombined[0].type,
      id: topLevelCombined[0].item.id
    } : null
  }), topLevelCombined.map((entry, i) => /*#__PURE__*/React.createElement(React.Fragment, {
    key: entry.type + ":" + entry.item.id
  }, entry.type === "system" ? /*#__PURE__*/React.createElement(LessonSystemBlock, {
    data: data,
    update: update,
    system: entry.item,
    onOpenLesson: setOpenLessonId,
    projectPath: projectPath,
    dnd: dnd
  }) : /*#__PURE__*/React.createElement(LessonCard, {
    data: data,
    lesson: entry.item,
    index: standaloneIndexById.get(entry.item.id),
    onOpen: () => setOpenLessonId(entry.item.id),
    onMove: dir => moveLesson(entry.item.id, dir),
    onToggleDone: () => toggleLessonCompleted(entry.item.id),
    drag: {
      onDragStart: () => dnd.startLessonDrag(entry.item.id, null),
      onDragEnd: dnd.endLessonDrag,
      onDrop: () => dnd.dropLessonBefore(entry.item.id, null)
    }
  }), /*#__PURE__*/React.createElement(TopLevelGap, {
    dnd: dnd,
    target: topLevelCombined[i + 1] ? {
      type: topLevelCombined[i + 1].type,
      id: topLevelCombined[i + 1].item.id
    } : null
  }))), /*#__PURE__*/React.createElement(ReservePool, {
    data: data,
    update: update,
    groupKind: groupKind,
    groupValue: groupValue
  }), archivedHere.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "panel",
    style: {
      marginTop: 14
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-head",
    style: {
      cursor: "pointer"
    },
    onClick: () => setArchOpen(v => !v)
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      flex: 1
    }
  }, "\u0410\u0440\u0445\u0438\u0432 ", /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "(", archivedHere.length, ")")), /*#__PURE__*/React.createElement(Icon, {
    name: "chev",
    size: 14
  })), archOpen && /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 8
    }
  }, "\u0421\u044E\u0434\u0430 \u043F\u043E\u043F\u0430\u0434\u0430\u044E\u0442 \u043F\u0440\u043E\u0439\u0434\u0435\u043D\u043D\u044B\u0435 \u0443\u0440\u043E\u043A\u0438 (\u0433\u0430\u043B\u043E\u0447\u043A\u0430 \u0441\u043B\u0435\u0432\u0430 \u0432 \u0441\u043F\u0438\u0441\u043A\u0435) \u0438 \u0443\u0440\u043E\u043A\u0438, \u043E\u0442\u043F\u0440\u0430\u0432\u043B\u0435\u043D\u043D\u044B\u0435 \u0432 \u0430\u0440\u0445\u0438\u0432 \u0432\u0440\u0443\u0447\u043D\u0443\u044E; \u043A\u043E\u0433\u0434\u0430 \u043F\u0440\u043E\u0439\u0434\u0435\u043D\u044B \u0432\u0441\u0435 \u0443\u0440\u043E\u043A\u0438 \u0441\u0438\u0441\u0442\u0435\u043C\u044B, \u0432 \u0430\u0440\u0445\u0438\u0432 \u0443\u0445\u043E\u0434\u0438\u0442 \u0438 \u0441\u0430\u043C\u0430 \u0441\u0438\u0441\u0442\u0435\u043C\u0430. \u0412 \u0441\u043F\u0438\u0441\u043A\u0435 \u0438\u0445 \u043D\u0435\u0442, \u043D\u043E \u0432 \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u0435 \u043E\u043D\u0438 \u043E\u0441\u0442\u0430\u044E\u0442\u0441\u044F. \xAB\u0420\u0430\u0437\u0430\u0440\u0445\u0438\u0432\u0438\u0440\u043E\u0432\u0430\u0442\u044C\xBB \u0432\u0435\u0440\u043D\u0451\u0442 \u0443\u0440\u043E\u043A \u0432 \u0441\u043F\u0438\u0441\u043E\u043A (\u043D\u0435\u043F\u0440\u043E\u0439\u0434\u0435\u043D\u043D\u044B\u043C) \u0438 \u043F\u043E\u0441\u0442\u0430\u0432\u0438\u0442 \u043D\u0430 \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u0443\u044E \u0434\u0430\u0442\u0443."), archivedHere.map(l => {
    const sys = l.systemId ? data.lessonSystems.find(x => x.id === l.systemId) : null;
    return /*#__PURE__*/React.createElement("div", {
      key: l.id,
      className: "row wrap",
      style: {
        padding: "6px 0",
        borderBottom: "1px solid var(--line)",
        gap: 8
      }
    }, l.plannedDate && /*#__PURE__*/React.createElement("span", {
      className: "chip",
      style: {
        flex: "none",
        padding: "2px 8px",
        fontSize: 11.5
      }
    }, weekdayShort(l.plannedDate), " ", formatDateRu(l.plannedDate).slice(0, 5)), /*#__PURE__*/React.createElement("div", {
      style: {
        flex: 1,
        minWidth: 160,
        fontSize: 13.5,
        fontWeight: 600
      }
    }, l.topic || "Без темы", l.completed ? /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, " \xB7 \u043F\u0440\u043E\u0439\u0434\u0435\u043D", l.actualDate ? " " + formatDateRu(l.actualDate) : "") : null, sys ? /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, " \xB7 \u0441\u0438\u0441\u0442\u0435\u043C\u0430 \xAB", sys.title || "без названия", "\xBB", sys.archived ? " (вся в архиве)" : "") : null), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => setOpenLessonId(l.id)
    }, "\u041E\u0442\u043A\u0440\u044B\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
      kind: "primary",
      sm: true,
      onClick: () => setUnarchiveId(l.id)
    }, "\u21A9 \u0420\u0430\u0437\u0430\u0440\u0445\u0438\u0432\u0438\u0440\u043E\u0432\u0430\u0442\u044C"));
  }))), unarchiveId && /*#__PURE__*/React.createElement(UnarchiveModal, {
    data: data,
    update: update,
    lessonId: unarchiveId,
    onClose: () => setUnarchiveId(null)
  }), sendOpen && /*#__PURE__*/React.createElement(SendLessonsModal, {
    data: data,
    onClose: () => setSendOpen(false),
    lessons: lessonsHere.filter(l => !l.archived).sort((a, b) => (a.plannedDate || "9999") < (b.plannedDate || "9999") ? -1 : (a.plannedDate || "9999") > (b.plannedDate || "9999") ? 1 : a.order - b.order),
    groupLabel: groupKind === "class" ? groupValue + (wantLetterId ? classLetterName(data, wantLetterId) : "") + " класс" : (data.courses.find(c => c.id === groupValue) || {}).name || "Курс"
  })));
}
;
// ============================================================
// calendar.js — Календарь (недели/месяц, каникулы) + Журнал (ученики, отметки).
// ============================================================

const MONTHS_RU = ["январь", "февраль", "март", "апрель", "май", "июнь", "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь"];
const MONTHS_RU_SHORT = ["янв", "фев", "мар", "апр", "мая", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"];
const WEEKDAYS_RU = ["Пн", "Вт", "Ср", "Чт", "Пт", "Сб", "Вс"];
const CLASS_COLORS = {
  "5": "#2678b3",
  "6": "#2f9e64",
  "7": "#e0ac00",
  "8": "#c0392b",
  "9": "#7e57c2",
  "10": "#00838f",
  "11": "#ad6800"
};
function pad2(n) {
  return String(n).padStart(2, "0");
}
function isoOfDate(d) {
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1) + "-" + pad2(d.getDate());
}
function fromIso(iso) {
  const p = iso.split("-");
  return new Date(+p[0], +p[1] - 1, +p[2]);
}
function addDaysIso(iso, n) {
  const d = fromIso(iso);
  d.setDate(d.getDate() + n);
  return isoOfDate(d);
}
function mondayOf(iso) {
  const d = fromIso(iso);
  const wd = (d.getDay() + 6) % 7;
  d.setDate(d.getDate() - wd);
  return isoOfDate(d);
}
function todayIso() {
  return isoOfDate(new Date());
}
function curMonthYm() {
  const d = new Date();
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1);
}
function monthDaysCount(ym) {
  const p = ym.split("-");
  return new Date(+p[0], +p[1], 0).getDate();
}
function addMonthsYm(ym, delta) {
  const p = ym.split("-");
  const d = new Date(+p[0], +p[1] - 1 + delta, 1);
  return d.getFullYear() + "-" + pad2(d.getMonth() + 1);
}
function ymLabel(ym) {
  const p = ym.split("-");
  return MONTHS_RU[+p[1] - 1] + " " + p[0];
}
function dayNum(iso) {
  return +iso.slice(8, 10);
}
function shortMon(iso) {
  return MONTHS_RU_SHORT[+iso.slice(5, 7) - 1];
}
function formatDateRu(iso) {
  if (!iso) return "";
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
}
function ruShortDate(iso) {
  return dayNum(iso) + " " + shortMon(iso);
}
function weeksOfMonthYm(ym) {
  const start = mondayOf(ym + "-01");
  const lastIso = ym + "-" + pad2(monthDaysCount(ym));
  const out = [];
  let cur = start,
    g = 0;
  while (cur <= lastIso && g < 10) {
    out.push(cur);
    cur = addDaysIso(cur, 7);
    g++;
  }
  return out;
}
function monthGridYm(ym) {
  const start = mondayOf(ym + "-01");
  const cells = [];
  for (let i = 0; i < 42; i++) {
    const iso = addDaysIso(start, i);
    cells.push({
      iso,
      inMonth: iso.slice(0, 7) === ym
    });
  }
  return cells;
}
function weekLabelRange(mon) {
  const sun = addDaysIso(mon, 6);
  if (mon.slice(5, 7) === sun.slice(5, 7)) return dayNum(mon) + "–" + dayNum(sun) + " " + shortMon(sun);
  return dayNum(mon) + " " + shortMon(mon) + " – " + dayNum(sun) + " " + shortMon(sun);
}
function sortedVacations(data) {
  return [...data.calendar.vacations].sort((a, b) => a.start < b.start ? -1 : a.start > b.start ? 1 : 0);
}
function vacationAt(data, iso) {
  return data.calendar.vacations.find(v => v.start && v.end && iso >= v.start && iso <= v.end) || null;
}
function quarterOfDate(data, iso) {
  let n = 0;
  sortedVacations(data).forEach(v => {
    if (v.end && v.end < iso) n++;
  });
  return n;
}
function quarterCountOf(data) {
  return data.calendar.vacations.length + 1;
}
function allTracks(data) {
  const fav = data.favorites || {
    classes: [],
    courses: []
  };
  const list = [...CLASSES.map(c => ({
    kind: "class",
    value: c,
    label: `${c} класс`,
    color: CLASS_COLORS[c] || "#888",
    fav: fav.classes.includes(c)
  })), ...data.courses.map(c => ({
    kind: "course",
    value: c.id,
    label: c.name,
    color: c.color || "#888",
    fav: fav.courses.includes(c.id)
  }))];
  return [...list.filter(t => t.fav), ...list.filter(t => !t.fav)]; // избранные — первыми
}
function trackKeyOf(kindOrLesson, value) {
  if (typeof kindOrLesson === "object") {
    const l = kindOrLesson;
    return l.courseId ? `course:${l.courseId}` : `class:${l.class}`;
  }
  return `${kindOrLesson}:${value}`;
}
function trackColorOf(data, kind, value) {
  return kind === "course" ? data.courses.find(c => c.id === value)?.color || "#888" : CLASS_COLORS[value] || "#888";
}
function lessonTrackColor(lesson, data) {
  return lesson.courseId ? trackColorOf(data, "course", lesson.courseId) : trackColorOf(data, "class", lesson.class);
}
function markColorFor(v, scale) {
  if (v === "Н") return "#8593a3";
  if (scale === "100") return v >= 85 ? "#1b7a4a" : v >= 65 ? "#2678b3" : v >= 50 ? "#c98a2b" : "#c0392b";
  return v >= 5 ? "#1b7a4a" : v >= 4 ? "#2678b3" : v >= 3 ? "#c98a2b" : "#c0392b";
}
function weekdayShort(iso) {
  const d = fromIso(iso);
  const wd = (d.getDay() + 6) % 7;
  return WEEKDAYS_RU[wd];
}
function semesterOfQuarter(q) {
  return Math.floor(q / 2);
}
function quartersInPeriod(periodMode, periodIdx, qc) {
  if (periodMode === "semester") {
    const start = periodIdx * 2;
    return [start, start + 1].filter(q => q < qc);
  }
  return [periodIdx];
}
function isLessonStudied(l) {
  return !!l.completed || !!l.actualDate;
}
function markTooltip(data, m) {
  const parts = [formatDateRu(m.date)];
  if (m.lessonId) {
    const l = data.lessons.find(x => x.id === m.lessonId);
    if (l) parts.push(l.topic || "(без темы)");
  }
  if (m.exerciseId) {
    const ex = data.examples.find(e => e.id === m.exerciseId);
    if (ex) parts.push(stripHtml(ex.text).slice(0, 60));
  }
  if (m.questionId) {
    for (const key of ["theoryHistory", "personalities", "works"]) {
      const ent = data[key].find(e => (e.course?.questions || []).some(q => q.id === m.questionId));
      if (ent) {
        const q = ent.course.questions.find(x => x.id === m.questionId);
        if (q) parts.push(q.text.slice(0, 60));
        break;
      }
    }
  }
  if (m.reason) parts.push(m.reason);
  return parts.join(" · ");
}
function avgColorFor(a) {
  return a >= 4.5 ? "#1b7a4a" : a >= 3.5 ? "#2678b3" : a >= 2.5 ? "#c98a2b" : "#c0392b";
}
function fmtAvgNum(a) {
  return a == null ? "—" : Number.isInteger(a) ? String(a) : a.toFixed(2);
}
function studentMarksOf(data, sid, scale) {
  return data.calendar.marks.filter(m => m.studentId === sid && (scale == null || (m.scale || "5") === scale));
}
// Только реальные числовые отметки — без "Н" (пропуск не участвует в среднем и подсчёте отметок).
function numericMarksOf(data, sid, scale) {
  return studentMarksOf(data, sid, scale).filter(m => m.value !== "Н");
}
function quarterAvgOf(data, sid, q, scale) {
  const ms = numericMarksOf(data, sid, scale || "5").filter(m => quarterOfDate(data, m.date) === q);
  if (!ms.length) return null;
  return Math.round(ms.reduce((a, m) => a + (+m.value || 0), 0) / ms.length * 100) / 100;
}
function yearAvgOf(data, sid, scale) {
  const ms = numericMarksOf(data, sid, scale || "5");
  if (!ms.length) return null;
  return Math.round(ms.reduce((a, m) => a + (+m.value || 0), 0) / ms.length * 100) / 100;
}
function roundToGrade(avg) {
  return avg == null ? null : Math.max(1, Math.min(5, Math.round(avg)));
}
function performancePct(grades) {
  const withGrade = grades.filter(g => g != null);
  if (!grades.length) return null;
  return Math.round(withGrade.filter(g => g >= 3).length / grades.length * 1000) / 10;
}
function qualityPct(grades) {
  const withGrade = grades.filter(g => g != null);
  if (!grades.length) return null;
  return Math.round(withGrade.filter(g => g >= 4).length / grades.length * 1000) / 10;
}
function urgencyColorFor(haveCount, requiredMarks, remainingLessons) {
  if (!requiredMarks) return null;
  const shortfall = requiredMarks - haveCount;
  if (shortfall <= 0) return null;
  if (remainingLessons <= 0) return "#c0392b";
  const urgency = shortfall / remainingLessons;
  if (urgency >= 1) return "#c0392b";
  if (urgency >= 0.5) return "#d97b3f";
  return "#dba556";
}

// ---- Литеры классов и подгруппы курсов ----
const GRADING_SYSTEM_LABELS = {
  "5": "5-балльная",
  "100": "100-балльная",
  combined: "Комбинированная"
};
function scalesInUse(gradingSystem) {
  return gradingSystem === "100" ? ["100"] : gradingSystem === "combined" ? ["5", "100"] : ["5"];
}
// «Без литеры» в фильтрах показываем только если в параллели есть урок или ученик без литеры.
function hasNoLetterItems(data, grade) {
  return data.lessons.some(l => !l.archived && !l.courseId && l.class === grade && !l.classLetterId) || data.calendar.students.some(st => !st.courseId && st.class === grade && !st.classLetterId);
}
function lettersForGrade(data, grade) {
  return (data.classLetters || []).filter(cl => cl.grade === grade).sort((a, b) => String(a.name).localeCompare(String(b.name), "ru"));
}
function classLetterName(data, id) {
  const cl = (data.classLetters || []).find(x => x.id === id);
  return cl ? cl.name : "";
}
// Подходит ли выбранная литера/подгруппа к классу или курсу (при смене класса неподходящая сбрасывается в «Все»).
function letterValidFor(data, kind, value, f) {
  if (f === "all" || f === "none") return true;
  return kind === "course" ? subgroupsForCourse(data, value).some(g => g.id === f) : lettersForGrade(data, value).some(l => l.id === f);
}
function subgroupsForCourse(data, courseId) {
  return (data.courseSubgroups || []).filter(g => g.courseId === courseId);
}
function classSettingsKey(kind, value, classLetterId) {
  return `${kind}:${value}:${classLetterId || "none"}`;
}
function getClassSettings(data, kind, value, classLetterId) {
  const key = classSettingsKey(kind, value, classLetterId);
  const found = (data.classSettings || []).find(s => classSettingsKey(s.kind, s.value, s.classLetterId) === key);
  return found || {
    gradingSystem: "5",
    periodMode: "quarter",
    requiredMarks: null
  };
}
function settingsForLesson(data, lesson) {
  return getClassSettings(data, lesson.courseId ? "course" : "class", lesson.courseId || lesson.class, lesson.classLetterId || null);
}
function settingsForStudent(data, student) {
  return getClassSettings(data, student.courseId ? "course" : "class", student.courseId || student.class, student.courseId ? student.courseSubgroupId || null : student.classLetterId || null);
}
function nextGrade(grade) {
  const idx = CLASSES.indexOf(grade);
  return idx >= 0 && idx < CLASSES.length - 1 ? CLASSES[idx + 1] : null;
}
// Литеры того же класса, в материалах которых этого упражнения нет ни в одном уроке —
// подсказка для «Запасных упражнений»: видно, где ещё не пройдено то же самое.
function lettersMissingExercise(data, grade, exerciseId, excludeLetterId) {
  return lettersForGrade(data, grade).filter(cl => cl.id !== excludeLetterId).filter(cl => !data.lessons.some(l => l.class === grade && !l.courseId && l.classLetterId === cl.id && (l.items || []).some(it => it.type === "exercise" && it.refId === exerciseId))).map(cl => `${grade}${cl.name}`);
}

// ---- Рассадка по партам ----
const SEATING_FILE_LABELS = ["У окна", "Посередине", "У двери"];
function shuffleArray(arr) {
  const a = [...arr];
  for (let i = a.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const t = a[i];
    a[i] = a[j];
    a[j] = t;
  }
  return a;
}
function seatingLayoutKey(kind, value, classLetterId) {
  return `${kind}:${value}:${classLetterId || "none"}`;
}
function getSeatingLayout(data, kind, value, classLetterId) {
  const key = seatingLayoutKey(kind, value, classLetterId);
  return (data.seatingLayouts || []).find(l => seatingLayoutKey(l.kind, l.value, l.classLetterId) === key) || null;
}
// В каждом ряду всегда есть парта со свободным местом; лишние пустые парты в хвосте убираются.
function ensureSeatingCapacity(layout) {
  layout.files.forEach(file => {
    while (file.length > 1 && file[file.length - 1].seats.every(s => !s) && file[file.length - 2].seats.every(s => !s)) file.pop();
    if (file.length === 0 || !file.some(desk => desk.seats.some(s => !s))) file.push({
      id: uid(),
      seats: [null, null]
    });
  });
}
function getOrCreateSeatingLayout(d, kind, value, classLetterId) {
  d.seatingLayouts = d.seatingLayouts || [];
  let layout = getSeatingLayout(d, kind, value, classLetterId);
  if (!layout) {
    layout = {
      id: uid(),
      kind,
      value,
      classLetterId: classLetterId || null,
      files: [[], [], []],
      visionLimits: [],
      seatFurther: [],
      avoidGroups: []
    };
    d.seatingLayouts.push(layout);
  }
  ensureSeatingCapacity(layout);
  return layout;
}
// Ищет свободное место — сперва в предпочтительном ряду, потом в любом другом.
// Ученик с ограничением по зрению садится только не дальше указанной парты. Возвращает true/false «посадили / нет места».
function placeInFirstEmptySeat(layout, studentId, preferFileIdx) {
  const lim = (layout.visionLimits || []).find(v => v.studentId === studentId);
  const maxDepth = lim ? lim.maxRow : Infinity;
  const order = [preferFileIdx, ...[0, 1, 2].filter(i => i !== preferFileIdx)];
  for (const fi of order) {
    for (let di = 0; di < layout.files[fi].length && di + 1 <= maxDepth; di++) {
      const desk = layout.files[fi][di];
      const si = desk.seats.findIndex(x => !x);
      if (si !== -1) {
        desk.seats[si] = studentId;
        return true;
      }
    }
  }
  if (layout.files[preferFileIdx].length + 1 <= maxDepth) {
    layout.files[preferFileIdx].push({
      id: uid(),
      seats: [studentId, null]
    });
    return true;
  }
  return false;
}
// Сидят ли двое «рядом или близко»: на одной парте, на соседних, впереди/сзади, в соседнем ряду, по диагонали.
function seatedNear(files, a, b) {
  let pa = null,
    pb = null;
  files.forEach((file, fi) => file.forEach((desk, di) => desk.seats.forEach(sid => {
    if (sid === a) pa = {
      fi,
      di
    };
    if (sid === b) pb = {
      fi,
      di
    };
  })));
  if (!pa || !pb) return false;
  return Math.abs(pa.fi - pb.fi) <= 1 && Math.abs(pa.di - pb.di) <= 1;
}
function countNearInGroups(files, groups) {
  let n = 0;
  groups.forEach(g => {
    for (let i = 0; i < g.members.length; i++) for (let j = i + 1; j < g.members.length; j++) if (seatedNear(files, g.members[i], g.members[j])) n++;
  });
  return n;
}
// Подсказка с именем ученика: свой выпадающий список (колесо и палец прокручивают), ввод фильтрует список.
function StudentNameInput({
  value,
  roster,
  flags,
  onChange,
  onBlur,
  autoFocus,
  style,
  placeholder
}) {
  const [open, setOpen] = useState(false);
  const q = (value || "").trim().toLowerCase();
  const exact = roster.some(r => r.name.trim().toLowerCase() === q);
  const list = roster.filter(r => !q || exact || r.name.toLowerCase().includes(q));
  return /*#__PURE__*/React.createElement("div", {
    style: {
      position: "relative",
      display: "inline-block"
    }
  }, /*#__PURE__*/React.createElement("input", {
    autoFocus: autoFocus,
    style: style,
    placeholder: placeholder,
    value: value || "",
    onFocus: () => setOpen(true),
    onChange: e => {
      onChange(e.target.value);
      setOpen(true);
    },
    onBlur: () => {
      setTimeout(() => setOpen(false), 120);
      onBlur && onBlur();
    }
  }), open && list.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      position: "absolute",
      top: "100%",
      left: 0,
      minWidth: "100%",
      zIndex: 60,
      background: "#fff",
      border: "1.5px solid #cfe3f2",
      borderRadius: 8,
      boxShadow: "0 6px 18px rgba(19,44,83,.18)",
      maxHeight: 210,
      overflowY: "auto",
      touchAction: "pan-y"
    }
  }, list.map(r => /*#__PURE__*/React.createElement("div", {
    key: r.id,
    onMouseDown: e => {
      e.preventDefault();
      onChange(r.name);
      setOpen(false);
    },
    style: {
      padding: "6px 10px",
      fontSize: 12.5,
      cursor: "pointer",
      whiteSpace: "nowrap",
      color: r.noCall ? "#8a6fb0" : "#12233a",
      ...(flags && callFlagStyle(flags.get(r.id)) ? {
        background: callFlagStyle(flags.get(r.id)).background
      } : null)
    }
  }, r.name, r.noCall ? " · не вызывать" : ""))));
}
function buildSeatingLayoutHtml(layout, nameById, small) {
  const cols = layout.files.map((file, fi) => {
    const deskRows = file.map((desk, di) => {
      const names = desk.seats.map(s => s ? escapeHtml(nameById(s)) : "&nbsp;");
      return `<div style="border:1px solid #999;border-radius:4px;padding:${small ? "2px 4px" : "6px 10px"};margin-bottom:${small ? 2 : 6}px;font-size:${small ? 8 : 13}px;min-height:${small ? 14 : 34}px;"><span style="color:#aaa;font-size:${small ? 6 : 10}px;margin-right:4px;">${di + 1}</span>${names[0]}<br/>${names[1]}</div>`;
    }).join("");
    return `<td style="vertical-align:top;padding:0 ${small ? 3 : 8}px;"><div style="text-align:center;font-weight:bold;font-size:${small ? 8 : 12}px;margin-bottom:${small ? 2 : 6}px;">${escapeHtml(SEATING_FILE_LABELS[fi])}</div>${deskRows}</td>`;
  }).join("");
  return `<div style="border:2px solid #333;padding:${small ? 6 : 16}px;${small ? "width:250px;display:inline-block;vertical-align:top;margin:6px;" : ""}"><div style="text-align:center;font-weight:bold;background:#333;color:#fff;padding:${small ? 4 : 10}px;margin-bottom:${small ? 6 : 14}px;font-size:${small ? 9 : 15}px;">ДОСКА</div><table style="width:100%"><tr>${cols}</tr></table></div>`;
}

// Кого предложить вызвать к доске: сначала те, у кого меньше всего отметок в текущей четверти, затем тот,
// у кого отметки самые низкие (средний < 3 по 5-балльной шкале, < 50 по 100-балльной). «Не вызывать» не предлагаются.
// Подсветка предлагаемых к доске: КРАСНЫЙ — мало оценок в этой четверти (меньше половины от среднего по классу,
// а если в «Списках классов» задана норма отметок — меньше половины нормы), ЖЁЛТЫЙ — не успевает (средний балл ниже 3 по
// 5-балльной / ниже 50 по 100-балльной шкале). Если верно и то и другое — красно-жёлтый. «Не вызывать» не учитываются.
function studentCallFlags(data, roster, scale, requiredMarks) {
  const q = quarterOfDate(data, todayIso());
  const rows = roster.filter(st => !st.noCall).map(st => ({
    id: st.id,
    n: numericMarksOf(data, st.id, scale).filter(m => quarterOfDate(data, m.date) === q).length,
    avg: quarterAvgOf(data, st.id, q, scale)
  }));
  const meanN = rows.length ? rows.reduce((sum, r) => sum + r.n, 0) / rows.length : 0;
  const limit = requiredMarks ? requiredMarks / 2 : meanN / 2;
  const enough = !!requiredMarks || meanN >= 1.5; // в самом начале четверти, когда отметок почти ни у кого нет, красным не красим
  const poorLimit = scale === "100" ? 50 : 3;
  const out = new Map();
  rows.forEach(r => out.set(r.id, {
    n: r.n,
    avg: r.avg,
    few: enough && r.n < Math.max(1, limit),
    poor: r.avg != null && r.avg < poorLimit
  }));
  return out;
}
function callFlagStyle(f) {
  if (!f || !f.few && !f.poor) return null;
  if (f.few && f.poor) return {
    background: "linear-gradient(90deg, #f9cfc9 50%, #fff0a8 50%)",
    border: "1.5px solid #e0a88a",
    borderRadius: 8
  };
  if (f.few) return {
    background: "#f9cfc9",
    border: "1.5px solid #e09a90",
    borderRadius: 8
  };
  return {
    background: "#fff0a8",
    border: "1.5px solid #e0c24a",
    borderRadius: 8
  };
}
function callFlagTitle(f) {
  if (!f) return "";
  return (f.few ? " · МАЛО ОЦЕНОК" : "") + (f.poor ? " · НЕ УСПЕВАЕТ" : "");
}
function CallFlagLegend() {
  return /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      fontSize: 11.5
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      background: "#f9cfc9",
      borderRadius: 4,
      padding: "0 6px"
    }
  }, "\u043A\u0440\u0430\u0441\u043D\u044B\u0439"), " \u2014 \u043C\u0430\u043B\u043E \u043E\u0446\u0435\u043D\u043E\u043A \xB7 ", /*#__PURE__*/React.createElement("span", {
    style: {
      background: "#fff0a8",
      borderRadius: 4,
      padding: "0 6px"
    }
  }, "\u0436\u0451\u043B\u0442\u044B\u0439"), " \u2014 \u043D\u0435 \u0443\u0441\u043F\u0435\u0432\u0430\u0435\u0442");
}
function recommendToCall(data, roster, scale, requiredMarks) {
  const flags = studentCallFlags(data, roster, scale, requiredMarks);
  const q = quarterOfDate(data, todayIso());
  const rows = roster.filter(st => !st.noCall).map(st => ({
    st,
    n: numericMarksOf(data, st.id, scale).filter(m => quarterOfDate(data, m.date) === q).length,
    avg: quarterAvgOf(data, st.id, q, scale)
  }));
  const byFew = [...rows].sort((a, b) => a.n - b.n || (a.avg == null ? 99 : a.avg) - (b.avg == null ? 99 : b.avg) || a.st.name.localeCompare(b.st.name, "ru"));
  const poorLimit = scale === "100" ? 50 : 3;
  const poor = rows.filter(r => r.avg != null && r.avg < poorLimit).sort((a, b) => a.avg - b.avg);
  const chosen = byFew.slice(0, 2);
  const third = poor.find(r => !chosen.includes(r)) || byFew.find(r => !chosen.includes(r));
  if (third) chosen.push(third);
  return chosen.map(r => ({
    ...r.st,
    _n: r.n,
    _avg: r.avg,
    _few: !!(flags.get(r.st.id) || {}).few,
    _poor: !!(flags.get(r.st.id) || {}).poor
  }));
}
// Подпись «где» у урока: «6Б класс» / название курса.
function lessonPlaceLabel(l, data) {
  if (l.courseId) return (data.courses.find(c => c.id === l.courseId) || {}).name || "Курс";
  return (l.class || "") + (l.classLetterId ? classLetterName(data, l.classLetterId) : "") + " класс";
}

/* ============================================================
   «Вечерний набор»: очередь «класс → дата» с автонабором оценок в школьный электронный журнал («Моя школа»).
   Программа нажимает клавиши в окне браузера, которое в фокусе: оценка → стрелка вниз → оценка → … (Windows).
   На каждом шаге вы сами открываете нужный журнал, встаёте в клетку первого ученика и нажимаете Enter (или F8);
   программа читает экран и сверяет предмет/класс/фамилии; только потом набирает. Esc — остановить.
   ============================================================ */
function buildEntryGroups(data) {
  const map = new Map();
  data.calendar.students.forEach(st => {
    const key = st.courseId ? "course:" + st.courseId : "class:" + st.class + ":" + (st.classLetterId || "");
    if (!map.has(key)) map.set(key, {
      key,
      courseId: st.courseId || null,
      grade: st.class || null,
      letterId: st.classLetterId || null,
      list: []
    });
    map.get(key).list.push(st);
  });
  const out = [];
  map.forEach(g => {
    const students = [...g.list].sort((a, b) => (a.name || "").localeCompare(b.name || "", "ru"));
    const ids = new Set(students.map(x => x.id));
    const marks = data.calendar.marks.filter(m => ids.has(m.studentId) && m.date && (m.scale || "5") === "5" && !m.posted);
    const dates = Array.from(new Set(marks.map(m => m.date))).sort();
    if (!dates.length) return;
    const letter = g.letterId ? classLetterName(data, g.letterId) : "";
    const label = g.courseId ? (data.courses.find(c => c.id === g.courseId) || {}).name || "Курс" : g.grade + letter + " класс";
    out.push({
      key: g.key,
      label,
      classLabel: g.courseId ? "" : letter ? g.grade + letter : "",
      isCourse: !!g.courseId,
      grade: g.grade,
      courseId: g.courseId,
      students,
      marks,
      dates
    });
  });
  return out.sort((a, b) => a.label.localeCompare(b.label, "ru", {
    numeric: true
  }));
}
// отметка → клавиши: «Н» вставляется из буфера обмена, цифры набираются; всё остальное пропускается
function entryTokenOf(value) {
  const v = String(value == null ? "" : value).trim();
  if (v === "Н") return "N";
  return /^[0-9]{1,3}$/.test(v) ? v : "";
}
function EntryQueueModal({
  data,
  update,
  onClose,
  focus
}) {
  const groups = buildEntryGroups(data);
  const recentFrom = (() => {
    const x = new Date();
    x.setDate(x.getDate() - 14);
    return isoOfDate(x);
  })(); // по умолчанию отмечены даты за последние 14 дней — старые чаще всего уже внесены, их отмечайте сами
  const inFocus = g => !focus || (focus.kind === "course" ? g.courseId === focus.value : !g.isCourse && String(g.grade) === String(focus.value));
  const [sel, setSel] = useState(() => Object.fromEntries(groups.map(g => [g.key, {
    on: inFocus(g),
    subject: g.isCourse ? "" : "Литература",
    checkClass: false,
    classLabel: g.classLabel,
    dates: Object.fromEntries(g.dates.map(d => [d, d >= recentFrom]))
  }])));
  const [trigger, setTrigger] = useState("Enter");
  const [verify, setVerify] = useState(true);
  const [stepMs, setStepMs] = useState(200);
  const [running, setRunning] = useState(false);
  const [prog, setProg] = useState(null);
  const [typedIds, setTypedIds] = useState([]);
  const [msg, setMsg] = useState("");
  const [ocrInfo, setOcrInfo] = useState(null);
  const stepsRef = useRef([]);
  const inp = {
    padding: "4px 8px",
    border: "1.5px solid var(--line)",
    borderRadius: 8,
    background: "var(--surface)"
  };
  useEffect(() => {
    if (!window.litAPI.onAutoTypeProgress) return undefined;
    return window.litAPI.onAutoTypeProgress(pr => {
      setProg(pr);
      if (pr.phase === "typed") setTypedIds(p => [...p, pr.id]);
    });
  }, []);
  const buildSteps = () => {
    const steps = [];
    groups.forEach(g => {
      const st = sel[g.key];
      if (!st || !st.on) return;
      g.dates.forEach(d => {
        if (!st.dates[d]) return;
        const rows = g.students.map(s => {
          const ms = g.marks.filter(m => m.studentId === s.id && m.date === d);
          return {
            s,
            ms,
            token: ms.length ? entryTokenOf(ms[0].value) : ""
          };
        });
        steps.push({
          id: g.key + "|" + d,
          title: g.label + " · " + formatDateRu(d).slice(0, 5),
          subject: st.subject.trim(),
          classLabel: st.checkClass ? st.classLabel.trim() : "",
          names: g.students.slice(0, 2).map(x => x.name),
          tokens: rows.map(r => r.token),
          firstName: g.students[0] ? g.students[0].name : "",
          markIds: rows.flatMap(r => r.ms.map(m => m.id)),
          multi: rows.filter(r => r.ms.length > 1).length
        });
      });
    });
    return steps;
  };
  const steps = buildSteps();
  const totalMarks = steps.reduce((n, x) => n + x.tokens.filter(Boolean).length, 0);
  const start = async () => {
    if (!steps.length) return;
    stepsRef.current = steps;
    setTypedIds([]);
    setMsg("");
    setProg(null);
    setRunning(true);
    try {
      const res = await window.litAPI.autoTypeQueue({
        steps: steps.map(({
          id,
          title,
          subject,
          classLabel,
          names,
          tokens
        }) => ({
          id,
          title,
          subject,
          classLabel,
          names,
          tokens
        })),
        trigger,
        verify,
        stepMs
      });
      if (res.ok) setMsg("Очередь выполнена. Проверьте журналы и отметьте «✓ Внесено» для шагов, где нажали «Сохранить».");else if (res.aborted) setMsg(res.error || "Очередь остановлена.");else setMsg("Не получилось: " + (res.error || "неизвестная ошибка"));
    } catch (e) {
      setMsg("Не получилось: " + (e && e.message || e));
    }
    setRunning(false);
  };
  const markPosted = ids => update(d => {
    const set = new Set(ids);
    d.calendar.marks.forEach(m => {
      if (set.has(m.id)) m.posted = true;
    });
    return d;
  });
  const ocrTest = async () => {
    setOcrInfo({
      busy: true
    });
    const r = await window.litAPI.autoTypeOcrTest(6000);
    setOcrInfo(r.ok ? {
      lang: r.lang,
      lines: (r.lines || []).slice(0, 25)
    } : {
      error: r.error || "не удалось"
    });
  };
  const curStep = prog && prog.index != null ? stepsRef.current[prog.index] : null;
  const doneSteps = stepsRef.current.filter(x => typedIds.includes(x.id));
  const phaseText = () => {
    if (!prog) return "";
    const k = prog.trigger || trigger;
    if (prog.phase === "waiting" && curStep) return "Шаг " + (prog.index + 1) + " из " + prog.total + ": «" + curStep.title + "». Откройте в журнале нужный класс и предмет" + (curStep.subject ? " («" + curStep.subject + "»)" : "") + ", щёлкните в клетку первого ученика (" + curStep.firstName + ") в столбце этой даты и нажмите " + k + ". Esc — остановить.";
    if (prog.phase === "checking") return "Сверяю экран…";
    if (prog.phase === "typing") return "Набираю… не трогайте мышь и клавиатуру. Esc — остановить.";
    return "";
  };
  const upd = (key, patch) => setSel(p => ({
    ...p,
    [key]: {
      ...p[key],
      ...patch
    }
  }));
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    style: {
      zIndex: 1260
    },
    onClick: e => e.target === e.currentTarget && !running && onClose()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 820
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, "\u2328 \u0410\u0432\u0442\u043E\u043D\u0430\u0431\u043E\u0440 \u043E\u0446\u0435\u043D\u043E\u043A \u0432 \u0436\u0443\u0440\u043D\u0430\u043B \xAB\u041C\u043E\u044F \u0448\u043A\u043E\u043B\u0430\xBB (\u043E\u0447\u0435\u0440\u0435\u0434\u044C)"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: running,
    onClick: onClose
  }, "\u0417\u0430\u043A\u0440\u044B\u0442\u044C")), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      display: "grid",
      gap: 12
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontSize: 12.5,
      lineHeight: 1.55
    }
  }, "\u041E\u0447\u0435\u0440\u0435\u0434\u044C \xAB\u043A\u043B\u0430\u0441\u0441 \u2192 \u0434\u0430\u0442\u0430\xBB. \u041D\u0430 \u043A\u0430\u0436\u0434\u043E\u043C \u0448\u0430\u0433\u0435 \u0432\u044B \u0441\u0430\u043C\u0438 \u043E\u0442\u043A\u0440\u044B\u0432\u0430\u0435\u0442\u0435 \u0432 \u0431\u0440\u0430\u0443\u0437\u0435\u0440\u0435 \u0436\u0443\u0440\u043D\u0430\u043B \u043D\u0443\u0436\u043D\u043E\u0433\u043E \u043A\u043B\u0430\u0441\u0441\u0430 \u0438 \u043F\u0440\u0435\u0434\u043C\u0435\u0442\u0430 \u0438 \u0432\u0441\u0442\u0430\u0451\u0442\u0435 (\u0449\u0435\u043B\u0447\u043A\u043E\u043C \u043C\u044B\u0448\u0438) \u0432 \u043A\u043B\u0435\u0442\u043A\u0443 \u043F\u0435\u0440\u0432\u043E\u0433\u043E \u0443\u0447\u0435\u043D\u0438\u043A\u0430 \u0432 \u0441\u0442\u043E\u043B\u0431\u0446\u0435 \u0434\u0430\u0442\u044B, \u0437\u0430\u0442\u0435\u043C \u043D\u0430\u0436\u0438\u043C\u0430\u0435\u0442\u0435 ", /*#__PURE__*/React.createElement("b", null, trigger), ". \u041F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0447\u0438\u0442\u0430\u0435\u0442 \u044D\u043A\u0440\u0430\u043D \u0438 \u0441\u0432\u0435\u0440\u044F\u0435\u0442 \u043F\u0440\u0435\u0434\u043C\u0435\u0442, \u043A\u043B\u0430\u0441\u0441 \u0438 \u0444\u0430\u043C\u0438\u043B\u0438\u0438 \u043F\u0435\u0440\u0432\u044B\u0445 \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432; \u0435\u0441\u043B\u0438 \u0432\u0441\u0451 \u0441\u043E\u0432\u043F\u0430\u043B\u043E \u2014 \u043D\u0430\u0431\u0438\u0440\u0430\u0435\u0442 \u043E\u0446\u0435\u043D\u043A\u0438 \u0441\u0442\u043E\u043B\u0431\u0446\u043E\u043C (\u043E\u0446\u0435\u043D\u043A\u0430 \u2192 \u0441\u0442\u0440\u0435\u043B\u043A\u0430 \u0432\u043D\u0438\u0437 \u2192 \u2026). \u041F\u043E\u0442\u043E\u043C \u0432\u044B \u043F\u0440\u043E\u0432\u0435\u0440\u044F\u0435\u0442\u0435 \u0441\u0442\u043E\u043B\u0431\u0435\u0446 \u0438 \u0441\u0430\u043C\u0438 \u043D\u0430\u0436\u0438\u043C\u0430\u0435\u0442\u0435 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435 \xAB\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C\xBB. ", /*#__PURE__*/React.createElement("b", null, "Esc \u2014 \u043E\u0441\u0442\u0430\u043D\u043E\u0432\u0438\u0442\u044C."), " \u0423\u0447\u0435\u043D\u0438\u043A\u0438 \u0438\u0434\u0443\u0442 \u043F\u043E \u0430\u043B\u0444\u0430\u0432\u0438\u0442\u0443 \u2014 \u043A\u0430\u043A \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435. \u0414\u0430\u0442\u044B \u0441\u0442\u0430\u0440\u0448\u0435 14 \u0434\u043D\u0435\u0439 \u043F\u043E \u0443\u043C\u043E\u043B\u0447\u0430\u043D\u0438\u044E \u043D\u0435 \u043E\u0442\u043C\u0435\u0447\u0435\u043D\u044B \u2014 \u043E\u0442\u043C\u0435\u0442\u044C\u0442\u0435 \u0441\u0430\u043C\u0438, \u0435\u0441\u043B\u0438 \u043E\u043D\u0438 \u043F\u0440\u0430\u0432\u0434\u0430 \u0435\u0449\u0451 \u043D\u0435 \u0432\u043D\u0435\u0441\u0435\u043D\u044B."), !running && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      alignItems: "center",
      fontSize: 12.5
    }
  }, "\u041D\u0430\u0447\u0438\u043D\u0430\u0442\u044C \u043F\u043E \u043A\u043B\u0430\u0432\u0438\u0448\u0435:", /*#__PURE__*/React.createElement(Btn, {
    kind: trigger === "Enter" ? "primary" : "ghost",
    sm: true,
    onClick: () => setTrigger("Enter")
  }, "Enter"), /*#__PURE__*/React.createElement(Btn, {
    kind: trigger === "F8" ? "primary" : "ghost",
    sm: true,
    onClick: () => setTrigger("F8")
  }, "F8"), /*#__PURE__*/React.createElement(ToggleSwitch, {
    checked: verify,
    onChange: setVerify,
    label: "\u0421\u0432\u0435\u0440\u044F\u0442\u044C \u0436\u0443\u0440\u043D\u0430\u043B \u043F\u043E \u044D\u043A\u0440\u0430\u043D\u0443",
    small: true
  }), "\u0421\u043A\u043E\u0440\u043E\u0441\u0442\u044C:", /*#__PURE__*/React.createElement("select", {
    value: stepMs,
    onChange: e => setStepMs(+e.target.value),
    style: inp
  }, /*#__PURE__*/React.createElement("option", {
    value: 120
  }, "\u0431\u044B\u0441\u0442\u0440\u043E"), /*#__PURE__*/React.createElement("option", {
    value: 200
  }, "\u043E\u0431\u044B\u0447\u043D\u043E"), /*#__PURE__*/React.createElement("option", {
    value: 400
  }, "\u043C\u0435\u0434\u043B\u0435\u043D\u043D\u043E")), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: ocrTest,
    title: "\u0427\u0435\u0440\u0435\u0437 6 \u0441\u0435\u043A\u0443\u043D\u0434 \u043F\u0440\u043E\u0447\u0438\u0442\u0430\u0435\u0442 \u043E\u043A\u043D\u043E, \u043A\u043E\u0442\u043E\u0440\u043E\u0435 \u0432 \u0444\u043E\u043A\u0443\u0441\u0435 (\u043E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0436\u0443\u0440\u043D\u0430\u043B) \u2014 \u043F\u043E\u043A\u0430\u0436\u0435\u0442, \u0447\u0442\u043E \u0432\u0438\u0434\u0438\u0442 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430"
  }, "\uD83D\uDD0E \u0422\u0435\u0441\u0442 \u0440\u0430\u0441\u043F\u043E\u0437\u043D\u0430\u0432\u0430\u043D\u0438\u044F (6 \u0441)")), trigger === "Enter" && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "Enter \u043F\u0440\u0438 \u044D\u0442\u043E\u043C \u043F\u043E\u043F\u0430\u0434\u0430\u0435\u0442 \u0438 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B; \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u0430\u043C\u0430 \u043A\u043E\u0440\u0440\u0435\u043A\u0442\u0438\u0440\u0443\u0435\u0442 \u043F\u043E\u043B\u043E\u0436\u0435\u043D\u0438\u0435 (\u043E\u0434\u0438\u043D \u0440\u0430\u0437 \u0441\u0442\u0440\u0435\u043B\u043A\u0430 \u0432\u0432\u0435\u0440\u0445). \u0415\u0441\u043B\u0438 \u0441 Enter \u0447\u0442\u043E-\u0442\u043E \u0432\u0435\u0434\u0451\u0442 \u0441\u0435\u0431\u044F \u0441\u0442\u0440\u0430\u043D\u043D\u043E \u2014 \u0432\u044B\u0431\u0435\u0440\u0438\u0442\u0435 F8."), ocrInfo && /*#__PURE__*/React.createElement("div", {
    className: "qrow",
    style: {
      cursor: "default",
      fontSize: 12
    }
  }, ocrInfo.busy ? "Через 6 секунд читаю окно в фокусе — переключитесь на журнал…" : ocrInfo.error ? "Распознавание недоступно: " + ocrInfo.error : "Распознано (" + ocrInfo.lang + "): " + ocrInfo.lines.join(" | ")), groups.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "qrow",
    style: {
      cursor: "default",
      color: "#2e8b57"
    }
  }, "\u2713 \u041D\u0435\u0432\u043D\u0435\u0441\u0451\u043D\u043D\u044B\u0445 \u043E\u0442\u043C\u0435\u0442\u043E\u043A (5-\u0431\u0430\u043B\u043B\u044C\u043D\u044B\u0445) \u043D\u0435\u0442."), groups.map(g => {
    const st = sel[g.key];
    if (!st) return null;
    return /*#__PURE__*/React.createElement("div", {
      key: g.key,
      className: "qrow",
      style: {
        cursor: "default",
        display: "grid",
        gap: 6,
        opacity: st.on ? 1 : 0.55
      }
    }, /*#__PURE__*/React.createElement("label", {
      style: {
        display: "flex",
        gap: 8,
        alignItems: "center",
        fontWeight: 800,
        cursor: "pointer"
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: st.on,
      onChange: e => upd(g.key, {
        on: e.target.checked
      })
    }), g.label, " ", /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, "\xB7 \u0434\u0430\u0442: ", g.dates.length, ", \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432: ", g.students.length)), st.on && /*#__PURE__*/React.createElement("div", {
      style: {
        marginLeft: 26,
        display: "grid",
        gap: 6
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        gap: 8,
        alignItems: "center",
        fontSize: 12.5
      }
    }, "\u041F\u0440\u0435\u0434\u043C\u0435\u0442 \u043D\u0430 \u044D\u043A\u0440\u0430\u043D\u0435:", /*#__PURE__*/React.createElement("input", {
      style: {
        ...inp,
        maxWidth: 190
      },
      value: st.subject,
      placeholder: "\u043D\u0435 \u043F\u0440\u043E\u0432\u0435\u0440\u044F\u0442\u044C",
      onChange: e => upd(g.key, {
        subject: e.target.value
      })
    }), g.classLabel && /*#__PURE__*/React.createElement("label", {
      style: {
        display: "flex",
        gap: 5,
        alignItems: "center",
        cursor: "pointer"
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: st.checkClass,
      onChange: e => upd(g.key, {
        checkClass: e.target.checked
      })
    }), "\u043F\u0440\u043E\u0432\u0435\u0440\u044F\u0442\u044C \u043A\u043B\u0430\u0441\u0441 \xAB", g.classLabel, "\xBB \u043D\u0430 \u044D\u043A\u0440\u0430\u043D\u0435")), /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        gap: 10,
        fontSize: 12.5
      }
    }, g.dates.map(d => /*#__PURE__*/React.createElement("label", {
      key: d,
      style: {
        display: "flex",
        gap: 4,
        alignItems: "center",
        cursor: "pointer"
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: !!st.dates[d],
      onChange: e => upd(g.key, {
        dates: {
          ...st.dates,
          [d]: e.target.checked
        }
      })
    }), formatDateRu(d).slice(0, 5), " (", g.marks.filter(m => m.date === d).length, ")")))));
  }), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 10,
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: !steps.length || !totalMarks,
    onClick: start
  }, "\u25B6 \u0417\u0430\u043F\u0443\u0441\u0442\u0438\u0442\u044C \u043E\u0447\u0435\u0440\u0435\u0434\u044C (", steps.length, " \u0448\u0430\u0433\u043E\u0432, ", totalMarks, " \u043E\u0442\u043C\u0435\u0442\u043E\u043A)"), steps.some(x => x.multi) && /*#__PURE__*/React.createElement("span", {
    style: {
      fontSize: 12,
      color: "var(--neg)"
    }
  }, "\u0413\u0434\u0435 \u0443 \u0443\u0447\u0435\u043D\u0438\u043A\u0430 \u043D\u0430 \u0434\u0430\u0442\u0443 \u0431\u043E\u043B\u044C\u0448\u0435 \u043E\u0434\u043D\u043E\u0439 \u043E\u0442\u043C\u0435\u0442\u043A\u0438, \u043D\u0430\u0431\u0438\u0440\u0430\u0435\u0442\u0441\u044F \u043F\u0435\u0440\u0432\u0430\u044F \u2014 \u043E\u0441\u0442\u0430\u043B\u044C\u043D\u044B\u0435 \u0432\u043D\u0435\u0441\u0438\u0442\u0435 \u0432\u0440\u0443\u0447\u043D\u0443\u044E."))), (running || prog) && /*#__PURE__*/React.createElement("div", {
    className: "qrow",
    style: {
      cursor: "default",
      display: "grid",
      gap: 8,
      background: running ? "#fffaf0" : undefined
    }
  }, running && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 14,
      lineHeight: 1.55
    }
  }, phaseText()), prog && prog.phase === "mismatch" && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13,
      color: "var(--neg)",
      lineHeight: 1.5
    }
  }, "\u26A0 \u0416\u0443\u0440\u043D\u0430\u043B \u043D\u0435 \u0441\u043E\u0432\u043F\u0430\u043B \u0441 \u0448\u0430\u0433\u043E\u043C \xAB", curStep ? curStep.title : "", "\xBB. \u041D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u043E \u043D\u0430 \u044D\u043A\u0440\u0430\u043D\u0435: ", /*#__PURE__*/React.createElement("b", null, (prog.failed || []).join(", ")), ".", /*#__PURE__*/React.createElement("br", null), "\u041E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u043F\u0440\u0430\u0432\u0438\u043B\u044C\u043D\u044B\u0439 \u0436\u0443\u0440\u043D\u0430\u043B \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435 ", prog.trigger || trigger, " \u0435\u0449\u0451 \u0440\u0430\u0437; ", /*#__PURE__*/React.createElement("b", null, "F9"), " \u2014 \u043D\u0430\u0431\u0440\u0430\u0442\u044C \u0432\u0441\u0451 \u0440\u0430\u0432\u043D\u043E; Esc \u2014 \u043E\u0441\u0442\u0430\u043D\u043E\u0432\u0438\u0442\u044C.", prog.seen && prog.seen.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 4
    }
  }, "\u041F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0432\u0438\u0434\u0438\u0442: ", prog.seen.join(" | "))), running && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement(Btn, {
    kind: "danger",
    sm: true,
    onClick: () => window.litAPI.autoTypeCancel()
  }, "\u25A0 \u041E\u0441\u0442\u0430\u043D\u043E\u0432\u0438\u0442\u044C")), msg && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13,
      color: msg.startsWith("Не получилось") ? "var(--neg)" : "#2e8b57"
    }
  }, msg), doneSteps.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gap: 4
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700,
      fontSize: 13
    }
  }, "\u041D\u0430\u0431\u0440\u0430\u043D\u043E \u0448\u0430\u0433\u043E\u0432: ", doneSteps.length), doneSteps.map(x => /*#__PURE__*/React.createElement("div", {
    key: x.id,
    className: "row",
    style: {
      gap: 8,
      alignItems: "center",
      fontSize: 12.5
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      flex: 1
    }
  }, "\u2713 ", x.title), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => markPosted(x.markIds),
    title: "\u041D\u0430\u0436\u043C\u0438\u0442\u0435 \u043F\u043E\u0441\u043B\u0435 \xAB\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C\xBB \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435 \u2014 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u043F\u043E\u043B\u0443\u0447\u0430\u0442 \u043F\u043E\u043C\u0435\u0442\u043A\u0443 \xAB\u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u0430\xBB"
  }, "\u2713 \u0412\u043D\u0435\u0441\u0435\u043D\u043E"))), doneSteps.length > 1 && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => markPosted(doneSteps.flatMap(x => x.markIds))
  }, "\u2713 \u041E\u0442\u043C\u0435\u0442\u0438\u0442\u044C \u0432\u043D\u0435\u0441\u0451\u043D\u043D\u044B\u043C\u0438 \u0432\u0441\u0435 \u043D\u0430\u0431\u0440\u0430\u043D\u043D\u044B\u0435")))))));
}

/* ============================================================
   СЕГОДНЯ: уроки дня, кого вызвать, что не выставлено в эл. журнал
   ============================================================ */
function TodayTab({
  data,
  update,
  onOpenLesson,
  onOpenJournal,
  hist
}) {
  const today = todayIso();
  const [touched, setTouched] = useState([]); // отметки, переключённые на этой странице, остаются в списке (серыми), пока вы не уйдёте
  const [queueOpen, setQueueOpen] = useState(false);
  const canQueue = !!(window.litAPI && window.litAPI.autoTypeQueue);
  const setPostedMarks = (ids, on) => {
    setTouched(p => Array.from(new Set([...p, ...ids])));
    update(d => {
      const set = new Set(ids);
      d.calendar.marks.forEach(m => {
        if (set.has(m.id)) m.posted = on;
      });
      return d;
    });
  };
  const tomD = new Date();
  tomD.setDate(tomD.getDate() + 1);
  const tomorrow = isoOfDate(tomD);
  const agoD = new Date();
  agoD.setDate(agoD.getDate() - 14);
  const ago = isoOfDate(agoD);
  const dateOf = l => l.actualDate || l.plannedDate || "";
  const byDate = day => {
    const ls = data.lessons.filter(l => isLessonInPlay(l) && dateOf(l) === day);
    const {
      map,
      emptySlots
    } = assignLessonNumbers(data, day, ls);
    ls.sort((a, b) => (map.get(a.id) || 99) - (map.get(b.id) || 99) || lessonPlaceLabel(a, data).localeCompare(lessonPlaceLabel(b, data), "ru") || a.order - b.order);
    return {
      list: ls,
      nums: map,
      emptySlots
    };
  };
  const todayInfo = byDate(today),
    tomorrowInfo = byDate(tomorrow);
  const overdue = data.lessons.filter(l => !l.archived && !l.completed && !l.excludeFromSchedule && !l.actualDate && l.plannedDate && l.plannedDate < today).sort((a, b) => a.plannedDate < b.plannedDate ? 1 : -1).slice(0, 8);
  const trackOf = l => ({
    trackKey: l.courseId ? "course:" + l.courseId : "class:" + l.class,
    letter: l.classLetterId || "all"
  });
  const toggleDone = lesson => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    if (!l) return d;
    setLessonCompleted(l, !l.completed);
    return d;
  });

  // Не выставлено в электронный журнал: отметки за последние 14 дней без флажка «выставлена», по группам и ученикам
  const groups = {};
  data.calendar.marks.forEach(m => {
    if (!m.date || m.date < ago || m.date > today) return;
    if (m.posted && !touched.includes(m.id)) return;
    const st = data.calendar.students.find(x => x.id === m.studentId);
    if (!st) return;
    const key = st.courseId ? "course:" + st.courseId : "class:" + st.class + ":" + (st.classLetterId || "");
    if (!groups[key]) groups[key] = {
      key,
      count: 0,
      people: {},
      label: st.courseId ? (data.courses.find(c => c.id === st.courseId) || {}).name || "Курс" : st.class + (st.classLetterId ? classLetterName(data, st.classLetterId) : "") + " класс",
      trackKey: st.courseId ? "course:" + st.courseId : "class:" + st.class,
      letter: st.courseId ? "all" : st.classLetterId || "none"
    };
    if (!m.posted) groups[key].count++;
    (groups[key].people[st.name] = groups[key].people[st.name] || []).push(m);
  });
  const groupList = Object.values(groups).sort((a, b) => a.label.localeCompare(b.label, "ru", {
    numeric: true
  }));
  const LessonCardT = ({
    l,
    num
  }) => {
    const roster = data.calendar.students.filter(s => (s.class || null) === (l.class || null) && (s.courseId || null) === (l.courseId || null) && (s.classLetterId || null) === (l.classLetterId || null)).sort((a, b) => (a.name || "").localeCompare(b.name || "", "ru"));
    const rec = recommendToCall(data, roster, settingsForLesson(data, l).gradingSystem === "100" ? "100" : "5", settingsForLesson(data, l).requiredMarks);
    const exCount = (l.items || []).filter(it => it.type === "exercise").length;
    return /*#__PURE__*/React.createElement("div", {
      className: "qrow",
      style: {
        cursor: "default"
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        justifyContent: "space-between",
        gap: 8
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        fontWeight: 800,
        fontSize: 14.5
      }
    }, num ? /*#__PURE__*/React.createElement("span", {
      style: {
        color: "var(--gold)"
      }
    }, num, " \u0443\u0440\u043E\u043A \xB7 ") : null, lessonPlaceLabel(l, data), " \xB7 ", l.topic || "Без темы", l.duplicateOf ? " (дубликат)" : "", l.completed && /*#__PURE__*/React.createElement("span", {
      style: {
        color: "#2e8b57",
        marginLeft: 8,
        fontSize: 12.5
      }
    }, "\u2713 \u043F\u0440\u043E\u0432\u0435\u0434\u0451\u043D")), /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        gap: 6
      }
    }, /*#__PURE__*/React.createElement(Btn, {
      kind: "primary",
      sm: true,
      onClick: () => onOpenLesson(l)
    }, "\u041E\u0442\u043A\u0440\u044B\u0442\u044C \u0443\u0440\u043E\u043A"), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => onOpenJournal(trackOf(l))
    }, "\uD83D\uDCD2 \u0416\u0443\u0440\u043D\u0430\u043B"), /*#__PURE__*/React.createElement(Btn, {
      kind: "ghost",
      sm: true,
      onClick: () => toggleDone(l)
    }, l.completed ? "Снять «проведён»" : "✓ Урок проведён"))), /*#__PURE__*/React.createElement("div", {
      className: "hint",
      style: {
        marginTop: 3
      }
    }, exCount, " \u0443\u043F\u0440\u0430\u0436\u043D\u0435\u043D\u0438\u0439 \xB7 ", (l.items || []).length - exCount, " \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u043E\u0432 \u0438 \u0432\u043E\u043F\u0440\u043E\u0441\u043E\u0432", l.needsWork ? " · требует доработки" : ""), rec.length > 0 && /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        gap: 6,
        marginTop: 6,
        fontSize: 12.5
      }
    }, /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, "\u041A\u043E\u0433\u043E \u0432\u044B\u0437\u0432\u0430\u0442\u044C (\u043C\u0430\u043B\u043E \u043E\u0442\u043C\u0435\u0442\u043E\u043A / \u043D\u0438\u0437\u043A\u0438\u0435 \u043E\u0446\u0435\u043D\u043A\u0438):"), rec.map(r => /*#__PURE__*/React.createElement("span", {
      key: r.id,
      title: `Отметок в четверти: ${r._n}${r._avg != null ? ", средний: " + String(r._avg).replace(".", ",") : ""}${callFlagTitle(r._few || r._poor ? {
        few: r._few,
        poor: r._poor
      } : null)}`,
      style: {
        fontWeight: 700,
        background: "var(--accent-soft)",
        borderRadius: 6,
        padding: "2px 8px",
        border: "1.5px solid transparent",
        ...(callFlagStyle({
          few: r._few,
          poor: r._poor
        }) || {})
      }
    }, r.name)), rec.some(r => r._few || r._poor) && /*#__PURE__*/React.createElement(CallFlagLegend, null)));
  };
  const H = ({
    children
  }) => /*#__PURE__*/React.createElement("h2", {
    style: {
      fontFamily: "var(--font-display)",
      fontSize: 17,
      margin: "14px 0 6px"
    }
  }, children);
  const rows = [...todayInfo.list.map(l => ({
    n: todayInfo.nums.get(l.id) || 99,
    l
  })), ...todayInfo.emptySlots.map(x => ({
    n: x.num,
    slot: x
  }))].sort((a, b) => a.n - b.n);
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      justifyContent: "space-between",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("h1", {
    style: {
      fontFamily: "var(--font-display)",
      fontSize: 22,
      margin: 0
    }
  }, "\u0421\u0435\u0433\u043E\u0434\u043D\u044F, ", formatDateRu(today)), hist && /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !hist.canUndo,
    onClick: hist.undo,
    title: "\u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C \u043F\u043E\u0441\u043B\u0435\u0434\u043D\u0435\u0435 \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0435 \u043E\u0442\u043C\u0435\u0442\u043E\u043A (Ctrl+Z)"
  }, "\u21B6 \u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !hist.canRedo,
    onClick: hist.redo,
    title: "\u0412\u0435\u0440\u043D\u0443\u0442\u044C \u043E\u0442\u043C\u0435\u043D\u0451\u043D\u043D\u043E\u0435 (Ctrl+Y)"
  }, "\u21B7 \u0412\u0435\u0440\u043D\u0443\u0442\u044C"))), /*#__PURE__*/React.createElement(H, null, "\u0423\u0440\u043E\u043A\u0438 \u0441\u0435\u0433\u043E\u0434\u043D\u044F"), rows.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "empty",
    style: {
      padding: 22
    }
  }, "\u041D\u0430 \u0441\u0435\u0433\u043E\u0434\u043D\u044F \u0443\u0440\u043E\u043A\u043E\u0432 \u043D\u0435\u0442. \u0414\u0430\u0442\u044B \u0443\u0440\u043E\u043A\u043E\u0432 \u0437\u0430\u0434\u0430\u044E\u0442\u0441\u044F \u0432 \xAB\u041F\u043B\u0430\u043D\u0438\u0440\u043E\u0432\u0430\u043D\u0438\u0438 \u2192 \u041A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044C\xBB (\u0442\u0430\u043C \u0436\u0435 \xAB\u041C\u043E\u0451 \u0440\u0430\u0441\u043F\u0438\u0441\u0430\u043D\u0438\u0435\xBB)."), rows.map(r => r.l ? /*#__PURE__*/React.createElement(LessonCardT, {
    key: r.l.id,
    l: r.l,
    num: todayInfo.nums.get(r.l.id)
  }) : /*#__PURE__*/React.createElement("div", {
    key: r.slot.id,
    className: "qrow",
    style: {
      cursor: "default",
      borderStyle: "dashed",
      color: "var(--muted)"
    }
  }, /*#__PURE__*/React.createElement("b", null, r.slot.num, " \u0443\u0440\u043E\u043A"), " \xB7 ", slotLabel(data, r.slot), " \u2014 \u0443\u0440\u043E\u043A \u043D\u0430 \u0441\u0435\u0433\u043E\u0434\u043D\u044F \u043D\u0435 \u0437\u0430\u043F\u043B\u0430\u043D\u0438\u0440\u043E\u0432\u0430\u043D (\u0434\u043E\u0431\u0430\u0432\u044C\u0442\u0435 \u0443\u0440\u043E\u043A\u0438 \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB\u0417\u0430\u043F\u043E\u043B\u043D\u0438\u0442\u044C \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044C\xBB \u0432 \xAB\u041C\u043E\u0451\u043C \u0440\u0430\u0441\u043F\u0438\u0441\u0430\u043D\u0438\u0438\xBB)")), tomorrowInfo.list.length > 0 && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(H, null, "\u0417\u0430\u0432\u0442\u0440\u0430"), tomorrowInfo.list.map(l => /*#__PURE__*/React.createElement("div", {
    key: l.id,
    className: "qrow row wrap",
    style: {
      cursor: "default",
      justifyContent: "space-between"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontWeight: 700
    }
  }, tomorrowInfo.nums.get(l.id) ? tomorrowInfo.nums.get(l.id) + " урок · " : "", lessonPlaceLabel(l, data), " \xB7 ", l.topic || "Без темы", l.duplicateOf ? " (дубликат)" : "", l.needsWork ? " · требует доработки" : ""), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => onOpenLesson(l)
  }, "\u041E\u0442\u043A\u0440\u044B\u0442\u044C")))), /*#__PURE__*/React.createElement(H, null, "\u041D\u0435 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043E \u0432 \u044D\u043B\u0435\u043A\u0442\u0440\u043E\u043D\u043D\u044B\u0439 \u0436\u0443\u0440\u043D\u0430\u043B (\u0437\u0430 14 \u0434\u043D\u0435\u0439)"), canQueue && groupList.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: 8
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setQueueOpen(true),
    title: "\u041E\u0447\u0435\u0440\u0435\u0434\u044C \u043F\u043E \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u0438\u043C \u043A\u043B\u0430\u0441\u0441\u0430\u043C: \u0432\u044B \u043E\u0442\u043A\u0440\u044B\u0432\u0430\u0435\u0442\u0435 \u0436\u0443\u0440\u043D\u0430\u043B \u043A\u0430\u0436\u0434\u043E\u0433\u043E \u043A\u043B\u0430\u0441\u0441\u0430 \u2014 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u0432\u0435\u0440\u044F\u0435\u0442 \u044D\u043A\u0440\u0430\u043D \u0438 \u043D\u0430\u0431\u0438\u0440\u0430\u0435\u0442 \u043E\u0446\u0435\u043D\u043A\u0438"
  }, "\u2328 \u0412\u0435\u0447\u0435\u0440\u043D\u0438\u0439 \u043D\u0430\u0431\u043E\u0440 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B (\u043E\u0447\u0435\u0440\u0435\u0434\u044C \u043F\u043E \u043A\u043B\u0430\u0441\u0441\u0430\u043C)")), queueOpen && /*#__PURE__*/React.createElement(EntryQueueModal, {
    data: data,
    update: update,
    onClose: () => setQueueOpen(false)
  }), groupList.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "qrow",
    style: {
      cursor: "default",
      color: "#2e8b57"
    }
  }, "\u2713 \u0412\u0441\u0435 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u0437\u0430 \u043F\u043E\u0441\u043B\u0435\u0434\u043D\u0438\u0435 \u0434\u0432\u0435 \u043D\u0435\u0434\u0435\u043B\u0438 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u044B.") : groupList.map(g => /*#__PURE__*/React.createElement("div", {
    key: g.key,
    className: "qrow",
    style: {
      cursor: "default"
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      justifyContent: "space-between",
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontWeight: 800
    }
  }, g.label, " ", g.count > 0 ? /*#__PURE__*/React.createElement("span", {
    style: {
      color: "#c0392b"
    }
  }, "\xB7 \u043D\u0435 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043E \u043E\u0442\u043C\u0435\u0442\u043E\u043A: ", g.count) : /*#__PURE__*/React.createElement("span", {
    style: {
      color: "#2e8b57"
    }
  }, "\xB7 \u0432\u0441\u0451 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043E \u2713")), /*#__PURE__*/React.createElement("span", {
    className: "row wrap",
    style: {
      gap: 6
    }
  }, g.count > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setPostedMarks(Object.values(g.people).flat().filter(m => !m.posted).map(m => m.id), true),
    title: "\u041E\u0442\u043C\u0435\u0442\u0438\u0442\u044C \u0432\u0441\u0435 \u044D\u0442\u0438 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043D\u044B\u043C\u0438"
  }, "\u2713 \u0412\u0441\u0435 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u044B"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: () => onOpenJournal({
      trackKey: g.trackKey,
      letter: g.letter
    })
  }, "\u041E\u0442\u043A\u0440\u044B\u0442\u044C \u0436\u0443\u0440\u043D\u0430\u043B"))), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gap: 6,
      marginTop: 6
    }
  }, Object.entries(g.people).sort((a, b) => a[0].localeCompare(b[0], "ru")).map(([name, ms]) => /*#__PURE__*/React.createElement("div", {
    key: name,
    className: "row wrap",
    style: {
      gap: 8,
      borderTop: "1px solid var(--line)",
      paddingTop: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontWeight: 700,
      minWidth: 150
    }
  }, name), [...ms].sort((a, b) => a.date < b.date ? -1 : 1).map(m => /*#__PURE__*/React.createElement("span", {
    key: m.id,
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: 7,
      border: "1.5px solid var(--line)",
      borderRadius: 9,
      padding: "3px 8px",
      background: m.posted ? "#e3e5e8" : "#fff"
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "jpill",
    style: {
      opacity: m.posted ? 0.7 : 1,
      background: m.value === "Н" ? m.excused ? "#8593a3" : "#c0392b" : markColorFor(m.value, m.scale)
    }
  }, m.value), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, formatDateRu(m.date).slice(0, 5)), /*#__PURE__*/React.createElement(ToggleSwitch, {
    small: true,
    checked: !!m.posted,
    onChange: v => setPostedMarks([m.id], v),
    label: "\u0412\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u0430"
  })))))))), overdue.length > 0 && /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(H, null, "\u041D\u0435 \u043F\u0440\u043E\u0432\u0435\u0434\u0435\u043D\u044B (\u0434\u0430\u0442\u0430 \u0443\u0436\u0435 \u043F\u0440\u043E\u0448\u043B\u0430)"), overdue.map(l => /*#__PURE__*/React.createElement("div", {
    key: l.id,
    className: "qrow row wrap",
    style: {
      cursor: "default",
      justifyContent: "space-between"
    }
  }, /*#__PURE__*/React.createElement("span", null, /*#__PURE__*/React.createElement("b", null, lessonPlaceLabel(l, data)), " \xB7 ", l.topic || "Без темы", " ", /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\xB7 \u043F\u043B\u0430\u043D: ", formatDateRu(l.plannedDate))), /*#__PURE__*/React.createElement("span", {
    className: "row",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => onOpenLesson(l)
  }, "\u041E\u0442\u043A\u0440\u044B\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => toggleDone(l)
  }, "\u2713 \u041F\u0440\u043E\u0432\u0435\u0434\u0451\u043D"))))));
}

// ---- «Моё расписание» ----
// Номер урока действителен, только пока урок стоит на той дате, для которой он был назначен.
function lessonNoOf(l) {
  return l.lessonNo && l.lessonNoDate === l.plannedDate ? l.lessonNo : null;
}
// Какая это неделя: "A" (нечётная) или "B" (чётная). В режиме «каждую неделю» всегда "A".
function scheduleWeekOf(data, iso) {
  const sc = data.schedule;
  if (!sc || sc.mode !== "biweekly") return "A";
  const a = new Date(mondayOf(sc.weekAAnchor || iso) + "T00:00:00");
  const b = new Date(mondayOf(iso) + "T00:00:00");
  const weeks = Math.round((b - a) / (7 * 86400000));
  return (weeks % 2 + 2) % 2 === 0 ? "A" : "B";
}
function scheduleSlotsFor(data, iso) {
  const sc = data.schedule;
  if (!sc || !sc.slots.length) return [];
  const dow = (new Date(iso + "T00:00:00").getDay() + 6) % 7 + 1;
  const wk = scheduleWeekOf(data, iso);
  return sc.slots.filter(x => x.week === wk && x.dow === dow).sort((a, b) => a.num - b.num);
}
function slotGroupKey(sl) {
  return sl.kind + ":" + sl.value + ":" + (sl.classLetterId || "");
}
function lessonMatchesSlot(l, sl) {
  return sl.kind === "course" ? l.courseId === sl.value : !l.courseId && l.class === sl.value && (l.classLetterId || "") === (sl.classLetterId || "");
}
function slotLabel(data, sl) {
  if (sl.kind === "course") return data.courses.find(c => c.id === sl.value)?.name || "Курс";
  return sl.value + (sl.classLetterId ? classLetterName(data, sl.classLetterId) : "") + (sl.classLetterId ? "" : " класс");
}
// Номера уроков на дату: у урока с действительным номером — он; остальным подбирается первый свободный слот его класса.
function assignLessonNumbers(data, iso, lessons) {
  const slots = scheduleSlotsFor(data, iso);
  const used = new Set();
  const map = new Map();
  lessons.forEach(l => {
    const n = lessonNoOf(l);
    if (n && slots.some(x => x.num === n && lessonMatchesSlot(l, x))) {
      map.set(l.id, n);
      used.add(n);
    }
  });
  lessons.forEach(l => {
    if (map.has(l.id)) return;
    const x = slots.find(y => !used.has(y.num) && lessonMatchesSlot(l, y));
    if (x) {
      map.set(l.id, x.num);
      used.add(x.num);
    }
  });
  return {
    map,
    emptySlots: slots.filter(x => !used.has(x.num))
  };
}
// Уроки дня в порядке расписания (для календаря): { list, nums }.
function lessonsOfDayByNumber(data, iso, lessons) {
  const {
    map
  } = assignLessonNumbers(data, iso, lessons.filter(isLessonInPlay));
  return {
    list: [...lessons].sort((a, b) => (map.get(a.id) || 99) - (map.get(b.id) || 99) || a.order - b.order),
    nums: map
  };
}
// Уроки группы (класс/литера или курс) в порядке списка «Уроки»: отдельные уроки и уроки систем вперемешку.
function groupLessonsInOrder(d, slot) {
  const sample = {
    class: slot.kind === "class" ? slot.value : null,
    courseId: slot.kind === "course" ? slot.value : null,
    classLetterId: slot.classLetterId || null
  };
  const out = [];
  topLevelEntries(d, sample).sort((a, b) => a.order - b.order).forEach(e => {
    if (e.type === "lesson") {
      const l = d.lessons.find(x => x.id === e.id);
      if (l && !out.includes(l)) out.push(l);
    } else {
      const sy = d.lessonSystems.find(x => x.id === e.id);
      (sy ? sy.lessonIds : []).forEach(id => {
        const l = d.lessons.find(x => x.id === id);
        if (l && !out.includes(l)) out.push(l);
      });
    }
  });
  d.lessons.filter(l => !out.includes(l)).sort((a, b) => a.order - b.order).forEach(l => out.push(l));
  return out.filter(l => lessonMatchesSlot(l, slot));
}
// Расставляет непроведённые уроки по слотам расписания, начиная с сегодняшнего дня (или со дня после последнего
// уже стоящего в календаре урока группы). Каникулы пропускаются. mode "undated" — только уроки без даты;
// mode "all" — все ещё не проведённые уроки группы (их даты меняются). Прежние даты запоминаются для «Вернуть прежние даты».
function fillCalendarFromSchedule(d, today, mode) {
  const stats = {
    placed: 0,
    changed: 0,
    groups: [],
    leftover: []
  };
  const prev = [];
  const seen = new Map();
  d.schedule.slots.forEach(x => {
    if (!seen.has(slotGroupKey(x))) seen.set(slotGroupKey(x), x);
  });
  const sysTouched = new Set();
  seen.forEach((slot, gk) => {
    const ordered = groupLessonsInOrder(d, slot);
    const eligible = l => !l.archived && !l.completed && !l.excludeFromSchedule && !l.actualDate;
    const pending = ordered.filter(l => eligible(l) && (mode === "all" || !l.plannedDate));
    if (!pending.length) return;
    const pendIds = new Set(pending.map(l => l.id));
    const others = ordered.filter(l => !pendIds.has(l.id) && isLessonInPlay(l) && !l.excludeFromSchedule);
    const lastPlanned = mode === "all" ? "" : others.filter(l => l.plannedDate).reduce((mx, l) => l.plannedDate > mx ? l.plannedDate : mx, "");
    const start = lastPlanned && lastPlanned > today ? lastPlanned : today;
    let idx = 0;
    for (let i = 0; i < 420 && idx < pending.length; i++) {
      const date = addDaysIso(start, i);
      if (vacationAt(d, date)) continue;
      const slots = scheduleSlotsFor(d, date).filter(x => slotGroupKey(x) === gk);
      if (!slots.length) continue;
      const existing = others.filter(l => l.plannedDate === date);
      const usedNums = new Set(existing.map(l => lessonNoOf(l)).filter(Boolean));
      const unnumbered = existing.filter(l => !lessonNoOf(l)).length;
      const free = slots.filter(x => !usedNums.has(x.num));
      free.splice(0, unnumbered);
      free.forEach(x => {
        if (idx >= pending.length) return;
        const l = pending[idx++];
        if (l.plannedDate !== date || lessonNoOf(l) !== x.num) {
          prev.push({
            id: l.id,
            plannedDate: l.plannedDate || "",
            lessonNo: l.lessonNo || null,
            lessonNoDate: l.lessonNoDate || ""
          });
          if (l.plannedDate && l.plannedDate !== date) stats.changed++;
        }
        l.plannedDate = date;
        l.lessonNo = x.num;
        l.lessonNoDate = date;
        if (l.systemId) sysTouched.add(l.systemId);
        stats.placed++;
      });
    }
    stats.groups.push({
      label: slotLabel(d, slot),
      placed: idx
    });
    if (idx < pending.length) stats.leftover.push(slotLabel(d, slot) + ": " + (pending.length - idx));
    resyncOrderForGroup(d, pending[0]);
  });
  sysTouched.forEach(id => resyncSystemLessonsByDate(d, id));
  if (prev.length) d.schedule.undo = {
    at: nowISO(),
    items: prev
  };
  return stats;
}
// Свободный номер урока по расписанию на дату для этого урока (или null, если расписания/слота нет).
function freeScheduleSlot(data, lesson, date) {
  const slots = scheduleSlotsFor(data, date).filter(x => lessonMatchesSlot(lesson, x));
  if (!slots.length) return null;
  const others = data.lessons.filter(l => l.id !== lesson.id && isLessonInPlay(l) && l.plannedDate === date && sameLessonGroup(l, lesson));
  const usedNums = new Set(others.map(l => lessonNoOf(l)).filter(Boolean));
  const unnumbered = others.filter(l => !lessonNoOf(l)).length;
  const free = slots.filter(x => !usedNums.has(x.num));
  free.splice(0, unnumbered);
  return free.length ? free[0].num : null;
}
// Дата, на которую предлагаем поставить разархивируемый урок: ближайший свободный слот «Моего расписания»
// (с сегодняшнего дня, без каникул), а без расписания — день после последнего урока группы (но не раньше сегодня).
function suggestUnarchiveDate(data, lesson) {
  const today = todayIso();
  if ((data.schedule && data.schedule.slots || []).some(x => lessonMatchesSlot(lesson, x))) {
    for (let i = 0; i < 400; i++) {
      const date = addDaysIso(today, i);
      if (vacationAt(data, date)) continue;
      if (freeScheduleSlot(data, lesson, date)) return date;
    }
  }
  const dates = data.lessons.filter(l => l.id !== lesson.id && isLessonInPlay(l) && sameLessonGroup(l, lesson)).map(l => l.plannedDate).filter(Boolean).sort();
  let date = dates.length ? addDaysIso(dates[dates.length - 1], 1) : today;
  if (date < today) date = today;
  for (let i = 0; i < 60 && vacationAt(data, date); i++) date = addDaysIso(date, 1);
  return date;
}
function unarchiveLesson(d, lessonId, date) {
  const l = d.lessons.find(x => x.id === lessonId);
  if (!l) return;
  if (l.completed) {
    l.completed = false;
    l.actualDate = "";
  }
  l.archived = false;
  l.plannedDate = date;
  const no = freeScheduleSlot(d, l, date);
  l.lessonNo = no || null;
  l.lessonNoDate = no ? date : "";
  resyncAfterDateChange(d, l);
}
function undoScheduleFill(d) {
  const u = d.schedule.undo;
  if (!u) return 0;
  let n = 0;
  u.items.forEach(it => {
    const l = d.lessons.find(x => x.id === it.id);
    if (!l) return;
    l.plannedDate = it.plannedDate;
    l.lessonNo = it.lessonNo;
    l.lessonNoDate = it.lessonNoDate;
    n++;
    resyncAfterDateChange(d, l);
  });
  d.schedule.undo = null;
  return n;
}

// «Моё расписание»: недельная сетка «день × номер урока → класс/литера или курс»; режим «каждую неделю» или
// «нечётная/чётная»; автозаполнение календаря уроками.
function ScheduleModal({
  data,
  update,
  onClose
}) {
  const sc = data.schedule;
  const [wk, setWk] = useState("A");
  const [result, setResult] = useState(null);
  const [fillMode, setFillMode] = useState("undated");
  const NUMS = [1, 2, 3, 4, 5, 6, 7, 8];
  const week = sc.mode === "biweekly" ? wk : "A";
  const curWeek = scheduleWeekOf(data, todayIso());
  const options = [];
  CLASSES.forEach(c => {
    const letters = lettersForGrade(data, c);
    if (letters.length) {
      letters.forEach(cl => options.push({
        v: "class:" + c + ":" + cl.id,
        label: c + cl.name
      }));
      if (hasNoLetterItems(data, c)) options.push({
        v: "class:" + c + ":",
        label: c + " (без литеры)"
      });
    } else options.push({
      v: "class:" + c + ":",
      label: c + " класс"
    });
  });
  data.courses.forEach(c => options.push({
    v: "course:" + c.id + ":",
    label: c.name
  }));
  const slotAt = (dow, num) => sc.slots.find(x => x.week === week && x.dow === dow && x.num === num);
  const setSlot = (dow, num, val) => update(d => {
    d.schedule.slots = d.schedule.slots.filter(x => !(x.week === week && x.dow === dow && x.num === num));
    if (val) {
      const [kind, value, letter] = val.split(":");
      d.schedule.slots.push({
        id: uid(),
        week,
        dow,
        num,
        kind,
        value,
        classLetterId: letter || null
      });
    }
    return d;
  });
  const setMode = mode => update(d => {
    d.schedule.mode = mode;
    if (mode === "biweekly" && !d.schedule.weekAAnchor) d.schedule.weekAAnchor = mondayOf(todayIso());
    return d;
  });
  const setCurrent = w => update(d => {
    const mon = mondayOf(todayIso());
    d.schedule.weekAAnchor = w === "A" ? mon : addDaysIso(mon, -7);
    return d;
  });
  const copyWeek = async () => {
    const to = week === "A" ? "B" : "A";
    if (!(await Confirm("Скопировать расписание «" + (week === "A" ? "нечётной" : "чётной") + "» недели в «" + (to === "A" ? "нечётную" : "чётную") + "»? Её нынешнее расписание будет заменено."))) return;
    update(d => {
      const src = d.schedule.slots.filter(x => x.week === week);
      d.schedule.slots = d.schedule.slots.filter(x => x.week !== to).concat(src.map(x => ({
        ...x,
        id: uid(),
        week: to
      })));
      return d;
    });
  };
  const fill = async () => {
    const copy = structuredClone(data);
    const stats = fillCalendarFromSchedule(copy, todayIso(), fillMode);
    if (!stats.placed) {
      setResult(stats);
      return;
    }
    if (fillMode === "all" || stats.changed > 0) {
      if (!(await Confirm("Будет расставлено уроков: " + stats.placed + (stats.changed ? ", из них у " + stats.changed + " изменится уже стоящая дата" : "") + ". Прежние даты можно вернуть кнопкой «Вернуть прежние даты». Продолжить?", "Расставить"))) return;
    }
    update(d => {
      fillCalendarFromSchedule(d, todayIso(), fillMode);
      return d;
    });
    setResult(stats);
  };
  const undoFill = async () => {
    if (!(await Confirm("Вернуть даты и номера уроков, какими они были до последнего «Заполнить календарь»?", "Вернуть"))) return;
    update(d => {
      undoScheduleFill(d);
      return d;
    });
    setResult(null);
    showToast("Прежние даты возвращены");
  };
  const used = sc.slots.filter(x => x.week === week).length;
  const cellSel = {
    border: "1px solid var(--line)",
    borderRadius: 7,
    padding: "4px 6px",
    fontSize: 12.5,
    minWidth: 86
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 900
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u041C\u043E\u0451 \u0440\u0430\u0441\u043F\u0438\u0441\u0430\u043D\u0438\u0435"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, "\u0423\u043A\u0430\u0436\u0438\u0442\u0435, \u0432 \u043A\u0430\u043A\u043E\u0439 \u0434\u0435\u043D\u044C, \u043F\u043E\u0434 \u043A\u0430\u043A\u0438\u043C \u043D\u043E\u043C\u0435\u0440\u043E\u043C \u0438 \u0432 \u043A\u0430\u043A\u043E\u043C \u043A\u043B\u0430\u0441\u0441\u0435 \u0443 \u0432\u0430\u0441 \u0443\u0440\u043E\u043A. \u041F\u043E \u044D\u0442\u043E\u0439 \u0441\u0435\u0442\u043A\u0435 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0440\u0430\u0441\u043A\u043B\u0430\u0434\u044B\u0432\u0430\u0435\u0442 \u0443\u0440\u043E\u043A\u0438 \u043F\u043E \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044E \u0438 \u0432\u044B\u0441\u0442\u0440\u0430\u0438\u0432\u0430\u0435\u0442 \u043F\u043E\u0440\u044F\u0434\u043E\u043A \u0443\u0440\u043E\u043A\u043E\u0432 \u0432 \u0440\u0430\u0437\u0434\u0435\u043B\u0435 \xAB\u0421\u0435\u0433\u043E\u0434\u043D\u044F\xBB."), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 8,
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041F\u043E\u0432\u0442\u043E\u0440:"), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u041A\u0430\u0436\u0434\u0443\u044E \u043D\u0435\u0434\u0435\u043B\u044E",
    color: "#8593a3",
    active: sc.mode === "weekly",
    onClick: () => setMode("weekly")
  }), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u041D\u0435\u0447\u0451\u0442\u043D\u0430\u044F / \u0447\u0451\u0442\u043D\u0430\u044F \u043D\u0435\u0434\u0435\u043B\u044F",
    color: "#8593a3",
    active: sc.mode === "biweekly",
    onClick: () => setMode("biweekly")
  })), sc.mode === "biweekly" && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 8,
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u0420\u0435\u0434\u0430\u043A\u0442\u0438\u0440\u0443\u044E:"), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u041D\u0435\u0447\u0451\u0442\u043D\u0430\u044F \u043D\u0435\u0434\u0435\u043B\u044F",
    color: "#8593a3",
    active: wk === "A",
    onClick: () => setWk("A")
  }), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0427\u0451\u0442\u043D\u0430\u044F \u043D\u0435\u0434\u0435\u043B\u044F",
    color: "#8593a3",
    active: wk === "B",
    onClick: () => setWk("B")
  }), /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      marginLeft: 10
    }
  }, "\u0422\u0435\u043A\u0443\u0449\u0430\u044F \u043D\u0435\u0434\u0435\u043B\u044F \u2014"), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u043D\u0435\u0447\u0451\u0442\u043D\u0430\u044F",
    color: "#8593a3",
    active: curWeek === "A",
    onClick: () => setCurrent("A")
  }), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0447\u0451\u0442\u043D\u0430\u044F",
    color: "#8593a3",
    active: curWeek === "B",
    onClick: () => setCurrent("B")
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: copyWeek
  }, "\u0421\u043A\u043E\u043F\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0432 ", wk === "A" ? "чётную" : "нечётную")), /*#__PURE__*/React.createElement("div", {
    style: {
      overflowX: "auto",
      marginBottom: 6
    }
  }, /*#__PURE__*/React.createElement("table", {
    style: {
      borderCollapse: "collapse",
      fontSize: 13
    }
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
    style: {
      padding: "6px 10px",
      border: "1.5px solid #aebdcc"
    }
  }, "\u2116 \u0443\u0440\u043E\u043A\u0430"), [1, 2, 3, 4, 5, 6].map(dw => /*#__PURE__*/React.createElement("th", {
    key: dw,
    style: {
      padding: "6px 10px",
      border: "1.5px solid #aebdcc"
    }
  }, WEEKDAYS_RU[dw - 1])))), /*#__PURE__*/React.createElement("tbody", null, NUMS.map(n => /*#__PURE__*/React.createElement("tr", {
    key: n
  }, /*#__PURE__*/React.createElement("td", {
    style: {
      padding: "4px 10px",
      border: "1.5px solid #aebdcc",
      textAlign: "center",
      fontWeight: 800
    }
  }, n), [1, 2, 3, 4, 5, 6].map(dw => {
    const sl = slotAt(dw, n);
    const val = sl ? sl.kind + ":" + sl.value + ":" + (sl.classLetterId || "") : "";
    const known = !val || options.some(o => o.v === val);
    return /*#__PURE__*/React.createElement("td", {
      key: dw,
      style: {
        padding: 3,
        border: "1.5px solid #aebdcc",
        background: val ? "var(--accent-soft)" : undefined
      }
    }, /*#__PURE__*/React.createElement("select", {
      value: val,
      onChange: e => setSlot(dw, n, e.target.value),
      style: {
        ...cellSel,
        fontWeight: val ? 700 : 400,
        background: val ? "var(--accent-soft)" : "#fff"
      }
    }, /*#__PURE__*/React.createElement("option", {
      value: ""
    }, "\u2014"), !known && /*#__PURE__*/React.createElement("option", {
      value: val
    }, slotLabel(data, sl)), options.map(o => /*#__PURE__*/React.createElement("option", {
      key: o.v,
      value: o.v
    }, o.label))));
  })))))), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 14
    }
  }, "\u0423\u0440\u043E\u043A\u043E\u0432 \u0432 \u043D\u0435\u0434\u0435\u043B\u0435: ", used, ". \u0420\u0430\u0441\u043F\u0438\u0441\u0430\u043D\u0438\u0435 \u0441\u043E\u0445\u0440\u0430\u043D\u044F\u0435\u0442\u0441\u044F \u0441\u0440\u0430\u0437\u0443."), /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 800,
      marginBottom: 6
    }
  }, "\u0417\u0430\u043F\u043E\u043B\u043D\u0438\u0442\u044C \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044C \u043F\u043E \u0440\u0430\u0441\u043F\u0438\u0441\u0430\u043D\u0438\u044E"), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 8
    }
  }, "\u0423\u0440\u043E\u043A\u0438 \u043F\u043E\u043B\u0443\u0447\u0430\u044E\u0442 \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0438\u0435 \u0434\u0430\u0442\u044B \u0438 \u043D\u043E\u043C\u0435\u0440\u0430 \u0438\u0437 \u0440\u0430\u0441\u043F\u0438\u0441\u0430\u043D\u0438\u044F, \u043D\u0430\u0447\u0438\u043D\u0430\u044F \u0441 \u0441\u0435\u0433\u043E\u0434\u043D\u044F\u0448\u043D\u0435\u0433\u043E \u0434\u043D\u044F. \u041A\u0430\u043D\u0438\u043A\u0443\u043B\u044B \u043F\u0440\u043E\u043F\u0443\u0441\u043A\u0430\u044E\u0442\u0441\u044F. \u0423\u0440\u043E\u043A\u0438 \u0438\u0434\u0443\u0442 \u0432 \u043F\u043E\u0440\u044F\u0434\u043A\u0435 \u0441\u043F\u0438\u0441\u043A\u0430 \xAB\u0423\u0440\u043E\u043A\u0438\xBB. \u041F\u0440\u043E\u0432\u0435\u0434\u0451\u043D\u043D\u044B\u0435 \u0443\u0440\u043E\u043A\u0438 \u043D\u0435 \u0442\u0440\u043E\u0433\u0430\u044E\u0442\u0441\u044F."), /*#__PURE__*/React.createElement("div", {
    className: "seg",
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": fillMode === "undated" ? "true" : undefined,
    onClick: () => setFillMode("undated")
  }, "\u0422\u043E\u043B\u044C\u043A\u043E \u0443\u0440\u043E\u043A\u0438 \u0431\u0435\u0437 \u0434\u0430\u0442\u044B"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": fillMode === "all" ? "true" : undefined,
    onClick: () => setFillMode("all")
  }, "\u0412\u0441\u0435 \u0435\u0449\u0451 \u043D\u0435 \u043F\u0440\u043E\u0432\u0435\u0434\u0451\u043D\u043D\u044B\u0435 \u0443\u0440\u043E\u043A\u0438")), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: fill,
    disabled: !sc.slots.length
  }, "\u26A1 \u0417\u0430\u043F\u043E\u043B\u043D\u0438\u0442\u044C \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044C"), sc.undo && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: undoFill,
    title: "Вернуть даты, какими они были до последнего заполнения (" + sc.undo.items.length + " ур.)"
  }, "\u21B6 \u0412\u0435\u0440\u043D\u0443\u0442\u044C \u043F\u0440\u0435\u0436\u043D\u0438\u0435 \u0434\u0430\u0442\u044B")), result && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12.5,
      lineHeight: 1.55,
      marginTop: 10
    }
  }, /*#__PURE__*/React.createElement("b", null, "\u0420\u0430\u0441\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043E \u0443\u0440\u043E\u043A\u043E\u0432: ", result.placed, "."), result.groups.length > 0 && /*#__PURE__*/React.createElement("div", null, result.groups.map(g => g.label + ": " + g.placed).join("; ")), result.groups.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, fillMode === "undated" ? "Нет уроков без даты в классах из расписания (если даты уже стоят — выберите «Все ещё не проведённые уроки»)." : "Нет непроведённых уроков в классах из расписания."), result.leftover.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--neg)"
    }
  }, "\u041D\u0435 \u0445\u0432\u0430\u0442\u0438\u043B\u043E \u0441\u043B\u043E\u0442\u043E\u0432 \u043D\u0430 \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0438\u0439 \u0433\u043E\u0434: ", result.leftover.join("; "))))))));
}
function TrackChip({
  label,
  color,
  active,
  onClick
}) {
  return /*#__PURE__*/React.createElement("span", {
    className: "chip",
    "aria-pressed": active ? "true" : undefined,
    style: {
      cursor: "pointer"
    },
    onClick: onClick
  }, /*#__PURE__*/React.createElement("span", {
    className: "cdot",
    style: {
      background: color
    }
  }), label);
}
function LessonChip({
  lesson,
  data,
  onOpen,
  onToggleDone,
  draggable,
  onDragStart,
  no
}) {
  const color = lessonTrackColor(lesson, data);
  const typeTag = (lesson.lessonTypeTagIds || []).map(id => data.tags.lessonType.find(t => t.id === id)).filter(Boolean)[0];
  return /*#__PURE__*/React.createElement("div", {
    className: "lchip" + (lesson.completed ? " done" : ""),
    style: {
      "--lc": color,
      ...(lesson.archived && !lesson.completed ? {
        opacity: 0.5,
        filter: "grayscale(.7)",
        outline: "1.5px dashed #b0bcc7"
      } : null)
    },
    title: lesson.archived ? lesson.completed ? "Урок пройден (в списке «Уроки» не показывается)" : "Урок в архиве (в списке «Уроки» не показывается)" : undefined,
    draggable: draggable,
    onDragStart: onDragStart,
    onClick: () => onOpen(lesson)
  }, /*#__PURE__*/React.createElement("div", {
    className: "lt"
  }, /*#__PURE__*/React.createElement("span", {
    className: "lcheck" + (lesson.completed ? " done" : ""),
    onClick: e => {
      e.stopPropagation();
      onToggleDone(lesson);
    }
  }, lesson.completed ? "✓" : ""), /*#__PURE__*/React.createElement("span", {
    className: "tx" + (lesson.completed ? " done" : "")
  }, lesson.topic || "(без темы)", lesson.duplicateOf ? " (дубликат)" : "")), /*#__PURE__*/React.createElement("div", {
    className: "lmeta"
  }, no ? /*#__PURE__*/React.createElement("b", null, no, " \u0443\u0440. \xB7 ") : null, lessonTrackLabel(lesson, data), lesson.archived && !lesson.completed ? " · в архиве" : ""), typeTag && /*#__PURE__*/React.createElement("span", {
    style: {
      display: "inline-block",
      fontSize: 9.5,
      background: typeTag.color + "33",
      color: typeTag.color,
      borderRadius: 5,
      padding: "0 5px",
      marginTop: 2
    }
  }, typeTag.name));
}
function lessonTrackLabel(lesson, data) {
  if (lesson.courseId) {
    const c = data.courses.find(x => x.id === lesson.courseId);
    return c ? c.name : "курс";
  }
  return lesson.class ? `${lesson.class} класс` : "";
}
function CalendarWeeksView({
  month,
  lessons,
  data,
  tIso,
  onAddLesson,
  onOpenLesson,
  onToggleDone,
  dragLessonId,
  setDragLessonId,
  onDropOnDate
}) {
  const weeks = weeksOfMonthYm(month);
  const tMon = mondayOf(tIso);
  return /*#__PURE__*/React.createElement("div", null, weeks.map(mon => {
    const isThis = mon === tMon;
    const weekLessons = lessons.filter(l => mondayOf(l.plannedDate) === mon);
    return /*#__PURE__*/React.createElement("div", {
      key: mon,
      className: "weekcard" + (isThis ? " thisweek" : "")
    }, /*#__PURE__*/React.createElement("div", {
      className: "wc-head"
    }, /*#__PURE__*/React.createElement("span", {
      className: "wl"
    }, weekLabelRange(mon)), isThis && /*#__PURE__*/React.createElement("span", {
      className: "tw"
    }, "\u0442\u0435\u043A\u0443\u0449\u0430\u044F"), /*#__PURE__*/React.createElement("span", {
      style: {
        marginLeft: "auto",
        fontSize: 12,
        color: "var(--faint)"
      }
    }, weekLessons.length, " \u0443\u0440\u043E\u043A\u043E\u0432")), /*#__PURE__*/React.createElement("div", {
      className: "daygrid"
    }, [0, 1, 2, 3, 4, 5, 6].map(i => {
      const dIso = addDaysIso(mon, i);
      const vac = vacationAt(data, dIso);
      const dayInfo = lessonsOfDayByNumber(data, dIso, weekLessons.filter(l => l.plannedDate === dIso));
      const dayLessons = dayInfo.list;
      return /*#__PURE__*/React.createElement("div", {
        key: dIso,
        className: "daycol" + (dIso === tIso ? " today" : "") + (vac ? " vac" : ""),
        onDragOver: e => dragLessonId && e.preventDefault(),
        onDrop: e => {
          e.preventDefault();
          onDropOnDate(dIso);
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "dnum"
      }, WEEKDAYS_RU[i], " ", dayNum(dIso)), vac && /*#__PURE__*/React.createElement("div", {
        className: "vaclabel"
      }, vac.name), dayLessons.map(l => /*#__PURE__*/React.createElement(LessonChip, {
        key: l.id,
        lesson: l,
        data: data,
        no: dayInfo.nums.get(l.id),
        onOpen: onOpenLesson,
        onToggleDone: onToggleDone,
        draggable: true,
        onDragStart: () => setDragLessonId(l.id)
      })), /*#__PURE__*/React.createElement("div", {
        className: "addmini",
        onClick: () => onAddLesson(dIso)
      }, "+ \u0443\u0440\u043E\u043A"));
    })));
  }));
}
function CalendarMonthView({
  month,
  lessons,
  data,
  tIso,
  onAddLesson,
  onOpenLesson,
  onToggleDone,
  dragLessonId,
  setDragLessonId,
  onDropOnDate
}) {
  const cells = monthGridYm(month);
  return /*#__PURE__*/React.createElement("div", {
    className: "monthcard"
  }, /*#__PURE__*/React.createElement("div", {
    className: "wdhead"
  }, WEEKDAYS_RU.map(d => /*#__PURE__*/React.createElement("div", {
    key: d
  }, d))), [0, 1, 2, 3, 4, 5].map(r => {
    const rowCells = cells.slice(r * 7, r * 7 + 7);
    if (r === 5 && rowCells.every(c => !c.inMonth)) return null;
    return /*#__PURE__*/React.createElement("div", {
      key: r,
      className: "mgrid"
    }, rowCells.map(c => {
      const vac = vacationAt(data, c.iso);
      const dayInfo = lessonsOfDayByNumber(data, c.iso, lessons.filter(l => l.plannedDate === c.iso));
      const dayLessons = dayInfo.list;
      return /*#__PURE__*/React.createElement("div", {
        key: c.iso,
        className: "mcell" + (c.iso === tIso ? " today" : "") + (!c.inMonth ? " out" : "") + (vac && c.inMonth ? " vac" : ""),
        onDragOver: e => dragLessonId && e.preventDefault(),
        onDrop: e => {
          e.preventDefault();
          onDropOnDate(c.iso);
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "dnum"
      }, dayNum(c.iso)), dayLessons.map(l => /*#__PURE__*/React.createElement(LessonChip, {
        key: l.id,
        lesson: l,
        data: data,
        no: dayInfo.nums.get(l.id),
        onOpen: onOpenLesson,
        onToggleDone: onToggleDone,
        draggable: true,
        onDragStart: () => setDragLessonId(l.id)
      })), /*#__PURE__*/React.createElement("div", {
        className: "addmini",
        onClick: () => onAddLesson(c.iso)
      }, "+"));
    }));
  }));
}
function CreateLessonModal({
  creating,
  setCreating,
  tracks,
  data,
  onSave,
  onCancel
}) {
  const [kind, value] = (creating.trackKey || "").split(":");
  const availableExisting = creating.trackKey ? data.lessons.filter(l => !l.plannedDate && !l.archived && (kind === "course" ? l.courseId === value : l.class === value && !l.courseId)) : [];
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onCancel
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 480
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, "\u0423\u0440\u043E\u043A \u043D\u0430 ", formatDateRu(creating.date))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "seg",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": creating.mode === "new" ? "true" : undefined,
    onClick: () => setCreating({
      ...creating,
      mode: "new"
    })
  }, "\u041D\u043E\u0432\u044B\u0439 \u0443\u0440\u043E\u043A"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": creating.mode === "existing" ? "true" : undefined,
    onClick: () => setCreating({
      ...creating,
      mode: "existing"
    })
  }, "\u0412\u044B\u0431\u0440\u0430\u0442\u044C \u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u044E\u0449\u0438\u0439")), /*#__PURE__*/React.createElement(Field, {
    label: "\u041A\u043B\u0430\u0441\u0441 / \u043A\u0443\u0440\u0441"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: creating.trackKey || "",
    onChange: e => setCreating({
      ...creating,
      trackKey: e.target.value,
      existingId: ""
    })
  }, tracks.map(t => /*#__PURE__*/React.createElement("option", {
    key: trackKeyOf(t.kind, t.value),
    value: trackKeyOf(t.kind, t.value)
  }, t.label)))), creating.mode === "new" ? /*#__PURE__*/React.createElement(Field, {
    label: "\u0422\u0435\u043C\u0430 \u0443\u0440\u043E\u043A\u0430"
  }, /*#__PURE__*/React.createElement("input", {
    autoFocus: true,
    value: creating.topic,
    onChange: e => setCreating({
      ...creating,
      topic: e.target.value
    }),
    placeholder: "\u0422\u0435\u043C\u0430 \u0443\u0440\u043E\u043A\u0430",
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  })) : /*#__PURE__*/React.createElement(Field, {
    label: "\u0423\u0440\u043E\u043A \u0431\u0435\u0437 \u0434\u0430\u0442\u044B"
  }, availableExisting.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041D\u0435\u0442 \u0443\u0440\u043E\u043A\u043E\u0432 \u0431\u0435\u0437 \u0434\u0430\u0442\u044B \u0434\u043B\u044F \u044D\u0442\u043E\u0433\u043E \u043A\u043B\u0430\u0441\u0441\u0430/\u043A\u0443\u0440\u0441\u0430.") : /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: creating.existingId,
    onChange: e => setCreating({
      ...creating,
      existingId: e.target.value
    })
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u2014 \u0432\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u0443\u0440\u043E\u043A \u2014"), availableExisting.map(l => /*#__PURE__*/React.createElement("option", {
    key: l.id,
    value: l.id
  }, l.topic || "(без темы)"))))), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onCancel
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: onSave
  }, "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C"))));
}
function CalendarTab({
  data,
  update,
  onOpenLesson,
  nav,
  setNav
}) {
  const month = nav.planning.calMonth;
  const view = nav.planning.calView;
  const filterKey = nav.planning.calFilterKey;
  const calLetterFilter = nav.planning.calLetterFilter;
  const setMonth = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      calMonth: v
    }
  }));
  const setView = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      calView: v
    }
  }));
  const setFilterKey = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      calFilterKey: v
    }
  }));
  const setCalLetterFilter = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      calLetterFilter: v
    }
  }));
  const [scheduleOpen, setScheduleOpen] = useState(false);
  const [vacOpen, setVacOpen] = useState(false);
  const [addingVac, setAddingVac] = useState(false);
  const [vacName, setVacName] = useState("");
  const [vacStart, setVacStart] = useState("");
  const [vacEnd, setVacEnd] = useState("");
  const [creating, setCreating] = useState(null);
  const [dragLessonId, setDragLessonId] = useState(null);
  const tracks = allTracks(data);
  const tIso = todayIso();
  const [filterKindPart, filterValuePart] = filterKey !== "all" ? filterKey.split(":") : [null, null];
  const calFirstRender = useRef(true);
  useEffect(() => {
    if (calFirstRender.current) {
      calFirstRender.current = false;
      return;
    }
    const [fk, fv] = filterKey !== "all" ? filterKey.split(":") : [null, null];
    if (!fk || !letterValidFor(data, fk, fv, calLetterFilter)) setCalLetterFilter("all");
  }, [filterKey]);
  const addVacation = async () => {
    if (!vacStart || !vacEnd) {
      showToast("Укажите начало и конец каникул");
      return;
    }
    let s = vacStart,
      e = vacEnd;
    if (e < s) {
      const tmp = s;
      s = e;
      e = tmp;
    }
    update(d => {
      d.calendar.vacations.push({
        id: uid(),
        name: vacName.trim() || "Каникулы",
        start: s,
        end: e
      });
      return d;
    });
    setVacName("");
    setVacStart("");
    setVacEnd("");
    setAddingVac(false);
  };
  const deleteVacation = async id => {
    if (!(await Confirm("Удалить эти каникулы?"))) return;
    update(d => {
      d.calendar.vacations = d.calendar.vacations.filter(v => v.id !== id);
      return d;
    });
  };
  const plannedLessons = data.lessons.filter(l => l.plannedDate);
  const visibleLessons = plannedLessons.filter(l => filterKey === "all" || trackKeyOf(l) === filterKey).filter(l => {
    if (filterKindPart !== "class" || calLetterFilter === "all") return true;
    if (calLetterFilter === "none") return !l.classLetterId;
    return l.classLetterId === calLetterFilter;
  });
  const setLessonDate = (lessonId, newDate) => update(d => {
    const l = d.lessons.find(x => x.id === lessonId);
    if (l) {
      l.plannedDate = newDate;
      resyncAfterDateChange(d, l);
    }
    return d;
  });
  const toggleLessonDone = lesson => update(d => {
    const l = d.lessons.find(x => x.id === lesson.id);
    if (!l) return d;
    setLessonCompleted(l, !l.completed);
    return d;
  });
  const openCreate = date => {
    const defaultTrackKey = filterKey !== "all" ? filterKey : tracks[0] ? trackKeyOf(tracks[0].kind, tracks[0].value) : null;
    setCreating({
      date,
      mode: "new",
      trackKey: defaultTrackKey,
      topic: "",
      existingId: ""
    });
  };
  const saveCreate = () => {
    const c = creating;
    if (!c || !c.trackKey) {
      showToast("Выберите класс или курс");
      return;
    }
    const [kind, value] = c.trackKey.split(":");
    if (c.mode === "new") {
      if (!c.topic.trim()) {
        showToast("Впишите тему урока");
        return;
      }
      const wantLetterId = kind === "class" && calLetterFilter !== "all" && calLetterFilter !== "none" ? calLetterFilter : null;
      update(d => {
        createLessonAuto(d, {
          class: kind === "class" ? value : null,
          courseId: kind === "course" ? value : null,
          classLetterId: wantLetterId,
          topic: c.topic.trim(),
          plannedDate: c.date
        });
        return d;
      });
    } else {
      if (!c.existingId) {
        showToast("Выберите урок из списка");
        return;
      }
      update(d => {
        const l = d.lessons.find(x => x.id === c.existingId);
        if (l) {
          l.plannedDate = c.date;
          resyncAfterDateChange(d, l);
        }
        return d;
      });
    }
    setCreating(null);
  };
  const onDropOnDate = date => {
    if (dragLessonId) setLessonDate(dragLessonId, date);
    setDragLessonId(null);
  };
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-head",
    style: {
      cursor: "pointer"
    },
    onClick: () => setVacOpen(v => !v)
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      flex: 1
    }
  }, "\u041A\u0430\u043D\u0438\u043A\u0443\u043B\u044B ", /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "(", data.calendar.vacations.length, ")")), /*#__PURE__*/React.createElement(Icon, {
    name: "chev",
    size: 14
  })), vacOpen && /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, data.calendar.vacations.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, "\u041E\u0442\u043C\u0435\u0442\u044C\u0442\u0435 \u043A\u0430\u043D\u0438\u043A\u0443\u043B\u044B \u2014 \u043E\u043D\u0438 \u043F\u043E\u0434\u0441\u0432\u0435\u0442\u044F\u0442\u0441\u044F \u0432 \u043A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u0435, \u0430 \u0433\u043E\u0434 \u0440\u0430\u0437\u043E\u0431\u044C\u0451\u0442\u0441\u044F \u043D\u0430 \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u0438 \u0434\u043B\u044F \u0436\u0443\u0440\u043D\u0430\u043B\u0430."), sortedVacations(data).map(v => /*#__PURE__*/React.createElement("div", {
    key: v.id,
    className: "row",
    style: {
      padding: "6px 0",
      borderBottom: "1px solid var(--line)"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      fontWeight: 700
    }
  }, v.name), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, formatDateRu(v.start), " \u2013 ", formatDateRu(v.end)), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => deleteVacation(v.id)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 12
  })))), !addingVac ? /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setAddingVac(true),
    style: {
      marginTop: 8
    }
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u043A\u0430\u043D\u0438\u043A\u0443\u043B\u044B") : /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginTop: 8
    }
  }, /*#__PURE__*/React.createElement("input", {
    placeholder: "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 (\u041E\u0441\u0435\u043D\u043D\u0438\u0435)",
    value: vacName,
    onChange: e => setVacName(e.target.value),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      maxWidth: 160
    }
  }), /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: vacStart,
    onChange: e => setVacStart(e.target.value),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px"
    }
  }), /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: vacEnd,
    onChange: e => setVacEnd(e.target.value),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px"
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: addVacation
  }, "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setAddingVac(false)
  }, "\u041E\u0442\u043C\u0435\u043D\u0430")))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "viewtoggle"
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": view === "weeks" ? "true" : undefined,
    onClick: () => setView("weeks")
  }, "\u041D\u0435\u0434\u0435\u043B\u0438"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": view === "month" ? "true" : undefined,
    onClick: () => setView("month")
  }, "\u041C\u0435\u0441\u044F\u0446")), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setMonth(addMonthsYm(month, -1))
  }, "\u2190"), /*#__PURE__*/React.createElement("span", {
    style: {
      fontFamily: "var(--font-display)",
      fontWeight: 700,
      minWidth: 130,
      textAlign: "center",
      textTransform: "capitalize"
    }
  }, ymLabel(month)), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setMonth(addMonthsYm(month, 1))
  }, "\u2192"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setMonth(curMonthYm())
  }, "\u0421\u0435\u0433\u043E\u0434\u043D\u044F"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setScheduleOpen(true),
    title: "\u0412 \u043A\u0430\u043A\u0438\u0435 \u0434\u043D\u0438 \u0438 \u043F\u043E\u0434 \u043A\u0430\u043A\u0438\u043C \u043D\u043E\u043C\u0435\u0440\u043E\u043C \u0443 \u0432\u0430\u0441 \u0443\u0440\u043E\u043A\u0438 \u0432 \u043A\u0430\u043A\u043E\u043C \u043A\u043B\u0430\u0441\u0441\u0435"
  }, "\uD83D\uDDD3 \u041C\u043E\u0451 \u0440\u0430\u0441\u043F\u0438\u0441\u0430\u043D\u0438\u0435"), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0412\u0441\u0435",
    color: "#8593a3",
    active: filterKey === "all",
    onClick: () => setFilterKey("all")
  }), tracks.map(t => /*#__PURE__*/React.createElement(TrackChip, {
    key: trackKeyOf(t.kind, t.value),
    label: t.label,
    color: t.color,
    active: filterKey === trackKeyOf(t.kind, t.value),
    onClick: () => setFilterKey(trackKeyOf(t.kind, t.value))
  })))), filterKindPart === "class" && lettersForGrade(data, filterValuePart).length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 14,
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041B\u0438\u0442\u0435\u0440\u0430:"), lettersForGrade(data, filterValuePart).map(cl => /*#__PURE__*/React.createElement(TrackChip, {
    key: cl.id,
    label: `${filterValuePart}${cl.name}`,
    color: "#8593a3",
    active: calLetterFilter === cl.id,
    onClick: () => setCalLetterFilter(cl.id)
  })), (calLetterFilter === "none" || hasNoLetterItems(data, filterValuePart)) && /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0411\u0435\u0437 \u043B\u0438\u0442\u0435\u0440\u044B",
    color: "#8593a3",
    active: calLetterFilter === "none",
    onClick: () => setCalLetterFilter("none")
  }), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0412\u0441\u0435",
    color: "#8593a3",
    active: calLetterFilter === "all",
    onClick: () => setCalLetterFilter("all")
  })), tracks.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u0421\u043D\u0430\u0447\u0430\u043B\u0430 \u0441\u043E\u0437\u0434\u0430\u0439\u0442\u0435 \u043A\u043B\u0430\u0441\u0441 \u0438\u043B\u0438 \u043A\u0443\u0440\u0441 \u0432 \u0440\u0430\u0437\u0434\u0435\u043B\u0435 \xAB\u0423\u0440\u043E\u043A\u0438\xBB \u2014 \u0442\u043E\u0433\u0434\u0430 \u0437\u0434\u0435\u0441\u044C \u043C\u043E\u0436\u043D\u043E \u0431\u0443\u0434\u0435\u0442 \u043F\u043B\u0430\u043D\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0443\u0440\u043E\u043A\u0438.") : view === "weeks" ? /*#__PURE__*/React.createElement(CalendarWeeksView, {
    month: month,
    lessons: visibleLessons,
    data: data,
    tIso: tIso,
    onAddLesson: openCreate,
    onOpenLesson: onOpenLesson,
    onToggleDone: toggleLessonDone,
    dragLessonId: dragLessonId,
    setDragLessonId: setDragLessonId,
    onDropOnDate: onDropOnDate
  }) : /*#__PURE__*/React.createElement(CalendarMonthView, {
    month: month,
    lessons: visibleLessons,
    data: data,
    tIso: tIso,
    onAddLesson: openCreate,
    onOpenLesson: onOpenLesson,
    onToggleDone: toggleLessonDone,
    dragLessonId: dragLessonId,
    setDragLessonId: setDragLessonId,
    onDropOnDate: onDropOnDate
  }), scheduleOpen && /*#__PURE__*/React.createElement(ScheduleModal, {
    data: data,
    update: update,
    onClose: () => setScheduleOpen(false)
  }), creating && /*#__PURE__*/React.createElement(CreateLessonModal, {
    creating: creating,
    setCreating: setCreating,
    tracks: tracks,
    data: data,
    onSave: saveCreate,
    onCancel: () => setCreating(null)
  }));
}
function MarkCellModal({
  student,
  date,
  lesson,
  existingMark,
  scale,
  onSave,
  onClear,
  onClose,
  onNavigate
}) {
  const min = scale === "100" ? 0 : 1;
  const max = scale === "100" ? 100 : 5;
  const [isAbsent, setIsAbsent] = useState(existingMark?.value === "Н");
  const [val, setVal] = useState(existingMark && existingMark.value !== "Н" ? String(existingMark.value) : "");
  const [excused, setExcused] = useState(existingMark?.excused !== false);
  const [reason, setReason] = useState(existingMark?.reason || "");
  const [posted, setPosted] = useState(!!existingMark?.posted);
  const buildPatch = () => {
    if (isAbsent) return {
      value: "Н",
      excused,
      reason,
      scale,
      posted
    };
    const n = +val;
    if (!val || Number.isNaN(n) || n < min || n > max) return null;
    return {
      value: n,
      excused: null,
      reason,
      scale,
      posted
    };
  };
  const save = () => {
    const patch = buildPatch();
    if (!patch) {
      showToast(`Введите оценку от ${min} до ${max}`);
      return;
    }
    onSave(patch);
  };
  const handleKeyDown = e => {
    const map = {
      ArrowUp: "up",
      ArrowDown: "down",
      ArrowLeft: "left",
      ArrowRight: "right"
    };
    const dir = map[e.key];
    if (!dir || e.target.tagName === "INPUT" && (dir === "left" || dir === "right")) return;
    e.preventDefault();
    onNavigate(dir, buildPatch());
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 360
    },
    onClick: e => e.stopPropagation(),
    onKeyDown: handleKeyDown
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, student.name), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, formatDateRu(date), lesson ? ` · ${lesson.topic || "(без темы)"}` : ""), /*#__PURE__*/React.createElement("div", {
    className: "seg",
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": !isAbsent ? "true" : undefined,
    onClick: () => setIsAbsent(false)
  }, "\u041E\u0442\u043C\u0435\u0442\u043A\u0430"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": isAbsent ? "true" : undefined,
    onClick: () => setIsAbsent(true)
  }, "\u041D \u2014 \u043E\u0442\u0441\u0443\u0442\u0441\u0442\u0432\u043E\u0432\u0430\u043B")), !isAbsent ? /*#__PURE__*/React.createElement(Field, {
    label: `Оценка (${min}–${max})`
  }, /*#__PURE__*/React.createElement("input", {
    autoFocus: true,
    type: "number",
    min: min,
    max: max,
    value: val,
    onChange: e => setVal(e.target.value),
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  })) : /*#__PURE__*/React.createElement("div", {
    className: "seg",
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": excused ? "true" : undefined,
    onClick: () => setExcused(true)
  }, "\u0423\u0432\u0430\u0436\u0438\u0442\u0435\u043B\u044C\u043D\u0430\u044F"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": !excused ? "true" : undefined,
    onClick: () => setExcused(false)
  }, "\u041D\u0435\u0443\u0432\u0430\u0436\u0438\u0442\u0435\u043B\u044C\u043D\u0430\u044F")), /*#__PURE__*/React.createElement(Field, {
    label: "\u041A\u043E\u043C\u043C\u0435\u043D\u0442\u0430\u0440\u0438\u0439"
  }, /*#__PURE__*/React.createElement("input", {
    value: reason,
    onChange: e => setReason(e.target.value),
    placeholder: "\u0434\u0438\u043A\u0442\u0430\u043D\u0442, \u0443\u0441\u0442\u043D\u044B\u0439 \u043E\u0442\u0432\u0435\u0442\u2026",
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  })), /*#__PURE__*/React.createElement(ToggleSwitch, {
    checked: posted,
    onChange: setPosted,
    label: "\u0412\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u0430 \u0432 \u0448\u043A\u043E\u043B\u044C\u043D\u044B\u0439 \u044D\u043B. \u0436\u0443\u0440\u043D\u0430\u043B"
  })), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, existingMark && /*#__PURE__*/React.createElement(Btn, {
    kind: "danger",
    onClick: onClear
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  }), " \u0423\u0434\u0430\u043B\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0443"), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: save
  }, "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C"))));
}

// Отчёт за день по журналу: сколько детей в классе, сколько присутствовало (все, кроме отмеченных
// «Н» в этот день), сколько каких отметок, таблица ученик/отметка и проценты успеваемости/качества.
// Считается только по 5-балльной шкале — у 100-балльной нет смысловых порогов «3/4/5».
function journalDateReportHtml(data, students, date, groupLabel) {
  const marksOf = s => studentMarksOf(data, s.id, "5").filter(m => m.date === date);
  const absent = students.filter(s => {
    const ms = marksOf(s);
    return ms.length > 0 && ms.every(m => m.value === "Н");
  }).length;
  const counts = {
    5: 0,
    4: 0,
    3: 0,
    2: 0,
    1: 0
  };
  const grades = [];
  let rows = "";
  students.forEach(s => {
    const ms = marksOf(s);
    rows += `<tr><td style="padding:4px 8px;">${escapeHtml(s.name)}</td><td style="text-align:center;padding:4px 8px;">${ms.length ? ms.map(m => m.value).join(" ") : "—"}</td></tr>`;
    const nums = ms.filter(m => m.value !== "Н");
    if (!nums.length && ms.length) return;
    nums.forEach(m => {
      if (counts[+m.value] !== undefined) counts[+m.value]++;
    });
    grades.push(nums.length ? Math.round(nums.reduce((a, m) => a + +m.value, 0) / nums.length) : null);
  });
  const perf = performancePct(grades),
    qual = qualityPct(grades);
  return `<h2>${escapeHtml(groupLabel)} — отчёт за ${formatDateRu(date)}</h2>` + `<div>Количество детей в классе: ${students.length}</div>` + `<div>Присутствовало на уроке: ${students.length - absent}</div>` + `<div style="margin-bottom:14px;">Отметок «5»: ${counts[5]}, «4»: ${counts[4]}, «3»: ${counts[3]}, «2»: ${counts[2]}, «1»: ${counts[1]}</div>` + `<table style="border-collapse:collapse;width:100%;margin-bottom:14px;"><tr><th style="text-align:left;border-bottom:1px solid #ccc;padding:4px 8px;">Ученик</th><th style="border-bottom:1px solid #ccc;padding:4px 8px;">Отметка</th></tr>${rows}</table>` + `<div>Успеваемость: ${perf == null ? "—" : perf + "%"}</div><div>Качество: ${qual == null ? "—" : qual + "%"}</div>`;
}
function JournalDateReportModal({
  data,
  students,
  date,
  groupLabel,
  onClose
}) {
  const title = `Отчёт за ${formatDateRu(date)}`;
  const bodyHtml = journalDateReportHtml(data, students, date, groupLabel);
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, title), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => downloadWord(safeFileName(`otchet-${date}`) + ".doc", title, bodyHtml)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 12
  }), " Word"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => printAsPdf(`<div style="font-family:Arial,sans-serif;color:#000;padding:24px;">${bodyHtml}</div>`)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "note",
    size: 12
  }), " PDF")), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13.5,
      lineHeight: 1.6
    },
    dangerouslySetInnerHTML: {
      __html: bodyHtml
    }
  }))));
}
function BulkMarkModal({
  students,
  date,
  data,
  gridScale,
  onSave,
  onClose
}) {
  const min = gridScale === "100" ? 0 : 1;
  const max = gridScale === "100" ? 100 : 5;
  const [entries, setEntries] = useState(() => students.map(s => {
    const existing = studentMarksOf(data, s.id, gridScale).find(m => m.date === date);
    return {
      studentId: s.id,
      name: s.name,
      value: existing && existing.value !== "Н" ? String(existing.value) : "",
      isAbsent: existing?.value === "Н",
      excused: existing?.excused !== false
    };
  }));
  const patchEntry = (id, patch) => setEntries(prev => prev.map(e => e.studentId === id ? {
    ...e,
    ...patch
  } : e));
  const save = () => {
    const results = entries.filter(e => e.isAbsent || e.value !== "").map(e => ({
      studentId: e.studentId,
      value: e.isAbsent ? "Н" : +e.value,
      excused: e.isAbsent ? e.excused !== false : null
    }));
    onSave(results);
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u041E\u0442\u043C\u0435\u0442\u043A\u0438 \u043D\u0430 ", formatDateRu(date)), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      maxHeight: 420,
      overflow: "auto"
    }
  }, entries.map(e => /*#__PURE__*/React.createElement("div", {
    key: e.studentId,
    className: "row",
    style: {
      padding: "6px 0",
      borderBottom: "1px solid var(--line)"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      fontSize: 13.5,
      fontWeight: 600
    }
  }, e.name), !e.isAbsent ? /*#__PURE__*/React.createElement("input", {
    type: "number",
    min: min,
    max: max,
    value: e.value,
    onChange: ev => patchEntry(e.studentId, {
      value: ev.target.value
    }),
    placeholder: "\u2014",
    style: {
      width: 60,
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "5px 7px"
    }
  }) : /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: e.excused ? "y" : "n",
    onChange: ev => patchEntry(e.studentId, {
      excused: ev.target.value === "y"
    })
  }, /*#__PURE__*/React.createElement("option", {
    value: "y"
  }, "\u0423\u0432\u0430\u0436."), /*#__PURE__*/React.createElement("option", {
    value: "n"
  }, "\u041D\u0435\u0443\u0432\u0430\u0436.")), /*#__PURE__*/React.createElement(Btn, {
    kind: e.isAbsent ? "gold" : "ghost",
    sm: true,
    onClick: () => patchEntry(e.studentId, {
      isAbsent: !e.isAbsent
    })
  }, "\u041D")))), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: save
  }, "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u0432\u0441\u0435"))));
}
function AddJournalMarkModal({
  students,
  gridScale,
  onSave,
  onClose
}) {
  const min = gridScale === "100" ? 0 : 1;
  const max = gridScale === "100" ? 100 : 5;
  const [studentId, setStudentId] = useState(students[0]?.id || "");
  const [date, setDate] = useState(todayIso());
  const [isAbsent, setIsAbsent] = useState(false);
  const [val, setVal] = useState("");
  const [excused, setExcused] = useState(true);
  const [reason, setReason] = useState("");
  const save = () => {
    if (!studentId) {
      showToast("Выберите ученика");
      return;
    }
    if (isAbsent) {
      onSave({
        studentId,
        date,
        value: "Н",
        excused,
        reason,
        scale: gridScale
      });
      return;
    }
    const n = +val;
    if (!val || Number.isNaN(n) || n < min || n > max) {
      showToast(`Введите оценку от ${min} до ${max}`);
      return;
    }
    onSave({
      studentId,
      date,
      value: n,
      excused: null,
      reason,
      scale: gridScale
    });
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 360
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u041E\u0442\u043C\u0435\u0442\u043A\u0430 \u0431\u0435\u0437 \u0443\u0440\u043E\u043A\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u0423\u0447\u0435\u043D\u0438\u043A"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: studentId,
    onChange: e => setStudentId(e.target.value)
  }, students.map(s => /*#__PURE__*/React.createElement("option", {
    key: s.id,
    value: s.id
  }, s.name)))), /*#__PURE__*/React.createElement(Field, {
    label: "\u0414\u0430\u0442\u0430"
  }, /*#__PURE__*/React.createElement("input", {
    type: "date",
    value: date,
    onChange: e => setDate(e.target.value),
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  })), /*#__PURE__*/React.createElement("div", {
    className: "seg",
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": !isAbsent ? "true" : undefined,
    onClick: () => setIsAbsent(false)
  }, "\u041E\u0442\u043C\u0435\u0442\u043A\u0430"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": isAbsent ? "true" : undefined,
    onClick: () => setIsAbsent(true)
  }, "\u041D \u2014 \u043E\u0442\u0441\u0443\u0442\u0441\u0442\u0432\u043E\u0432\u0430\u043B")), !isAbsent ? /*#__PURE__*/React.createElement(Field, {
    label: `Оценка (${min}–${max})`
  }, /*#__PURE__*/React.createElement("input", {
    type: "number",
    min: min,
    max: max,
    value: val,
    onChange: e => setVal(e.target.value),
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  })) : /*#__PURE__*/React.createElement("div", {
    className: "seg",
    style: {
      marginBottom: 10
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": excused ? "true" : undefined,
    onClick: () => setExcused(true)
  }, "\u0423\u0432\u0430\u0436\u0438\u0442\u0435\u043B\u044C\u043D\u0430\u044F"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": !excused ? "true" : undefined,
    onClick: () => setExcused(false)
  }, "\u041D\u0435\u0443\u0432\u0430\u0436\u0438\u0442\u0435\u043B\u044C\u043D\u0430\u044F")), /*#__PURE__*/React.createElement(Field, {
    label: "\u041A\u043E\u043C\u043C\u0435\u043D\u0442\u0430\u0440\u0438\u0439"
  }, /*#__PURE__*/React.createElement("input", {
    value: reason,
    onChange: e => setReason(e.target.value),
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: save
  }, "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C"))));
}

// ---- Импорт отметок из таблицы школьного электронного журнала (.xlsx) ----
// Лист: в строке 1 заголовок («6Б: Литература (Минасян А. А.) - I четверть»), в строке-шапке даты («1сен»),
// ниже строки учеников: № | Ф.И. | отметки по датам (5, 4, …, «Н»). Каждый столбец-дата = один урок.
async function xlsxToSheets(buf) {
  const dec = b => new TextDecoder("utf-8").decode(b);
  const xml = t => new DOMParser().parseFromString(t, "application/xml");
  let wb, rels;
  try {
    wb = xml(dec(await readZipEntryBytes(buf, "xl/workbook.xml")));
    rels = xml(dec(await readZipEntryBytes(buf, "xl/_rels/workbook.xml.rels")));
  } catch (e) {
    throw new Error("это не файл Excel (.xlsx)");
  }
  let ss = [];
  try {
    const sx = xml(dec(await readZipEntryBytes(buf, "xl/sharedStrings.xml")));
    ss = xmlAll(sx, "si").map(si => xmlAll(si, "t").filter(t => !t.parentNode || t.parentNode.localName !== "rPh").map(t => t.textContent).join(""));
  } catch (e) {/* без общих строк */}
  const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const out = [];
  for (const sh of xmlAll(wb, "sheet")) {
    const rid = sh.getAttributeNS(NS_R, "id") || sh.getAttribute("r:id");
    const rel = xmlAll(rels, "Relationship").find(r => r.getAttribute("Id") === rid);
    if (!rel) continue;
    let target = (rel.getAttribute("Target") || "").replace(/^\//, "");
    if (!/^xl\//.test(target)) target = "xl/" + target;
    let sdoc;
    try {
      sdoc = xml(dec(await readZipEntryBytes(buf, target)));
    } catch (e) {
      continue;
    }
    const rows = {};
    xmlAll(sdoc, "c").forEach(c => {
      const m = (c.getAttribute("r") || "").match(/^([A-Z]+)(\d+)$/);
      if (!m) return;
      const t = c.getAttribute("t");
      const v = Array.from(c.children).find(k => k.localName === "v");
      const is = Array.from(c.children).find(k => k.localName === "is");
      let val = null;
      if (t === "s" && v) val = ss[+v.textContent];else if (t === "inlineStr" && is) val = xmlAll(is, "t").map(x => x.textContent).join("");else if (v) val = t === "str" || t === "e" || t === "b" ? v.textContent : isNaN(+v.textContent) ? v.textContent : +v.textContent;
      if (val === null || val === undefined || val === "") return;
      (rows[+m[2]] = rows[+m[2]] || {})[m[1]] = val;
    });
    out.push({
      name: sh.getAttribute("name") || "",
      rows
    });
  }
  return out;
}
const JR_MONTHS = {
  янв: 1,
  фев: 2,
  мар: 3,
  апр: 4,
  май: 5,
  мая: 5,
  июн: 6,
  июл: 7,
  авг: 8,
  сен: 9,
  окт: 10,
  ноя: 11,
  дек: 12
};
function jrParseDate(text, startYear) {
  const s = String(text).trim().toLowerCase();
  let day = null,
    mon = null;
  let m = s.match(/^(\d{1,2})\s*([а-яё]{3,})\.?$/);
  if (m) {
    day = +m[1];
    mon = JR_MONTHS[m[2].slice(0, 3)] || null;
  } else if (m = s.match(/^(\d{1,2})[./](\d{1,2})(?:[./]\d{2,4})?$/)) {
    day = +m[1];
    mon = +m[2];
  }
  if (!day || !mon || day > 31 || mon > 12) return null;
  const year = mon >= 9 ? startYear : startYear + 1;
  return `${year}-${String(mon).padStart(2, "0")}-${String(day).padStart(2, "0")}`;
}
function jrNameKey(n) {
  return String(n || "").toLowerCase().replace(/ё/g, "е").replace(/[^a-zа-я0-9 ]+/g, " ").split(/\s+/).filter(Boolean).sort().join(" ");
}
// Разбирает листы одного файла в удобную структуру: группа, даты-столбцы, ученики и их отметки.
function parseJournalSheets(fileName, sheets) {
  const fy = fileName.match(/(\d{4})\s*[-–]\s*(\d{4})/);
  const now = new Date();
  const startYear = fy ? +fy[1] : now.getMonth() + 1 >= 9 ? now.getFullYear() : now.getFullYear() - 1;
  const fnClass = fileName.replace(/^.*?\d{4}\s*[-–]\s*\d{4}_/, "").match(/^(\d{1,2})\s*([А-Яа-яA-Za-z])?(?=[_\s.])/);
  const res = [];
  sheets.forEach(sh => {
    const rowNums = Object.keys(sh.rows).map(Number).sort((a, b) => a - b);
    if (!rowNums.length) return;
    const title = String(Object.values(sh.rows[rowNums[0]])[0] || "");
    const hdrRow = rowNums.find(r => Object.values(sh.rows[r]).filter(v => typeof v === "string" && jrParseDate(v, startYear)).length >= 2);
    if (!hdrRow) return;
    const columns = [];
    Object.entries(sh.rows[hdrRow]).forEach(([col, v]) => {
      const dt = typeof v === "string" ? jrParseDate(v, startYear) : null;
      if (dt) columns.push({
        col,
        date: dt,
        label: v
      });
    });
    const students = [];
    rowNums.filter(r => r > hdrRow).forEach(r => {
      const row = sh.rows[r];
      if (typeof row.A !== "number" || typeof row.B !== "string") return;
      const cells = {};
      columns.forEach(c => {
        const raw = row[c.col];
        if (raw === undefined) return;
        const toks = String(raw).replace(/\s+/g, " ").trim().split(" ").filter(Boolean);
        const nums = toks.filter(t => /^[1-5]$/.test(t));
        const abs = toks.some(t => /^[НHнh]$/.test(t));
        const bad = toks.filter(t => !/^[1-5]$/.test(t) && !/^[НHнh]$/.test(t));
        if (nums.length || abs) cells[c.col] = {
          value: nums.length ? +nums[nums.length - 1] : "Н",
          both: nums.length > 0 && abs,
          bad
        };else if (bad.length) cells[c.col] = {
          value: null,
          bad
        };
      });
      students.push({
        name: row.B.trim().replace(/\s+/g, " "),
        cells
      });
    });
    const tm = title.match(/^\s*(\d{1,2})\s*([А-Яа-яA-Za-z])?\s*:/);
    const grade = tm ? tm[1] : fnClass ? fnClass[1] : "";
    const letterName = tm && tm[2] || fnClass && fnClass[2] || "";
    const teacher = ((title.match(/\(([^)]+)\)/) || [])[1] || "").trim();
    const groupTag = ((title.match(/\/\s*([^(]+?)\s*\(/) || [])[1] || "").trim();
    res.push({
      sheetName: sh.name,
      title,
      grade,
      letterName,
      teacher,
      groupTag,
      columns,
      students
    });
  });
  return res;
}
// ======================= Импорт из эл. журнала v2: сверка по КОЛИЧЕСТВУ оценок =======================
const dayDiff = (a, b) => Math.abs((new Date(a) - new Date(b)) / 86400000);

// Разбор отметок одного ученика: сопоставляет отметки таблицы и журнала по значению (сначала точное совпадение
// даты, затем ближайшие даты) и возвращает, что надо добавить (остаток таблицы — «не хватает в журнале») и что
// есть только в журнале (остаток журнала — «ещё не выставлено в электронный журнал»).
function planStudentMarks(d, student, st, sheet) {
  const dates = sheet.columns.map(c => c.date).sort();
  const minD = dates[0],
    maxD = dates[dates.length - 1];
  const excel = [];
  sheet.columns.forEach(c => {
    const cell = st.cells[c.col];
    if (cell && cell.value !== null) excel.push({
      date: c.date,
      value: cell.value,
      label: c.label
    });
  });
  // отметки журнала позже последней даты файла считаются «ещё не выставленными» (файл мог быть выгружен несколько дней назад)
  const qIdx = quarterOfDate(d, minD);
  const journal = d.calendar.marks.filter(m => m.studentId === student.id && (m.scale || "5") === "5" && m.date >= minD && quarterOfDate(d, m.date) === qIdx).map(m => ({
    id: m.id,
    date: m.date,
    value: m.value === "Н" ? "Н" : +m.value,
    mark: m
  }));
  const toAdd = [];
  const surplus = journal.filter(j => j.date > maxD);
  const journalIn = journal.filter(j => j.date <= maxD);
  const fileCounts = {
    1: 0,
    2: 0,
    3: 0,
    4: 0,
    5: 0,
    "Н": 0
  };
  const matched = [];
  excel.forEach(e => {
    fileCounts[e.value]++;
  });
  JR_VALUES.forEach(v => {
    let X = excel.filter(e => e.value === v);
    let Y = journalIn.filter(j => j.value === v);
    // 1) точные совпадения даты
    X = X.filter(e => {
      const i = Y.findIndex(j => j.date === e.date);
      if (i === -1) return true;
      matched.push(Y[i].id);
      Y.splice(i, 1);
      return false;
    });
    // 2) оставшиеся — парами по ближайшим датам (даты не совпали, но отметка в журнале уже есть)
    while (X.length && Y.length) {
      let best = null;
      X.forEach((e, xi) => Y.forEach((j, yi) => {
        const diff = dayDiff(e.date, j.date);
        if (!best || diff < best.diff) best = {
          xi,
          yi,
          diff
        };
      }));
      X.splice(best.xi, 1);
      matched.push(Y[best.yi].id);
      Y.splice(best.yi, 1);
    }
    toAdd.push(...X);
    surplus.push(...Y);
  });
  return {
    toAdd,
    surplus,
    fileCounts,
    matched
  };
}
// План импорта (без изменения данных): по каждому выбранному листу — найденные ученики, что добавить в журнал,
// чего в файле нет (значит, не выставлено в электронный журнал) и счётчики отметок файла для сохранения.
function planJournalImport(d, items) {
  return items.map(it => {
    const sh = it.sheet,
      grade = it.grade,
      letterId = it.letterId || null;
    const dates = sh.columns.map(c => c.date).sort();
    const rep = {
      label: `${grade}${it.letterLabel || ""}${sh.groupTag ? " · " + sh.groupTag : ""}`,
      grade,
      letterId,
      fileName: it.fileName || "",
      minDate: dates[0],
      maxDate: dates[dates.length - 1],
      sheet: sh,
      studentsTotal: sh.students.length,
      matched: 0,
      unmatched: [],
      people: [],
      fileCounts: [],
      matchedIds: [],
      bad: [],
      doubles: 0
    };
    const pool = d.calendar.students.filter(s => s.class === grade && !s.courseId);
    const pref = pool.filter(s => letterId ? s.classLetterId === letterId : !s.classLetterId);
    const fallback = letterId ? pool.filter(s => !s.classLetterId) : [];
    sh.students.forEach(st => {
      const key = jrNameKey(st.name);
      const s = pref.find(x => jrNameKey(x.name) === key) || fallback.find(x => jrNameKey(x.name) === key) || null;
      Object.entries(st.cells).forEach(([col, cell]) => {
        const c = sh.columns.find(x => x.col === col);
        if (cell.bad && cell.bad.length) rep.bad.push(`${st.name}, ${c ? c.label : col}: «${cell.bad.join(" ")}»`);
        if (cell.both) rep.doubles++;
      });
      if (!s) {
        rep.unmatched.push(st.name);
        return;
      }
      rep.matched++;
      const p = planStudentMarks(d, s, st, sh);
      rep.fileCounts.push({
        studentId: s.id,
        counts: p.fileCounts
      });
      rep.matchedIds.push(...p.matched);
      if (p.toAdd.length || p.surplus.length) rep.people.push({
        student: s,
        toAdd: p.toAdd,
        surplus: p.surplus
      });
    });
    rep.addCount = rep.people.reduce((n, p) => n + p.toAdd.filter(x => x.value !== "Н").length, 0);
    rep.absCount = rep.people.reduce((n, p) => n + p.toAdd.filter(x => x.value === "Н").length, 0);
    rep.surplusCount = rep.people.reduce((n, p) => n + p.surplus.length, 0);
    rep.skipped = rep.matched === 0;
    return rep;
  });
}
// Записывает план в данные. mode "lessons" — недостающие отметки ставятся на даты из файла, недостающие уроки
// создаются; mode "marks" — только отметки: каждая встаёт в ближайшую по дате к файловой ПУСТУЮ клетку ученика
// в уже существующем столбце-дате до сегодняшнего дня (уроки не создаются). Запоминает счётчики файла (journalFileSnapshots).
function applyJournalPlan(d, plans, mode) {
  const result = {
    added: 0,
    absAdded: 0,
    lessonsCreated: 0,
    blocked: [],
    moved: 0
  };
  const today = todayIso();
  d.calendar.journalFileSnapshots = d.calendar.journalFileSnapshots || [];
  plans.forEach(rep => {
    if (rep.skipped) return;
    const group = {
      class: rep.grade,
      courseId: null,
      classLetterId: rep.letterId
    };
    const lessonByDate = {};
    const groupLessons = d.lessons.filter(l => sameLessonGroup(l, group) && !l.archived);
    const studentIds = new Set(d.calendar.students.filter(s => s.class === rep.grade && !s.courseId && (s.classLetterId || null) === (rep.letterId || null)).map(s => s.id));
    const columnDates = Array.from(new Set([...groupLessons.map(l => l.actualDate || l.plannedDate).filter(Boolean), ...d.calendar.marks.filter(m => studentIds.has(m.studentId) && (m.scale || "5") === "5").map(m => m.date)]));
    rep.people.forEach(p => {
      p.toAdd.forEach(e => {
        let date = e.date;
        if (mode === "marks") {
          const qIdx = quarterOfDate(d, e.date);
          const free = columnDates.filter(dt => dt < today && quarterOfDate(d, dt) === qIdx && !d.calendar.marks.some(m => m.studentId === p.student.id && m.date === dt && (m.scale || "5") === "5")).sort((a, b) => dayDiff(a, e.date) - dayDiff(b, e.date) || (a < b ? -1 : 1));
          if (!free.length) {
            result.blocked.push(`${p.student.name}: ${e.value} (по таблице — ${formatDateRu(e.date)}) — нет свободной клетки до сегодняшнего дня`);
            return;
          }
          date = free[0];
          if (date !== e.date) result.moved++;
        }
        let lesson = lessonByDate[date] || groupLessons.find(l => (l.actualDate || l.plannedDate) === date) || null;
        if (!lesson && mode !== "marks") {
          lesson = createLessonAuto(d, {
            ...group,
            plannedDate: date,
            needsWork: true
          });
          result.lessonsCreated++;
        }
        if (lesson) lessonByDate[date] = lesson;
        d.calendar.marks.push({
          id: uid(),
          studentId: p.student.id,
          value: e.value,
          date,
          reason: "",
          scale: "5",
          excused: e.value === "Н" ? true : null,
          lessonId: lesson ? lesson.id : null,
          exerciseId: null,
          questionId: null,
          posted: true
        });
        if (e.value === "Н") result.absAdded++;else result.added++;
      });
    });
    // отметки, уже имеющиеся в таблице школьного журнала, считаются выставленными
    const postedSet = new Set(rep.matchedIds || []);
    d.calendar.marks.forEach(m => {
      if (postedSet.has(m.id)) m.posted = true;
    });
    // запоминаем таблицу: одна запись на группу и четверть, новая заменяет старую
    const qIdx = quarterOfDate(d, rep.minDate);
    d.calendar.journalFileSnapshots = d.calendar.journalFileSnapshots.filter(sn => !(sn.grade === rep.grade && (sn.letterId || null) === (rep.letterId || null) && quarterOfDate(d, sn.minDate) === qIdx));
    const students = {};
    rep.fileCounts.forEach(fc => {
      students[fc.studentId] = fc.counts;
    });
    d.calendar.journalFileSnapshots.push({
      id: uid(),
      grade: rep.grade,
      letterId: rep.letterId || null,
      minDate: rep.minDate,
      maxDate: rep.maxDate,
      importedAt: new Date().toISOString(),
      fileName: rep.fileName,
      students
    });
  });
  return result;
}
function JournalXlsxImportModal({
  data,
  update,
  onClose
}) {
  const [files, setFiles] = useState([]);
  const [busy, setBusy] = useState(false);
  const [step, setStep] = useState("pick"); // pick | mode | done
  const [result, setResult] = useState(null);
  const [err, setErr] = useState("");
  const inputRef = useRef(null);
  const sheetsFlat = files.flatMap((f, fi) => f.sheets.map((s, si) => ({
    fi,
    si,
    s
  })));
  const patchSheet = (fi, si, patch) => setFiles(prev => prev.map((f, i) => i !== fi ? f : {
    ...f,
    sheets: f.sheets.map((s, j) => j !== si ? s : {
      ...s,
      ...patch
    })
  }));
  const onPick = async e => {
    const list = Array.from(e.target.files || []);
    e.target.value = "";
    if (!list.length) return;
    setBusy(true);
    setErr("");
    const added = [];
    for (const f of list) {
      try {
        const parsed = parseJournalSheets(f.name, await xlsxToSheets(await f.arrayBuffer()));
        if (!parsed.length) {
          setErr("В файле «" + f.name + "» не найдена таблица с датами и отметками");
          continue;
        }
        added.push({
          name: f.name,
          sheets: parsed
        });
      } catch (ex) {
        setErr("Не удалось прочитать «" + f.name + "»: " + (ex && ex.message || ex));
      }
    }
    setFiles(prev => {
      const all = [...prev, ...added];
      const freq = {};
      all.forEach(f => f.sheets.forEach(s => {
        freq[s.teacher] = (freq[s.teacher] || 0) + 1;
      }));
      const top = Object.keys(freq).sort((a, b) => freq[b] - freq[a])[0];
      return all.map(f => added.includes(f) ? {
        ...f,
        sheets: f.sheets.map(s => {
          const letter = lettersForGrade(data, s.grade).find(cl => cl.name.toLowerCase() === (s.letterName || "").toLowerCase());
          return {
            ...s,
            checked: s.teacher === top,
            letterId: letter ? letter.id : ""
          };
        })
      } : f);
    });
    setBusy(false);
  };
  const chosen = sheetsFlat.filter(x => x.s.checked).map(x => ({
    sheet: x.s,
    grade: x.s.grade,
    letterId: x.s.letterId,
    letterLabel: x.s.letterId ? classLetterName(data, x.s.letterId) : "",
    fileName: files[x.fi].name
  }));
  const preview = useMemo(() => planJournalImport(data, chosen), [data, files]);
  const totalAdd = preview.reduce((n, r) => n + r.addCount + r.absCount, 0);
  const totalSurplus = preview.reduce((n, r) => n + r.surplusCount, 0);
  const matchedAny = preview.some(r => r.matched > 0);
  const finish = mode => {
    const res = applyJournalPlan(structuredClone(data), preview, mode); // отчёт считаем на копии (update выполняется позже)
    update(d => {
      applyJournalPlan(d, preview, mode);
      return d;
    });
    setResult(res);
    setStep("done");
  };
  const start = () => totalAdd > 0 ? setStep("mode") : finish("lessons");
  const choiceBox = {
    border: "1.5px solid var(--line)",
    borderRadius: 12,
    padding: "12px 14px",
    cursor: "pointer",
    display: "grid",
    gap: 4,
    marginBottom: 8
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: () => step === "pick" && onClose()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      width: "min(900px, 96vw)"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, "\u0418\u043C\u043F\u043E\u0440\u0442 \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u0438\u0437 \u044D\u043B\u0435\u043A\u0442\u0440\u043E\u043D\u043D\u043E\u0433\u043E \u0436\u0443\u0440\u043D\u0430\u043B\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, step === "done" ? "Закрыть" : /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("input", {
    ref: inputRef,
    type: "file",
    multiple: true,
    accept: ".xlsx",
    style: {
      display: "none"
    },
    onChange: onPick
  }), step === "pick" && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: {
      marginTop: 0
    }
  }, "\u0412\u044B\u0431\u0435\u0440\u0438\u0442\u0435 \u0441\u043A\u0430\u0447\u0430\u043D\u043D\u044B\u0435 \u0438\u0437 \u0448\u043A\u043E\u043B\u044C\u043D\u043E\u0433\u043E \u0436\u0443\u0440\u043D\u0430\u043B\u0430 \u0442\u0430\u0431\u043B\u0438\u0446\u044B (\u043C\u043E\u0436\u043D\u043E \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u043E \u0441\u0440\u0430\u0437\u0443). \u041F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u043D\u0430\u0439\u0434\u0451\u0442 \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432 \u043F\u043E \u0444\u0430\u043C\u0438\u043B\u0438\u0438 \u0438 \u0438\u043C\u0435\u043D\u0438 \u0438 ", /*#__PURE__*/React.createElement("b", null, "\u0441\u0440\u0430\u0432\u043D\u0438\u0442 \u043A\u043E\u043B\u0438\u0447\u0435\u0441\u0442\u0432\u043E \u043E\u0446\u0435\u043D\u043E\u043A \u043A\u0430\u0436\u0434\u043E\u0433\u043E \u0437\u043D\u0430\u0447\u0435\u043D\u0438\u044F (1, 2, 3, 4, 5, \xAB\u041D\xBB)"), " \u0432 \u0442\u0430\u0431\u043B\u0438\u0446\u0435 \u0438 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435. \u041D\u0435\u0434\u043E\u0441\u0442\u0430\u044E\u0449\u0438\u0435 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435 \u043E\u0446\u0435\u043D\u043A\u0438 \u0434\u043E\u0431\u0430\u0432\u044F\u0442\u0441\u044F. \u041E\u0446\u0435\u043D\u043A\u0438, \u043A\u043E\u0442\u043E\u0440\u044B\u0435 \u0435\u0441\u0442\u044C \u0443 \u0432\u0430\u0441, \u043D\u043E \u0435\u0449\u0451 \u043D\u0435\u0442 \u0432 \u0442\u0430\u0431\u043B\u0438\u0446\u0435 (\u043D\u0435 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u044B \u0432 \u0448\u043A\u043E\u043B\u044C\u043D\u044B\u0439 \u0436\u0443\u0440\u043D\u0430\u043B), \u043F\u043E\u043A\u0430\u0436\u0435\u0442 \u043A\u043E\u043B\u043E\u043D\u043A\u0430 \xAB\u0412\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043D\u044B\u0435 \u043E\u0442\u043C\u0435\u0442\u043A\u0438\xBB \u0432 \u043A\u043E\u043D\u0446\u0435 \u0442\u0430\u0431\u043B\u0438\u0446\u044B \u0436\u0443\u0440\u043D\u0430\u043B\u0430. \u0415\u0441\u043B\u0438 \u043A\u043E\u043B\u0438\u0447\u0435\u0441\u0442\u0432\u0430 \u0441\u043E\u0432\u043F\u0430\u043B\u0438, \u0430 \u0434\u0430\u0442\u044B \u043D\u0435\u0442 \u2014 \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u043C\u0435\u043D\u044F\u0435\u0442\u0441\u044F."), /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    disabled: busy,
    onClick: () => inputRef.current && inputRef.current.click()
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 13
  }), " ", busy ? "Читаю…" : files.length ? "Добавить ещё файлы…" : "Выбрать файлы (.xlsx)…"), err && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--neg)",
      marginTop: 8
    }
  }, err), sheetsFlat.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      overflow: "auto",
      maxHeight: "44vh",
      marginTop: 10
    }
  }, sheetsFlat.map(({
    fi,
    si,
    s
  }) => {
    const rep = s.checked ? preview[chosen.findIndex(c => c.sheet === s)] : null;
    const letters = lettersForGrade(data, s.grade);
    return /*#__PURE__*/React.createElement("div", {
      key: fi + "-" + si,
      style: {
        padding: "8px 10px",
        borderBottom: "1px solid var(--line)",
        opacity: s.checked ? 1 : 0.65
      }
    }, /*#__PURE__*/React.createElement("div", {
      className: "row wrap",
      style: {
        gap: 8,
        alignItems: "center"
      }
    }, /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      checked: !!s.checked,
      onChange: e => patchSheet(fi, si, {
        checked: e.target.checked
      })
    }), /*#__PURE__*/React.createElement("b", {
      style: {
        flex: 1,
        minWidth: 200
      }
    }, s.title || s.sheetName), /*#__PURE__*/React.createElement("span", {
      className: "hint"
    }, "\u043A\u043B\u0430\u0441\u0441:"), /*#__PURE__*/React.createElement("b", null, s.grade || "?"), /*#__PURE__*/React.createElement("select", {
      className: "toolselect",
      value: s.letterId || "",
      onChange: e => patchSheet(fi, si, {
        letterId: e.target.value
      })
    }, /*#__PURE__*/React.createElement("option", {
      value: ""
    }, "\u0431\u0435\u0437 \u043B\u0438\u0442\u0435\u0440\u044B"), letters.map(cl => /*#__PURE__*/React.createElement("option", {
      key: cl.id,
      value: cl.id
    }, s.grade, cl.name)))), /*#__PURE__*/React.createElement("div", {
      className: "hint",
      style: {
        marginLeft: 26
      }
    }, "\u0424\u0430\u0439\u043B: ", files[fi].name, " \xB7 \u0434\u0430\u0442: ", s.columns.length, " \xB7 \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432 \u0432 \u0442\u0430\u0431\u043B\u0438\u0446\u0435: ", s.students.length), rep && /*#__PURE__*/React.createElement("div", {
      style: {
        marginLeft: 26,
        fontSize: 13,
        lineHeight: 1.5
      }
    }, "\u041D\u0430\u0439\u0434\u0435\u043D\u043E \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u044B: ", /*#__PURE__*/React.createElement("b", null, rep.matched), " \u0438\u0437 ", rep.studentsTotal, " \xB7 \u043D\u0435 \u0445\u0432\u0430\u0442\u0430\u0435\u0442 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435: ", /*#__PURE__*/React.createElement("b", null, rep.addCount), " \u043E\u0442\u043C\u0435\u0442\u043E\u043A, \xAB\u041D\xBB: ", /*#__PURE__*/React.createElement("b", null, rep.absCount), rep.surplusCount > 0 && /*#__PURE__*/React.createElement("span", {
      style: {
        color: "var(--neg)"
      }
    }, " \xB7 \u0435\u0449\u0451 \u043D\u0435 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043E \u0432 \u0448\u043A\u043E\u043B\u044C\u043D\u044B\u0439 \u0436\u0443\u0440\u043D\u0430\u043B: ", rep.surplusCount), rep.doubles > 0 && ` · в ячейке и оценка, и «Н» (взята оценка): ${rep.doubles}`, rep.skipped && /*#__PURE__*/React.createElement("div", {
      style: {
        color: "var(--neg)"
      }
    }, "\u041D\u0438 \u043E\u0434\u043D\u043E\u0433\u043E \u0443\u0447\u0435\u043D\u0438\u043A\u0430 \u0438\u0437 \u044D\u0442\u043E\u0433\u043E \u043B\u0438\u0441\u0442\u0430 \u043D\u0435\u0442 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B\u0435 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u044B (\u0434\u043B\u044F \u0432\u044B\u0431\u0440\u0430\u043D\u043D\u043E\u0439 \u043B\u0438\u0442\u0435\u0440\u044B) \u2014 \u043B\u0438\u0441\u0442 \u0431\u0443\u0434\u0435\u0442 \u043F\u0440\u043E\u043F\u0443\u0449\u0435\u043D. \u0421\u043D\u0430\u0447\u0430\u043B\u0430 \u0434\u043E\u0431\u0430\u0432\u044C\u0442\u0435 \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432 \u0432\u043E \u0432\u043A\u043B\u0430\u0434\u043A\u0435 \xAB\u0421\u043F\u0438\u0441\u043A\u0438 \u043A\u043B\u0430\u0441\u0441\u043E\u0432\xBB."), rep.unmatched.length > 0 && rep.matched > 0 && /*#__PURE__*/React.createElement("div", {
      className: "hint"
    }, "\u041D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u044B (\u043D\u0435 \u0431\u0443\u0434\u0443\u0442 \u043E\u0431\u0440\u0430\u0431\u043E\u0442\u0430\u043D\u044B): ", rep.unmatched.join(", ")), rep.bad.length > 0 && /*#__PURE__*/React.createElement("div", {
      className: "hint"
    }, "\u041D\u0435\u043F\u043E\u043D\u044F\u0442\u043D\u044B\u0435 \u0437\u043D\u0430\u0447\u0435\u043D\u0438\u044F \u0432 \u044F\u0447\u0435\u0439\u043A\u0430\u0445 \u043F\u0440\u043E\u043F\u0443\u0449\u0435\u043D\u044B: ", rep.bad.slice(0, 5).join("; "), rep.bad.length > 5 ? "…" : "")));
  })), sheetsFlat.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginTop: 6
    }
  }, "\u041F\u043E \u0443\u043C\u043E\u043B\u0447\u0430\u043D\u0438\u044E \u043E\u0442\u043C\u0435\u0447\u0435\u043D\u044B \u043B\u0438\u0441\u0442\u044B \u0441\u0430\u043C\u043E\u0433\u043E \u0447\u0430\u0441\u0442\u043E\u0433\u043E \u0443\u0447\u0438\u0442\u0435\u043B\u044F \u0432 \u0444\u0430\u0439\u043B\u0430\u0445 (\u044D\u0442\u043E \u0432\u0430\u0448\u0438 \u0433\u0440\u0443\u043F\u043F\u044B); \u043B\u0438\u0441\u0442\u044B \u0434\u0440\u0443\u0433\u0438\u0445 \u0443\u0447\u0438\u0442\u0435\u043B\u0435\u0439 \u043C\u043E\u0436\u043D\u043E \u0432\u043A\u043B\u044E\u0447\u0438\u0442\u044C \u0433\u0430\u043B\u043E\u0447\u043A\u043E\u0439. \u0418\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0435 \u043C\u043E\u0436\u043D\u043E \u043E\u0442\u043C\u0435\u043D\u0438\u0442\u044C \u043A\u043D\u043E\u043F\u043A\u043E\u0439 \xAB\u21B6 \u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C\xBB (Ctrl+Z).")), step === "mode" && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("p", {
    style: {
      marginTop: 0
    }
  }, "\u0412 \u0436\u0443\u0440\u043D\u0430\u043B\u0435 \u043D\u0435 \u0445\u0432\u0430\u0442\u0430\u0435\u0442 \u043E\u0442\u043C\u0435\u0442\u043E\u043A: ", /*#__PURE__*/React.createElement("b", null, totalAdd), ". \u041A\u0430\u043A \u0438\u0445 \u0434\u043E\u0431\u0430\u0432\u0438\u0442\u044C?"), /*#__PURE__*/React.createElement("div", {
    style: choiceBox,
    onClick: () => finish("lessons")
  }, /*#__PURE__*/React.createElement("b", null, "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0443\u0440\u043E\u043A\u0438, \u043A\u043E\u0442\u043E\u0440\u044B\u0445 \u043D\u0435 \u0445\u0432\u0430\u0442\u0430\u0435\u0442, \u0438 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u043D\u0430 \u0434\u0430\u0442\u044B \u0438\u0437 \u0442\u0430\u0431\u043B\u0438\u0446\u044B"), /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      fontSize: 12.5
    }
  }, "\u041E\u0442\u043C\u0435\u0442\u043A\u0438 \u0432\u0441\u0442\u0430\u043D\u0443\u0442 \u0440\u043E\u0432\u043D\u043E \u043D\u0430 \u0442\u0435 \u0447\u0438\u0441\u043B\u0430, \u0447\u0442\u043E \u0432 \u0442\u0430\u0431\u043B\u0438\u0446\u0435. \u0414\u043B\u044F \u0434\u0430\u0442 \u0431\u0435\u0437 \u0443\u0440\u043E\u043A\u0430 \u0431\u0443\u0434\u0443\u0442 \u0441\u043E\u0437\u0434\u0430\u043D\u044B \u043F\u0443\u0441\u0442\u044B\u0435 \u0443\u0440\u043E\u043A\u0438 (\u0431\u0435\u0437 \u0442\u0435\u043C\u044B, \u0441 \u043F\u043E\u043C\u0435\u0442\u043A\u043E\u0439 \xAB\u0422\u0440\u0435\u0431\u0443\u0435\u0442 \u0434\u043E\u0440\u0430\u0431\u043E\u0442\u043A\u0438\xBB).")), /*#__PURE__*/React.createElement("div", {
    style: choiceBox,
    onClick: () => finish("marks")
  }, /*#__PURE__*/React.createElement("b", null, "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0442\u043E\u043B\u044C\u043A\u043E \u043E\u0442\u043C\u0435\u0442\u043A\u0438"), /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      fontSize: 12.5
    }
  }, "\u0423\u0440\u043E\u043A\u0438 \u043D\u0435 \u0441\u043E\u0437\u0434\u0430\u044E\u0442\u0441\u044F. \u041A\u0430\u0436\u0434\u0430\u044F \u043D\u0435\u0434\u043E\u0441\u0442\u0430\u044E\u0449\u0430\u044F \u043E\u0442\u043C\u0435\u0442\u043A\u0430 \u0432\u0441\u0442\u0430\u043D\u0435\u0442 \u0432 \u043F\u0443\u0441\u0442\u0443\u044E \u043A\u043B\u0435\u0442\u043A\u0443 \u044D\u0442\u043E\u0433\u043E \u0443\u0447\u0435\u043D\u0438\u043A\u0430 \u0432 \u0443\u0436\u0435 \u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u044E\u0449\u0435\u043C \u0441\u0442\u043E\u043B\u0431\u0446\u0435-\u0434\u0430\u0442\u0435 \u0434\u043E \u0441\u0435\u0433\u043E\u0434\u043D\u044F\u0448\u043D\u0435\u0433\u043E \u0434\u043D\u044F, \u0431\u043B\u0438\u0436\u0430\u0439\u0448\u0443\u044E \u043F\u043E \u0434\u0430\u0442\u0435 \u043A \u0434\u0430\u0442\u0435 \u0438\u0437 \u0442\u0430\u0431\u043B\u0438\u0446\u044B.")), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setStep("pick")
  }, "\u2190 \u041D\u0430\u0437\u0430\u0434")), step === "done" && result && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gap: 8,
      lineHeight: 1.55
    }
  }, /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("b", null, "\u0413\u043E\u0442\u043E\u0432\u043E."), " \u0414\u043E\u0431\u0430\u0432\u043B\u0435\u043D\u043E \u043E\u0442\u043C\u0435\u0442\u043E\u043A: ", result.added, ", \xAB\u041D\xBB: ", result.absAdded, ". \u0421\u043E\u0437\u0434\u0430\u043D\u043E \u0443\u0440\u043E\u043A\u043E\u0432: ", result.lessonsCreated, ".", result.moved > 0 && ` Отметок поставлено не на дату из таблицы (в ближайшую свободную клетку): ${result.moved}.`), totalSurplus > 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041E\u0442\u043C\u0435\u0442\u043A\u0438, \u043A\u043E\u0442\u043E\u0440\u044B\u0445 \u0435\u0449\u0451 \u043D\u0435\u0442 \u0432 \u0442\u0430\u0431\u043B\u0438\u0446\u0435 \u0448\u043A\u043E\u043B\u044C\u043D\u043E\u0433\u043E \u0436\u0443\u0440\u043D\u0430\u043B\u0430 (", totalSurplus, "), \u043E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u044E\u0442\u0441\u044F \u0432 \u043A\u043E\u043B\u043E\u043D\u043A\u0435 \xAB\u0412\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043D\u044B\u0435 \u043E\u0442\u043C\u0435\u0442\u043A\u0438\xBB \u0432 \u043A\u043E\u043D\u0446\u0435 \u0442\u0430\u0431\u043B\u0438\u0446\u044B \u0436\u0443\u0440\u043D\u0430\u043B\u0430."), result.lessonsCreated > 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0421\u043E\u0437\u0434\u0430\u043D\u043D\u044B\u0435 \u0443\u0440\u043E\u043A\u0438 \u043F\u0443\u0441\u0442\u044B\u0435 (\u0431\u0435\u0437 \u0442\u0435\u043C\u044B) \u0438 \u043F\u043E\u043C\u0435\u0447\u0435\u043D\u044B \xAB\u0422\u0440\u0435\u0431\u0443\u0435\u0442 \u0434\u043E\u0440\u0430\u0431\u043E\u0442\u043A\u0438\xBB: \u043E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0438\u0445 \u0438 \u0437\u0430\u043F\u043E\u043B\u043D\u0438\u0442\u0435."), result.blocked.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--neg)"
    }
  }, "\u041D\u0435 \u0434\u043E\u0431\u0430\u0432\u043B\u0435\u043D\u043E (", result.blocked.length, "):", /*#__PURE__*/React.createElement("ul", {
    style: {
      margin: "4px 0 0 18px"
    }
  }, result.blocked.slice(0, 20).map((b, i) => /*#__PURE__*/React.createElement("li", {
    key: i
  }, b)))), /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041F\u043E\u0432\u0442\u043E\u0440\u043D\u044B\u0439 \u0437\u0430\u043F\u0443\u0441\u043A \u0441 \u0442\u0435\u043C\u0438 \u0436\u0435 \u0444\u0430\u0439\u043B\u0430\u043C\u0438 \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u0434\u043E\u0431\u0430\u0432\u0438\u0442 \u2014 \u043A\u043E\u043B\u0438\u0447\u0435\u0441\u0442\u0432\u0430 \u0443\u0436\u0435 \u0441\u043E\u0432\u043F\u0430\u0434\u0430\u044E\u0442."))), step === "pick" && /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: !chosen.length || !matchedAny,
    onClick: start
  }, totalAdd > 0 ? `Продолжить (не хватает в журнале: ${totalAdd})` : "Не хватает ничего — запомнить таблицу для сверки"))));
}

// ======================= Журнал v3 (перенос из «Конструктора уроков») =======================
const JOURNAL_MARKED_BG = "#e3e5e8";
const JR_VALUES = [1, 2, 3, 4, 5, "Н"];
const RU_COUNT_WORDS = {
  1: "одну",
  2: "две",
  3: "три",
  4: "четыре",
  5: "пять"
};

// Переключатель-ползунок.
function ToggleSwitch({
  checked,
  onChange,
  label,
  small
}) {
  return /*#__PURE__*/React.createElement("label", {
    style: {
      display: "inline-flex",
      alignItems: "center",
      gap: 7,
      cursor: "pointer",
      fontSize: small ? 12 : 12.5,
      userSelect: "none"
    }
  }, /*#__PURE__*/React.createElement("span", {
    onClick: e => {
      e.preventDefault();
      onChange(!checked);
    },
    style: {
      width: 34,
      height: 19,
      borderRadius: 10,
      background: checked ? "#3b82c4" : "#b7c4d1",
      position: "relative",
      flexShrink: 0,
      transition: "background .15s"
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      position: "absolute",
      top: 2,
      left: checked ? 17 : 2,
      width: 15,
      height: 15,
      borderRadius: "50%",
      background: "#fff",
      transition: "left .15s",
      boxShadow: "0 1px 2px rgba(0,0,0,.3)"
    }
  })), label);
}
// Разбор введённых в клетку отметок: «5 4», «н», «5, 4» → значения; bad — непонятные части.
function parseMarkTokens(text, scale) {
  const toks = String(text || "").replace(/[;,]+/g, " ").trim().split(/\s+/).filter(Boolean);
  const out = [],
    bad = [];
  const lo = scale === "100" ? 0 : 1,
    hi = scale === "100" ? 100 : 5;
  toks.forEach(t => {
    const u = t.toUpperCase();
    if (u === "Н" || u === "H" || u === "N") {
      out.push("Н");
      return;
    }
    const n = Number(t.replace(",", "."));
    if (Number.isInteger(n) && n >= lo && n <= hi) out.push(n);else bad.push(t);
  });
  return {
    out,
    bad
  };
}
// Поле одной итоговой отметки (четвертная/годовая/экзамен): записывается при выходе из поля или по Enter.
function PeriodMarkInput({
  value,
  placeholder,
  onCommit,
  width
}) {
  const cur = value == null ? "" : String(value);
  const [text, setText] = useState(cur);
  const [bad, setBad] = useState(false);
  useEffect(() => {
    setText(cur);
    setBad(false);
  }, [cur]);
  const commit = () => {
    if (text.trim() === cur) return;
    if (!onCommit(text)) setBad(true);
  };
  return /*#__PURE__*/React.createElement("input", {
    value: text,
    placeholder: placeholder,
    onChange: e => {
      setText(e.target.value);
      setBad(false);
    },
    onBlur: commit,
    onKeyDown: e => {
      if (e.key === "Enter") e.target.blur();
      if (e.key === "Escape") {
        setText(cur);
        e.target.blur();
      }
    },
    style: {
      width: width || 46,
      textAlign: "center",
      fontWeight: 800,
      border: bad ? "1.5px solid #c0392b" : "1.5px solid var(--line)",
      borderRadius: 6,
      padding: "4px 2px",
      background: "#fff"
    }
  });
}

// Колонка «Выставленные отметки»: сколько отметок каждого значения есть в журнале Конструктора, но ещё нет в
// импортированной таблице электронного журнала. null — таблица для этой четверти не импортировалась / сверять не с чем.
function missingInFile(data, student, q) {
  const sn = (data.calendar.journalFileSnapshots || []).find(x => x.grade === student.class && (x.letterId || null) === (student.classLetterId || null) && quarterOfDate(data, x.minDate) === q);
  const fc = sn && sn.students[student.id] ? sn.students[student.id] : null;
  const marksQ = data.calendar.marks.filter(m => m.studentId === student.id && (m.scale || "5") === "5" && quarterOfDate(data, m.date) === q);
  const anyPosted = marksQ.some(m => m.posted);
  if (!fc && !anyPosted) return null;
  const zero = () => ({
    1: 0,
    2: 0,
    3: 0,
    4: 0,
    5: 0,
    "Н": 0
  });
  const inTot = zero(),
    inPosted = zero(),
    laterTot = zero(),
    laterPosted = zero();
  marksQ.forEach(m => {
    const v = m.value === "Н" ? "Н" : +m.value;
    if (inTot[v] === undefined) return;
    if (sn && m.date < sn.minDate && !m.posted) return;
    const later = sn && m.date > sn.maxDate;
    (later ? laterTot : inTot)[v]++;
    if (m.posted) (later ? laterPosted : inPosted)[v]++;
  });
  const missing = [];
  JR_VALUES.forEach(v => {
    const inFile = Math.max(fc ? fc[v] || 0 : 0, inPosted[v]);
    const n = Math.max(0, inTot[v] - inFile) + (laterTot[v] - laterPosted[v]);
    if (n > 0) missing.push({
      value: v,
      n
    });
  });
  return {
    missing,
    importedAt: sn ? sn.importedAt : null,
    fileName: sn ? sn.fileName : ""
  };
}
function missingText(missing) {
  return "Не выставлено: " + missing.map(m => `${RU_COUNT_WORDS[m.n] || m.n} «${m.value}»`).join(", ");
}

// ---- Выгрузка четверти в Excel (.xlsx) своим мини-писателем (ZIP без сжатия) ----
const xmlEsc = t => String(t).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
const xlsxCol = i => {
  let n = i + 1,
    out = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    out = String.fromCharCode(65 + r) + out;
    n = Math.floor((n - 1) / 26);
  }
  return out;
};
// sheets: [{name, rows: [[ячейка…]], widths?: [число…]}]; ячейка — строка, число, либо {v, s: 1 (шапка) | 2 (заголовок)}
function buildXlsxBlob(sheets) {
  const sheetXml = sh => {
    const rows = sh.rows.map((row, ri) => {
      const cells = row.map((cell, ci) => {
        if (cell === null || cell === undefined || cell === "") return "";
        const o = typeof cell === "object" ? cell : {
          v: cell
        };
        const ref = xlsxCol(ci) + (ri + 1);
        const st = o.s ? ` s="${o.s}"` : "";
        return typeof o.v === "number" ? `<c r="${ref}"${st}><v>${o.v}</v></c>` : `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${xmlEsc(o.v)}</t></is></c>`;
      }).join("");
      return `<row r="${ri + 1}">${cells}</row>`;
    }).join("");
    const cols = sh.widths ? "<cols>" + sh.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join("") + "</cols>" : "";
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${cols}<sheetData>${rows}</sheetData></worksheet>`;
  };
  const safeName = (n, i) => String(n).replace(/[\[\]:*?\/\\]/g, " ").slice(0, 31) || "Лист" + (i + 1);
  const files = [{
    name: "[Content_Types].xml",
    data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' + sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("") + "</Types>"
  }, {
    name: "_rels/.rels",
    data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>'
  }, {
    name: "xl/workbook.xml",
    data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' + sheets.map((sh, i) => `<sheet name="${xmlEsc(safeName(sh.name, i))}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("") + "</sheets></workbook>"
  }, {
    name: "xl/_rels/workbook.xml.rels",
    data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("") + `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`
  }, {
    name: "xl/styles.xml",
    data: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFDCE9F5"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs></styleSheet>'
  }, ...sheets.map((sh, i) => ({
    name: `xl/worksheets/sheet${i + 1}.xml`,
    data: sheetXml(sh)
  }))].map(f => ({
    name: f.name,
    data: strToBytes(f.data)
  }));
  return new Blob([makeZip(files)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
  });
}
// Выгрузка отметок одной четверти выбранных учеников: лист «Журнал» (ученики × даты) и «Список отметок».
function buildQuarterXlsx(data, students, q, groupLabel, scale) {
  const ids = new Set(students.map(st => st.id));
  const marks = data.calendar.marks.filter(m => ids.has(m.studentId) && quarterOfDate(data, m.date) === q);
  const dates = Array.from(new Set(marks.map(m => m.date))).sort();
  const sc = scale || "5";
  const nameOf = m => (data.calendar.students.find(x => x.id === m.studentId) || {}).name || "";
  const head = [{
    v: "№",
    s: 1
  }, {
    v: "Ф.И.О.",
    s: 1
  }, ...dates.map(d => ({
    v: formatDateRu(d).slice(0, 5),
    s: 1
  })), {
    v: "Средний балл",
    s: 1
  }, {
    v: "Кол-во",
    s: 1
  }, {
    v: "Четвертная",
    s: 1
  }];
  const rows = [[{
    v: `${groupLabel} — ${q + 1} четверть`,
    s: 2
  }], [`Выгружено ${formatDateRu(todayIso())}`], [], head];
  students.forEach((st, i) => {
    const ms = marks.filter(m => m.studentId === st.id);
    const cells = dates.map(d => {
      const here = ms.filter(m => m.date === d);
      if (!here.length) return "";
      if (here.length === 1 && typeof here[0].value === "number") return here[0].value;
      return here.map(m => m.value).join(" ");
    });
    const avg = quarterAvgOf(data, st.id, q, sc);
    const cnt = numericMarksOf(data, st.id, sc).filter(m => quarterOfDate(data, m.date) === q).length;
    const pm = (data.calendar.periodMarks || []).find(p => p.studentId === st.id && p.kind === "quarter" && p.q === q);
    rows.push([i + 1, st.name, ...cells, avg == null ? "" : avg, cnt, pm ? pm.value : ""]);
  });
  const list = [[{
    v: "Ученик",
    s: 1
  }, {
    v: "Дата",
    s: 1
  }, {
    v: "Отметка",
    s: 1
  }, {
    v: "Шкала",
    s: 1
  }, {
    v: "Причина «Н»",
    s: 1
  }, {
    v: "Комментарий",
    s: 1
  }, {
    v: "Тема урока",
    s: 1
  }, {
    v: "Выставлена в эл. журнал",
    s: 1
  }]];
  [...marks].sort((a, b) => nameOf(a) < nameOf(b) ? -1 : nameOf(a) > nameOf(b) ? 1 : a.date < b.date ? -1 : 1).forEach(m => {
    const ls = m.lessonId ? data.lessons.find(l => l.id === m.lessonId) : null;
    list.push([nameOf(m), formatDateRu(m.date), typeof m.value === "number" ? m.value : String(m.value), (m.scale || "5") === "100" ? "100-балльная" : "5-балльная", m.value === "Н" ? m.excused === false ? "неуважительная" : "уважительная" : "", m.reason || "", ls ? ls.topic || "" : "", m.posted ? "да" : "нет"]);
  });
  return {
    blob: buildXlsxBlob([{
      name: "Журнал",
      rows,
      widths: [5, 28, ...dates.map(() => 8), 13, 8, 11]
    }, {
      name: "Список отметок",
      rows: list,
      widths: [28, 12, 9, 13, 16, 30, 34, 22]
    }]),
    marksCount: marks.length
  };
}
// Окно «Начать новую четверть»: выбор четверти, предупреждение, подтверждение.
function NewQuarterModal({
  data,
  students,
  groupLabel,
  qc,
  defaultQ,
  onClose,
  onConfirm
}) {
  const ids = new Set(students.map(st => st.id));
  const counts = Array.from({
    length: qc
  }).map((_, q) => data.calendar.marks.filter(m => ids.has(m.studentId) && quarterOfDate(data, m.date) === q).length);
  const [q, setQ] = useState(defaultQ);
  const hasPeriod = (data.calendar.periodMarks || []).filter(p => ids.has(p.studentId) && p.kind === "quarter" && p.q === q).length;
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      width: "min(520px, 96vw)"
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", null, "\u041D\u0430\u0447\u0430\u0442\u044C \u043D\u043E\u0432\u0443\u044E \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u044C \u2014 ", groupLabel), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement("p", {
    style: {
      marginTop: 0
    }
  }, "\u041A\u0430\u043A\u0443\u044E \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u044C \u0437\u0430\u043A\u0440\u044B\u0442\u044C? \u0412\u0441\u0435 \u0435\u0451 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u0438 \xAB\u041D\xBB \u044D\u0442\u0438\u0445 \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432 \u0441\u043E\u0445\u0440\u0430\u043D\u044F\u0442\u0441\u044F \u0432 \u0444\u0430\u0439\u043B Excel, \u0430 \u0438\u0437 \u0436\u0443\u0440\u043D\u0430\u043B\u0430 \u0438\u0441\u0447\u0435\u0437\u043D\u0443\u0442. \u0427\u0435\u0442\u0432\u0435\u0440\u0442\u043D\u0430\u044F \u043E\u0442\u043C\u0435\u0442\u043A\u0430 \u043E\u0441\u0442\u0430\u043D\u0435\u0442\u0441\u044F."), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 6,
      marginBottom: 10
    }
  }, counts.map((c, i) => /*#__PURE__*/React.createElement(TrackChip, {
    key: i,
    label: `${i + 1} четверть · отметок: ${c}`,
    color: "#8593a3",
    active: q === i,
    onClick: () => setQ(i)
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12.5,
      lineHeight: 1.55,
      background: "var(--accent-soft)",
      borderRadius: 10,
      padding: "9px 12px"
    }
  }, "1. \u0421\u043A\u0430\u0447\u0430\u0435\u0442\u0441\u044F \u0444\u0430\u0439\u043B Excel (\u043B\u0438\u0441\u0442 \xAB\u0416\u0443\u0440\u043D\u0430\u043B\xBB \u2014 \u0443\u0447\u0435\u043D\u0438\u043A\u0438 \xD7 \u0434\u0430\u0442\u044B, \u043B\u0438\u0441\u0442 \xAB\u0421\u043F\u0438\u0441\u043E\u043A \u043E\u0442\u043C\u0435\u0442\u043E\u043A\xBB \u2014 \u043F\u043E \u043E\u0434\u043D\u043E\u0439 \u043E\u0442\u043C\u0435\u0442\u043A\u0435 \u0432 \u0441\u0442\u0440\u043E\u043A\u0435).", /*#__PURE__*/React.createElement("br", null), "2. \u041F\u043E\u0441\u043B\u0435 \u0432\u0430\u0448\u0435\u0433\u043E \u043F\u043E\u0434\u0442\u0432\u0435\u0440\u0436\u0434\u0435\u043D\u0438\u044F, \u0447\u0442\u043E \u0444\u0430\u0439\u043B \u0441\u043E\u0445\u0440\u0430\u043D\u0451\u043D, \u0432 ", groupLabel, " \u0431\u0443\u0434\u0435\u0442 \u0443\u0434\u0430\u043B\u0435\u043D\u043E \u043E\u0442\u043C\u0435\u0442\u043E\u043A: ", /*#__PURE__*/React.createElement("b", null, counts[q]), ". \u0427\u0435\u0442\u0432\u0435\u0440\u0442\u043D\u044B\u0445 \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u043E\u0441\u0442\u0430\u043D\u0435\u0442\u0441\u044F: ", /*#__PURE__*/React.createElement("b", null, hasPeriod), ".", /*#__PURE__*/React.createElement("br", null), "3. \u041F\u0443\u0441\u0442\u044B\u0435 \u0434\u0430\u0442\u044B \u0437\u0430\u043A\u0440\u044B\u0442\u043E\u0439 \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u0438 \u0441\u043A\u0440\u043E\u044E\u0442\u0441\u044F. \u0418\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0435 \u043C\u043E\u0436\u043D\u043E \u043E\u0442\u043C\u0435\u043D\u0438\u0442\u044C \u043A\u043D\u043E\u043F\u043A\u043E\u0439 \xAB\u21B6 \u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C\xBB (Ctrl+Z)."), counts[q] === 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--neg)",
      marginTop: 8
    }
  }, "\u0412 \u044D\u0442\u043E\u0439 \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u0438 \u043D\u0435\u0442 \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u2014 \u0432\u044B\u0433\u0440\u0443\u0436\u0430\u0442\u044C \u0438 \u043E\u0447\u0438\u0449\u0430\u0442\u044C \u043D\u0435\u0447\u0435\u0433\u043E.")), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: counts[q] === 0,
    onClick: () => onConfirm(q)
  }, "\uD83D\uDCE5 \u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u0432 Excel \u0438 \u043E\u0447\u0438\u0441\u0442\u0438\u0442\u044C"))));
}
function JournalSection({
  data,
  update,
  trackKey,
  setTrackKey,
  tracks,
  letterFilter,
  setLetterFilter,
  jui,
  setJui
}) {
  const [cellModal, setCellModal] = useState(null);
  const [reportDate, setReportDate] = useState(null);
  const [addMarkOpen, setAddMarkOpen] = useState(false);
  const [xlsxOpen, setXlsxOpen] = useState(false);
  const [autoTypeOpen, setAutoTypeOpen] = useState(false);
  const [newQuarterOpen, setNewQuarterOpen] = useState(false);
  const [hoverCell, setHoverCell] = useState(null);
  const [hoverMark, setHoverMark] = useState(null);
  const hoverTimer = useRef(null);
  const effectiveKey = trackKey && tracks.some(t => trackKeyOf(t.kind, t.value) === trackKey) ? trackKey : tracks[0] ? trackKeyOf(tracks[0].kind, tracks[0].value) : null;
  useEffect(() => {
    if (trackKey !== effectiveKey) setTrackKey(effectiveKey);
  }, [trackKey, effectiveKey]);
  const [kind, value] = (effectiveKey || "").split(":");
  const jFirstRender = useRef(true);
  useEffect(() => {
    if (jFirstRender.current) {
      jFirstRender.current = false;
      return;
    }
    const [jk, jv] = (effectiveKey || "").split(":");
    if (!letterValidFor(data, jk, jv, letterFilter)) setLetterFilter("all");
  }, [effectiveKey]);
  if (tracks.length === 0) return /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u0421\u043D\u0430\u0447\u0430\u043B\u0430 \u0441\u043E\u0437\u0434\u0430\u0439\u0442\u0435 \u043A\u043B\u0430\u0441\u0441 \u0438\u043B\u0438 \u043A\u0443\u0440\u0441 \u0432 \u0440\u0430\u0437\u0434\u0435\u043B\u0435 \xAB\u0423\u0440\u043E\u043A\u0438\xBB.");
  const letters = kind === "class" ? lettersForGrade(data, value) : [];
  const students = data.calendar.students.filter(s => kind === "course" ? s.courseId === value : s.class === value && !s.courseId).filter(s => {
    if (kind !== "class" || letterFilter === "all") return true;
    if (letterFilter === "none") return !s.classLetterId;
    return s.classLetterId === letterFilter;
  }).sort((a, b) => (a.name || "").localeCompare(b.name || "", "ru"));
  const qc = quarterCountOf(data);
  const scalesHere = Array.from(new Set(students.flatMap(s => scalesInUse(settingsForStudent(data, s).gradingSystem))));
  const gridScale = scalesHere.includes("5") ? "5" : scalesHere[0] || "5";
  const settingsHere = getClassSettings(data, kind, value, kind === "class" ? letterFilter === "all" || letterFilter === "none" ? null : letterFilter : null);
  const tIso = todayIso();
  const groupLessons = data.lessons.filter(l => !l.archived && (kind === "course" ? l.courseId === value : l.class === value && !l.courseId) && (kind !== "class" || letterFilter === "all" || (letterFilter === "none" ? !l.classLetterId : l.classLetterId === letterFilter)));
  const groupLabel = kind === "course" ? data.courses.find(c => c.id === value)?.name || "Курс" : `${value}${letterFilter !== "all" && letterFilter !== "none" ? classLetterName(data, letterFilter) : ""} класс`;
  const lessonForDate = dt => groupLessons.find(l => (l.actualDate || l.plannedDate) === dt) || null;
  const marksOfCell = (sid, date) => studentMarksOf(data, sid, gridScale).filter(m => m.date === date);
  const datesForQuarter = q => {
    const lessonDates = groupLessons.map(l => l.actualDate || l.plannedDate).filter(Boolean);
    const markDates = students.flatMap(s => studentMarksOf(data, s.id, gridScale).map(m => m.date));
    const hide = jui.hideEmpty && !jui.editing; // в режиме правки пустые даты видны — иначе в них нельзя поставить отметку
    return Array.from(new Set([...lessonDates, ...markDates])).filter(dt => quarterOfDate(data, dt) === q).filter(dt => !hide || students.some(s => marksOfCell(s.id, dt).length > 0)).sort();
  };

  // «Выставлена» — отметка уже внесена в школьный электронный журнал (клетка серая, из колонки «Выставленные отметки» исчезает)
  const setPosted = (ids, on) => update(d => {
    const set = new Set(ids);
    d.calendar.marks.forEach(m => {
      if (set.has(m.id)) m.posted = on;
    });
    return d;
  });
  const dateMarks = date => students.flatMap(st => marksOfCell(st.id, date));
  const isDateMarked = date => {
    const ms = dateMarks(date);
    return ms.length > 0 && ms.every(m => m.posted);
  };
  const toggleDateMarked = date => {
    const ms = dateMarks(date);
    if (!ms.length) return;
    setPosted(ms.map(m => m.id), !isDateMarked(date));
  };
  const showMarkHover = (e, id) => {
    clearTimeout(hoverTimer.current);
    const r = e.currentTarget.getBoundingClientRect();
    const below = r.top < 90;
    setHoverMark({
      id,
      x: r.left,
      y: below ? r.bottom + 2 : r.top - 3,
      below
    });
  };
  const hideMarkHover = () => {
    clearTimeout(hoverTimer.current);
    hoverTimer.current = setTimeout(() => setHoverMark(null), 220);
  };

  // ---- режим «Править»: отметки вводятся прямо в клетках (несколько через пробел), ✓ записывает всё разом ----
  const draft = jui.draft || {};
  const existingText = (sid, date) => marksOfCell(sid, date).map(m => m.value).join(" ");
  const cellText = (sid, date) => draft[sid + "|" + date] !== undefined ? draft[sid + "|" + date] : existingText(sid, date);
  const sameValues = (a, b) => a.map(String).sort().join(",") === b.map(String).sort().join(",");
  const draftPlans = () => {
    const plans = [],
      errors = [];
    Object.entries(draft).forEach(([key, text]) => {
      const [sid, date] = key.split("|");
      const st = students.find(x => x.id === sid);
      const {
        out,
        bad
      } = parseMarkTokens(text, gridScale);
      if (bad.length) {
        errors.push(`${st ? st.name : "?"}, ${formatDateRu(date)}: «${bad.join(" ")}»`);
        return;
      }
      if (sameValues(marksOfCell(sid, date).map(m => m.value), out)) return;
      plans.push({
        sid,
        date,
        out
      });
    });
    return {
      plans,
      errors
    };
  };
  const changedCount = draftPlans().plans.length;
  const commitDraft = async () => {
    const {
      plans,
      errors
    } = draftPlans();
    if (errors.length) {
      await Confirm("Не удалось понять:\n" + errors.slice(0, 8).join("\n") + "\n\nДопустимо " + (gridScale === "100" ? "0–100" : "1–5") + " или Н; несколько отметок в одной клетке — через пробел.", "Понятно");
      return;
    }
    if (plans.length) {
      update(d => {
        plans.forEach(({
          sid,
          date,
          out
        }) => {
          const free = d.calendar.marks.filter(m => m.studentId === sid && m.date === date && (m.scale || "5") === gridScale);
          const toAdd = [];
          out.forEach(v => {
            const i = free.findIndex(m => String(m.value) === String(v));
            if (i >= 0) free.splice(i, 1);else toAdd.push(v);
          });
          const removedIds = free.map(m => m.id);
          if (removedIds.length) {
            d.calendar.marks = d.calendar.marks.filter(m => !removedIds.includes(m.id));
            d.lessons.forEach(l => (l.items || []).forEach(it => {
              if (Array.isArray(it.students)) it.students = it.students.filter(en => !removedIds.includes(en.markId));
            }));
          }
          const l = lessonForDate(date);
          toAdd.forEach(v => d.calendar.marks.push({
            id: uid(),
            studentId: sid,
            date,
            value: v,
            reason: "",
            scale: gridScale,
            excused: v === "Н" ? true : null,
            lessonId: l ? l.id : null,
            exerciseId: null,
            questionId: null,
            posted: false
          }));
        });
        return d;
      });
    }
    setJui({
      editing: false,
      draft: {}
    });
  };
  const startInlineEdit = (sid, date, append) => {
    setJui(p => {
      const key = sid + "|" + date;
      const dr = {
        ...(p.draft || {})
      };
      if (append) dr[key] = (dr[key] !== undefined ? dr[key] : existingText(sid, date)) + " ";
      return {
        editing: true,
        draft: dr
      };
    });
    setTimeout(() => {
      const el = document.querySelector(`input[data-cell="${sid}|${date}"]`);
      if (!el) return;
      el.focus();
      if (append) el.setSelectionRange(el.value.length, el.value.length);else el.select();
    }, 60);
  };
  const focusCell = (sid, date) => {
    const el = document.querySelector(`input[data-cell="${sid}|${date}"]`);
    if (el) {
      el.focus();
      el.select();
    }
  };

  // ---- «Начать новую четверть» ----
  const startNewQuarter = async q => {
    let built;
    try {
      built = buildQuarterXlsx(data, students, q, groupLabel, gridScale);
    } catch (e) {
      await Confirm("Не удалось подготовить файл Excel: " + (e && e.message || e) + "\nОтметки не тронуты.", "Понятно");
      return;
    }
    const fileName = `Журнал ${groupLabel} — ${q + 1} четверть (${todayIso()}).xlsx`;
    const url = URL.createObjectURL(built.blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = fileName;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 10000);
    await new Promise(r => setTimeout(r, 1200));
    if (!(await Confirm(`Файл «${fileName}» сохранился?\n\nЕсли вы нажали «Отмена» в окне сохранения, нажмите «Нет, отмена» здесь — отметки останутся на месте.\n\nУдалить из журнала ${built.marksCount} отметок ${q + 1} четверти (четвертная останется)?`, "Да, файл сохранён — удалить"))) return;
    const idSet = new Set(students.map(st => st.id));
    update(d => {
      const removed = new Set(d.calendar.marks.filter(m => idSet.has(m.studentId) && quarterOfDate(d, m.date) === q).map(m => m.id));
      d.calendar.marks = d.calendar.marks.filter(m => !removed.has(m.id));
      d.lessons.forEach(l => (l.items || []).forEach(it => (it.students || []).forEach(en => {
        if (en.markId && removed.has(en.markId)) en.markId = null;
      })));
      d.calendar.students.forEach(st => {
        if (idSet.has(st.id)) st.closedQ = Array.from(new Set([...(st.closedQ || []), q]));
      });
      return d;
    });
    setNewQuarterOpen(false);
    setJui({
      editing: false,
      draft: {}
    });
  };
  const reopenQuarter = q => {
    const idSet = new Set(students.map(st => st.id));
    update(d => {
      d.calendar.students.forEach(st => {
        if (idSet.has(st.id)) st.closedQ = (st.closedQ || []).filter(x => x !== q);
      });
      return d;
    });
  };
  const currentQ = quarterOfDate(data, todayIso());
  const defaultCloseQ = (() => {
    const ids = new Set(students.map(st => st.id));
    const has = q => data.calendar.marks.some(m => ids.has(m.studentId) && quarterOfDate(data, m.date) === q);
    for (let q = Math.min(currentQ - 1, qc - 1); q >= 0; q--) if (has(q)) return q;
    for (let q = qc - 1; q >= 0; q--) if (has(q)) return q;
    return Math.max(0, Math.min(currentQ, qc - 1));
  })();

  // ---- итоговые отметки ----
  const periodMark = (sid, kindP, q) => (data.calendar.periodMarks || []).find(p => p.studentId === sid && p.kind === kindP && (kindP !== "quarter" || p.q === q));
  const setPeriodMark = (sid, kindP, q, text) => {
    const {
      out,
      bad
    } = parseMarkTokens(text, gridScale);
    if (bad.length || out.length > 1 || out[0] === "Н") return false;
    update(d => {
      d.calendar.periodMarks = d.calendar.periodMarks || [];
      const i = d.calendar.periodMarks.findIndex(p => p.studentId === sid && p.kind === kindP && (kindP !== "quarter" || p.q === q));
      if (!out.length) {
        if (i >= 0) d.calendar.periodMarks.splice(i, 1);
      } else if (i >= 0) d.calendar.periodMarks[i].value = out[0];else d.calendar.periodMarks.push({
        id: uid(),
        studentId: sid,
        kind: kindP,
        q: kindP === "quarter" ? q : null,
        value: out[0]
      });
      return d;
    });
    return true;
  };
  const semesterMode = students.length > 0 && settingsForStudent(data, students[0]).periodMode === "semester";

  // Подробная правка одной отметки в окне (комментарий, причина «Н», «выставлена») — правый клик / «Подробнее…»
  const saveCellMark = (target, patch) => update(d => {
    let mark = target.markId ? d.calendar.marks.find(m => m.id === target.markId) : null;
    if (mark) {
      Object.assign(mark, patch);
      d.lessons.forEach(l => (l.items || []).forEach(it => itemStudents(it).forEach(entry => {
        if (entry.markId === mark.id) {
          entry.grade = String(patch.value);
          entry.gradeScale = patch.scale || "5";
          entry.gradeExcused = patch.value === "Н" ? patch.excused !== false : null;
        }
      })));
    } else {
      const l = lessonForDate(target.date);
      d.calendar.marks.push({
        id: uid(),
        studentId: target.studentId,
        date: target.date,
        lessonId: l?.id || null,
        exerciseId: null,
        questionId: null,
        posted: false,
        ...patch
      });
    }
    return d;
  });
  const navigateCell = (dir, patch) => {
    if (patch) saveCellMark(cellModal, patch);
    const dates = datesForQuarter(quarterOfDate(data, cellModal.date));
    const sIdx = students.findIndex(s => s.id === cellModal.studentId);
    let nS = sIdx,
      nD = dates.indexOf(cellModal.date);
    if (dir === "down") nS = Math.min(students.length - 1, sIdx + 1);
    if (dir === "up") nS = Math.max(0, sIdx - 1);
    if (dir === "right") nD = Math.min(dates.length - 1, nD + 1);
    if (dir === "left") nD = Math.max(0, nD - 1);
    const ns = students[nS],
      nd = dates[nD];
    if (!ns || !nd) return;
    setCellModal({
      studentId: ns.id,
      date: nd,
      markId: (marksOfCell(ns.id, nd)[0] || {}).id || null
    });
  };
  const quarterTables = [];
  for (let q = 0; q < qc; q++) {
    const idsHere = new Set(students.map(st => st.id));
    const closedHere = students.length > 0 && students.every(st => (st.closedQ || []).includes(q)) && !data.calendar.marks.some(m => idsHere.has(m.studentId) && quarterOfDate(data, m.date) === q);
    if (closedHere) {
      quarterTables.push( /*#__PURE__*/React.createElement("div", {
        key: q,
        style: {
          marginBottom: 20,
          minWidth: 0
        }
      }, /*#__PURE__*/React.createElement("div", {
        className: "row wrap",
        style: {
          fontFamily: "var(--font-display)",
          fontWeight: 700,
          fontSize: 15,
          marginBottom: 8,
          gap: 10
        }
      }, "\u0427\u0435\u0442\u0432\u0435\u0440\u0442\u044C ", q + 1, " \u2014 \u0437\u0430\u043A\u0440\u044B\u0442\u0430", /*#__PURE__*/React.createElement("span", {
        className: "hint",
        style: {
          fontWeight: 400
        }
      }, "\u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u0441\u043E\u0445\u0440\u0430\u043D\u0435\u043D\u044B \u0432 \u0444\u0430\u0439\u043B\u0435 Excel; \u043E\u0441\u0442\u0430\u043B\u0438\u0441\u044C \u0442\u043E\u043B\u044C\u043A\u043E \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u043D\u044B\u0435"), /*#__PURE__*/React.createElement(Btn, {
        kind: "ghost",
        sm: true,
        onClick: () => reopenQuarter(q),
        title: "\u0421\u043D\u043E\u0432\u0430 \u043F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u0442\u044C \u0442\u0430\u0431\u043B\u0438\u0446\u0443 \u0434\u0430\u0442 \u044D\u0442\u043E\u0439 \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u0438 (\u0435\u0441\u043B\u0438 \u0441\u0442\u0430\u0432\u0438\u0442\u0435 \u0432 \u043D\u0435\u0439 \u043E\u0442\u043C\u0435\u0442\u043A\u0438)"
      }, "\u041F\u043E\u043A\u0430\u0437\u0430\u0442\u044C \u0441\u043D\u043E\u0432\u0430")), /*#__PURE__*/React.createElement("div", {
        className: "jwrap"
      }, /*#__PURE__*/React.createElement("table", {
        className: "jtable"
      }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", null, "\u2116"), /*#__PURE__*/React.createElement("th", null, "\u0424.\u0418.\u041E."), /*#__PURE__*/React.createElement("th", null, semesterMode ? "Полугодовая" : "Четвертная"))), /*#__PURE__*/React.createElement("tbody", null, students.map((st, i) => /*#__PURE__*/React.createElement("tr", {
        key: st.id,
        style: {
          background: st.noCall ? "#f6f2fb" : undefined
        }
      }, /*#__PURE__*/React.createElement("td", null, i + 1), /*#__PURE__*/React.createElement("td", {
        className: "jname"
      }, st.name), /*#__PURE__*/React.createElement("td", null, /*#__PURE__*/React.createElement(PeriodMarkInput, {
        value: (periodMark(st.id, "quarter", q) || {}).value,
        placeholder: "",
        onCommit: t => setPeriodMark(st.id, "quarter", q, t)
      })))))))));
      continue;
    }
    const dates = datesForQuarter(q);
    if (!dates.length) continue;
    const monthGroups = [];
    dates.forEach(dt => {
      const label = shortMon(dt);
      const last = monthGroups[monthGroups.length - 1];
      if (last && last.label === label) last.count++;else monthGroups.push({
        label,
        count: 1
      });
    });
    quarterTables.push( /*#__PURE__*/React.createElement("div", {
      key: q,
      style: {
        marginBottom: 20,
        minWidth: 0
      }
    }, /*#__PURE__*/React.createElement("div", {
      style: {
        fontFamily: "var(--font-display)",
        fontWeight: 700,
        fontSize: 15,
        marginBottom: 8
      }
    }, "\u0427\u0435\u0442\u0432\u0435\u0440\u0442\u044C ", q + 1, gridScale === "100" ? " · 100-балльная шкала" : ""), /*#__PURE__*/React.createElement("div", {
      className: "jwrap"
    }, /*#__PURE__*/React.createElement("table", {
      className: "jtable"
    }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", {
      rowSpan: 3,
      className: "jsn"
    }, "\u2116"), /*#__PURE__*/React.createElement("th", {
      rowSpan: 3,
      className: "jsf"
    }, "\u0424.\u0418.\u041E."), monthGroups.map((g, i) => /*#__PURE__*/React.createElement("th", {
      key: i,
      colSpan: g.count
    }, g.label)), /*#__PURE__*/React.createElement("th", {
      rowSpan: 3
    }, "\u0421\u0440. \u0431\u0430\u043B\u043B"), /*#__PURE__*/React.createElement("th", {
      rowSpan: 3
    }, "\u041A\u043E\u043B-\u0432\u043E"), (!semesterMode || q % 2 === 1) && /*#__PURE__*/React.createElement("th", {
      rowSpan: 3,
      title: "\u0418\u0442\u043E\u0433\u043E\u0432\u0430\u044F \u043E\u0442\u043C\u0435\u0442\u043A\u0430 \u0437\u0430 \u043F\u0435\u0440\u0438\u043E\u0434. \u0421\u0435\u0440\u044B\u043C \u2014 \u043F\u043E\u0434\u0441\u043A\u0430\u0437\u043A\u0430 \u043F\u043E \u0441\u0440\u0435\u0434\u043D\u0435\u043C\u0443 \u0431\u0430\u043B\u043B\u0443"
    }, semesterMode ? "Полугодовая" : "Четвертная"), /*#__PURE__*/React.createElement("th", {
      rowSpan: 3,
      title: "\u0421\u0432\u0435\u0440\u043A\u0430 \u0441 \u0438\u043C\u043F\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u043D\u043D\u043E\u0439 \u0442\u0430\u0431\u043B\u0438\u0446\u0435\u0439 \u044D\u043B\u0435\u043A\u0442\u0440\u043E\u043D\u043D\u043E\u0433\u043E \u0436\u0443\u0440\u043D\u0430\u043B\u0430: \u043A\u0430\u043A\u0438\u0445 \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u0438\u0437 \u041A\u043E\u043D\u0441\u0442\u0440\u0443\u043A\u0442\u043E\u0440\u0430 \u0442\u0430\u043C \u0435\u0449\u0451 \u043D\u0435\u0442"
    }, "\u0412\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u043D\u044B\u0435 \u043E\u0442\u043C\u0435\u0442\u043A\u0438")), /*#__PURE__*/React.createElement("tr", null, dates.map(dt => /*#__PURE__*/React.createElement("th", {
      key: dt,
      className: isDateMarked(dt) ? "jposted" : ""
    }, weekdayShort(dt)))), /*#__PURE__*/React.createElement("tr", null, dates.map(dt => /*#__PURE__*/React.createElement("th", {
      key: dt,
      className: (isDateMarked(dt) ? "jposted" : "") + (hoverCell?.date === dt ? " jcol-hover" : ""),
      onMouseEnter: () => setHoverCell({
        date: dt
      }),
      onMouseLeave: () => setHoverCell(null)
    }, /*#__PURE__*/React.createElement("span", {
      className: "jdatenum",
      title: "\u041E\u0442\u0447\u0451\u0442 \u0437\u0430 \u044D\u0442\u043E\u0442 \u0434\u0435\u043D\u044C",
      onClick: () => setReportDate(dt)
    }, dayNum(dt)), /*#__PURE__*/React.createElement("span", {
      className: "jbulk",
      title: "\u041F\u0440\u0430\u0432\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u044D\u0442\u043E\u0433\u043E \u0434\u043D\u044F \u043F\u0440\u044F\u043C\u043E \u0432 \u0442\u0430\u0431\u043B\u0438\u0446\u0435",
      onClick: () => {
        setJui(p => ({
          editing: true,
          draft: p.draft || {}
        }));
        setTimeout(() => students[0] && focusCell(students[0].id, dt), 60);
      }
    }, "\u270F\uFE0F"), /*#__PURE__*/React.createElement("input", {
      type: "checkbox",
      className: "jdonecheck",
      checked: isDateMarked(dt),
      onChange: () => toggleDateMarked(dt),
      title: "\u041E\u0442\u043C\u0435\u0442\u043A\u0438 \u0432\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u044B \u2014 \u0437\u0430\u043A\u0440\u0430\u0441\u0438\u0442\u044C \u0441\u0442\u043E\u043B\u0431\u0435\u0446"
    }))))), /*#__PURE__*/React.createElement("tbody", null, students.map((s, idx) => {
      const marksByDate = {};
      studentMarksOf(data, s.id, gridScale).filter(m => quarterOfDate(data, m.date) === q).forEach(m => (marksByDate[m.date] = marksByDate[m.date] || []).push(m));
      const sSettings = settingsForStudent(data, s);
      const periodIdx = sSettings.periodMode === "semester" ? semesterOfQuarter(q) : q;
      const periodQuarters = quartersInPeriod(sSettings.periodMode, periodIdx, qc);
      const haveCount = periodQuarters.reduce((sum, qq) => sum + numericMarksOf(data, s.id, gridScale).filter(m => quarterOfDate(data, m.date) === qq).length, 0);
      const remLessons = groupLessons.filter(l => {
        if (isLessonStudied(l)) return false;
        const ld = l.plannedDate || l.actualDate;
        if (!ld || ld < tIso) return false;
        return periodQuarters.includes(quarterOfDate(data, ld));
      }).length;
      const urgency = urgencyColorFor(haveCount, sSettings.requiredMarks, remLessons);
      const rowHover = hoverCell?.studentId === s.id;
      const rowBg = rowHover ? "var(--accent-soft)" : s.noCall ? "#ece5f6" : "var(--surface)";
      return /*#__PURE__*/React.createElement("tr", {
        key: s.id,
        style: {
          background: rowHover ? "var(--accent-soft)" : s.noCall ? "#f6f2fb" : undefined
        }
      }, /*#__PURE__*/React.createElement("td", {
        className: "jsn",
        style: {
          background: rowBg,
          color: "var(--faint)"
        }
      }, idx + 1), /*#__PURE__*/React.createElement("td", {
        className: "jname jsf",
        style: {
          background: rowBg,
          color: urgency || undefined
        },
        title: (s.noCall ? "Не вызывается к доске (домашнее обучение). " : "") + (sSettings.requiredMarks ? `Отметок за период: ${haveCount} из ${sSettings.requiredMarks}, уроков впереди: ${remLessons}` : "")
      }, s.name), dates.map(dt => {
        const cellMarks = marksByDate[dt] || [];
        const allPosted = cellMarks.length > 0 && cellMarks.every(m => m.posted);
        const hoverHere = hoverCell?.studentId === s.id && hoverCell?.date === dt;
        const highlighted = hoverCell?.studentId === s.id || hoverCell?.date === dt;
        if (jui.editing) {
          const text = cellText(s.id, dt);
          const badText = parseMarkTokens(text, gridScale).bad.length > 0;
          const changed = draft[s.id + "|" + dt] !== undefined && draft[s.id + "|" + dt] !== existingText(s.id, dt);
          return /*#__PURE__*/React.createElement("td", {
            key: dt,
            className: "jcell" + (allPosted ? " jposted" : ""),
            style: {
              padding: 2
            }
          }, /*#__PURE__*/React.createElement("input", {
            "data-cell": s.id + "|" + dt,
            value: text,
            className: "jinput" + (badText ? " bad" : changed ? " changed" : ""),
            onChange: e => setJui(p => ({
              draft: {
                ...p.draft,
                [s.id + "|" + dt]: e.target.value
              }
            })),
            onFocus: e => e.target.select(),
            onKeyDown: e => {
              const i = students.findIndex(x => x.id === s.id);
              if (e.key === "Enter" || e.key === "ArrowDown") {
                e.preventDefault();
                if (students[i + 1]) focusCell(students[i + 1].id, dt);
              } else if (e.key === "ArrowUp") {
                e.preventDefault();
                if (students[i - 1]) focusCell(students[i - 1].id, dt);
              }
            }
          }));
        }
        return /*#__PURE__*/React.createElement("td", {
          key: dt,
          className: "jcell" + (allPosted ? " jposted" : highlighted ? " jcol-hover" : ""),
          onClick: () => startInlineEdit(s.id, dt, false),
          onMouseEnter: () => setHoverCell({
            studentId: s.id,
            date: dt
          }),
          onMouseLeave: () => setHoverCell(null),
          title: cellMarks.length ? undefined : formatDateRu(dt)
        }, cellMarks.length ? /*#__PURE__*/React.createElement("span", {
          style: {
            display: "inline-flex",
            gap: 3,
            flexWrap: "nowrap",
            justifyContent: "center",
            alignItems: "center",
            whiteSpace: "nowrap"
          }
        }, cellMarks.map(m => /*#__PURE__*/React.createElement("span", {
          key: m.id,
          className: "jpill",
          onClick: e => {
            e.stopPropagation();
            startInlineEdit(s.id, dt, false);
          },
          onContextMenu: e => {
            e.preventDefault();
            setCellModal({
              studentId: s.id,
              date: dt,
              markId: m.id
            });
          },
          onMouseEnter: e => showMarkHover(e, m.id),
          onMouseLeave: hideMarkHover,
          style: {
            opacity: m.posted ? 0.72 : 1,
            background: m.value === "Н" ? m.excused ? "#8593a3" : "#c0392b" : markColorFor(m.value, m.scale)
          }
        }, m.value)), hoverHere && /*#__PURE__*/React.createElement("span", {
          onClick: e => {
            e.stopPropagation();
            startInlineEdit(s.id, dt, true);
          },
          style: {
            color: "#3b82c4",
            fontWeight: 900,
            fontSize: 18,
            padding: "0 5px",
            cursor: "pointer",
            lineHeight: "26px"
          },
          title: "\u0414\u043E\u043F\u0438\u0441\u0430\u0442\u044C \u0435\u0449\u0451 \u043E\u0434\u043D\u0443 \u043E\u0442\u043C\u0435\u0442\u043A\u0443 \u0432 \u044D\u0442\u043E\u0439 \u043A\u043B\u0435\u0442\u043A\u0435"
        }, "+")) : /*#__PURE__*/React.createElement("span", {
          className: "jempty"
        }, "\xB7"));
      }), /*#__PURE__*/React.createElement("td", {
        style: {
          fontWeight: 800
        }
      }, fmtAvgNum(quarterAvgOf(data, s.id, q, gridScale))), /*#__PURE__*/React.createElement("td", {
        style: {
          color: "var(--muted)"
        }
      }, numericMarksOf(data, s.id, gridScale).filter(m => quarterOfDate(data, m.date) === q).length), (!semesterMode || q % 2 === 1) && /*#__PURE__*/React.createElement("td", null, /*#__PURE__*/React.createElement(PeriodMarkInput, {
        value: (periodMark(s.id, "quarter", q) || {}).value,
        placeholder: (() => {
          const a = quarterAvgOf(data, s.id, q, gridScale);
          return a == null ? "" : String(Math.round(a));
        })(),
        onCommit: t => setPeriodMark(s.id, "quarter", q, t)
      })), (() => {
        const mf = gridScale === "5" ? missingInFile(data, s, q) : null;
        if (!mf) return /*#__PURE__*/React.createElement("td", {
          style: {
            color: "#c3ceda"
          },
          title: "\u0422\u0430\u0431\u043B\u0438\u0446\u0430 \u044D\u043B\u0435\u043A\u0442\u0440\u043E\u043D\u043D\u043E\u0433\u043E \u0436\u0443\u0440\u043D\u0430\u043B\u0430 \u0434\u043B\u044F \u044D\u0442\u043E\u0439 \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u0438 \u043D\u0435 \u0438\u043C\u043F\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u043B\u0430\u0441\u044C"
        }, "\u2014");
        const when = mf.importedAt ? new Date(mf.importedAt).toLocaleDateString("ru-RU") : "";
        return mf.missing.length ? /*#__PURE__*/React.createElement("td", {
          style: {
            color: "#c0392b",
            fontWeight: 700
          },
          title: `Сверка с таблицей от ${when}${mf.fileName ? " (" + mf.fileName + ")" : ""}`
        }, missingText(mf.missing)) : /*#__PURE__*/React.createElement("td", {
          style: {
            color: "#2e8b57",
            fontWeight: 800,
            fontSize: 16
          },
          title: `Все отметки выставлены (сверка с таблицей от ${when})`
        }, "\u2713");
      })());
    }))))));
  }
  const cellStudent = cellModal ? students.find(s => s.id === cellModal.studentId) : null;
  const cellExisting = cellModal && cellModal.markId ? data.calendar.marks.find(m => m.id === cellModal.markId) : undefined;
  const cellLesson = cellModal ? lessonForDate(cellModal.date) : null;
  return /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10
    }
  }, tracks.map(t => /*#__PURE__*/React.createElement(TrackChip, {
    key: trackKeyOf(t.kind, t.value),
    label: t.label,
    color: t.color,
    active: effectiveKey === trackKeyOf(t.kind, t.value),
    onClick: () => setTrackKey(trackKeyOf(t.kind, t.value))
  }))), letters.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10,
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041B\u0438\u0442\u0435\u0440\u0430:"), letters.map(cl => /*#__PURE__*/React.createElement(TrackChip, {
    key: cl.id,
    label: `${value}${cl.name}`,
    color: "#8593a3",
    active: letterFilter === cl.id,
    onClick: () => setLetterFilter(cl.id)
  })), (letterFilter === "none" || hasNoLetterItems(data, value)) && /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0411\u0435\u0437 \u043B\u0438\u0442\u0435\u0440\u044B",
    color: "#8593a3",
    active: letterFilter === "none",
    onClick: () => setLetterFilter("none")
  }), /*#__PURE__*/React.createElement(TrackChip, {
    label: "\u0412\u0441\u0435",
    color: "#8593a3",
    active: letterFilter === "all",
    onClick: () => setLetterFilter("all")
  })), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, "\u0427\u0435\u0442\u0432\u0435\u0440\u0442\u0435\u0439: ", /*#__PURE__*/React.createElement("b", {
    style: {
      color: "var(--ink)"
    }
  }, qc), " ", data.calendar.vacations.length ? "(по датам каникул)" : "(добавьте каникулы в календаре, чтобы разделить год на четверти)", " · шкала: ", /*#__PURE__*/React.createElement("b", {
    style: {
      color: "var(--ink)"
    }
  }, GRADING_SYSTEM_LABELS[settingsHere.gradingSystem]), settingsHere.requiredMarks && /*#__PURE__*/React.createElement(React.Fragment, null, " \xB7 \u043D\u043E\u0440\u043C\u0430 \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u0437\u0430 \u043F\u0435\u0440\u0438\u043E\u0434: ", /*#__PURE__*/React.createElement("b", {
    style: {
      color: "var(--ink)"
    }
  }, settingsHere.requiredMarks)), " · список учеников — во вкладке «Списки классов»."), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10,
      gap: 8,
      alignItems: "center"
    }
  }, students.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setAddMarkOpen(true)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u041E\u0442\u043C\u0435\u0442\u043A\u0430 \u0431\u0435\u0437 \u0443\u0440\u043E\u043A\u0430"), kind === "class" && /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    sm: true,
    onClick: () => setXlsxOpen(true),
    title: "\u0417\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u0438\u0437 \u0442\u0430\u0431\u043B\u0438\u0446 \u0448\u043A\u043E\u043B\u044C\u043D\u043E\u0433\u043E \u044D\u043B\u0435\u043A\u0442\u0440\u043E\u043D\u043D\u043E\u0433\u043E \u0436\u0443\u0440\u043D\u0430\u043B\u0430 (.xlsx): \u0441\u0432\u0435\u0440\u0438\u0442 \u043A\u043E\u043B\u0438\u0447\u0435\u0441\u0442\u0432\u043E \u043E\u0446\u0435\u043D\u043E\u043A, \u0434\u043E\u0431\u0430\u0432\u0438\u0442 \u043D\u0435\u0434\u043E\u0441\u0442\u0430\u044E\u0449\u0438\u0435"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 12
  }), " \u0418\u043C\u043F\u043E\u0440\u0442 \u0438\u0437 \u044D\u043B. \u0436\u0443\u0440\u043D\u0430\u043B\u0430 (.xlsx)"), students.length > 0 && !!(window.litAPI && window.litAPI.autoTypeQueue) && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setAutoTypeOpen(true),
    title: "\u041F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u0430\u043C\u0430 \u043D\u0430\u0431\u0435\u0440\u0451\u0442 \u043D\u0435\u0432\u043D\u0435\u0441\u0451\u043D\u043D\u044B\u0435 \u043E\u0446\u0435\u043D\u043A\u0438 \u0432 \u0448\u043A\u043E\u043B\u044C\u043D\u044B\u0439 \u044D\u043B\u0435\u043A\u0442\u0440\u043E\u043D\u043D\u044B\u0439 \u0436\u0443\u0440\u043D\u0430\u043B (\xAB\u041C\u043E\u044F \u0448\u043A\u043E\u043B\u0430\xBB) \u0441\u0442\u043E\u043B\u0431\u0446\u043E\u043C \u043F\u043E \u0434\u0430\u0442\u0430\u043C"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u2328 \u0410\u0432\u0442\u043E\u043D\u0430\u0431\u043E\u0440 \u0432 \u0436\u0443\u0440\u043D\u0430\u043B"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !jui.canUndo,
    onClick: jui.undo,
    title: "\u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C \u043F\u043E\u0441\u043B\u0435\u0434\u043D\u0435\u0435 \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u0435 \u043E\u0442\u043C\u0435\u0442\u043E\u043A (Ctrl+Z)"
  }, "\u21B6 \u041E\u0442\u043C\u0435\u043D\u0438\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: !jui.canRedo,
    onClick: jui.redo,
    title: "\u0412\u0435\u0440\u043D\u0443\u0442\u044C \u043E\u0442\u043C\u0435\u043D\u0451\u043D\u043D\u043E\u0435 (Ctrl+Y)"
  }, "\u21B7 \u0412\u0435\u0440\u043D\u0443\u0442\u044C"), /*#__PURE__*/React.createElement(ToggleSwitch, {
    checked: !!jui.hideEmpty,
    onChange: v => setJui({
      hideEmpty: v
    }),
    label: "\u0421\u043A\u0440\u044B\u0442\u044C \u043F\u0443\u0441\u0442\u044B\u0435 \u0434\u0430\u0442\u044B"
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: students.length === 0,
    onClick: () => setNewQuarterOpen(true),
    title: "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u0438 \u0432 Excel \u0438 \u043E\u0447\u0438\u0441\u0442\u0438\u0442\u044C \u0436\u0443\u0440\u043D\u0430\u043B, \u043E\u0441\u0442\u0430\u0432\u0438\u0432 \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u043D\u0443\u044E \u043E\u0442\u043C\u0435\u0442\u043A\u0443"
  }, "\uD83E\uDDF9 \u041D\u0430\u0447\u0430\u0442\u044C \u043D\u043E\u0432\u0443\u044E \u0447\u0435\u0442\u0432\u0435\u0440\u0442\u044C"), !jui.editing ? /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    disabled: students.length === 0,
    onClick: () => setJui({
      editing: true,
      draft: {}
    }),
    title: "\u0412\u0432\u043E\u0434\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u043F\u0440\u044F\u043C\u043E \u0432 \u0442\u0430\u0431\u043B\u0438\u0446\u0435; \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u043E \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u0432 \u043A\u043B\u0435\u0442\u043A\u0435 \u2014 \u0447\u0435\u0440\u0435\u0437 \u043F\u0440\u043E\u0431\u0435\u043B; \u043F\u043E\u0442\u043E\u043C \u2713"
  }, "\u270F \u041F\u0440\u0430\u0432\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0438") : /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: commitDraft
  }, "\u2713 \u0417\u0430\u043F\u0438\u0441\u0430\u0442\u044C", changedCount ? ` (изменено клеток: ${changedCount})` : ""), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setJui({
      editing: false,
      draft: {}
    })
  }, "\u2715 \u041E\u0442\u043C\u0435\u043D\u0430"))), jui.editing && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 12,
      color: "#2f6aa3",
      background: "#eaf2fa",
      borderRadius: 8,
      padding: "6px 10px",
      marginBottom: 10
    }
  }, "\u0420\u0435\u0436\u0438\u043C \u043F\u0440\u0430\u0432\u043A\u0438 (\u043A\u043B\u0438\u043A \u043F\u043E \u043B\u044E\u0431\u043E\u0439 \u043A\u043B\u0435\u0442\u043A\u0435 \u0438\u043B\u0438 \u043E\u0442\u043C\u0435\u0442\u043A\u0435 \u0432\u043A\u043B\u044E\u0447\u0430\u0435\u0442 \u0435\u0433\u043E): \u043F\u0438\u0448\u0438\u0442\u0435 \u043E\u0442\u043C\u0435\u0442\u043A\u0438 \u0432 \u043A\u043B\u0435\u0442\u043A\u0430\u0445 (5, 4, \u041D; \u043D\u0435\u0441\u043A\u043E\u043B\u044C\u043A\u043E \u0432 \u043E\u0434\u043D\u043E\u0439 \u043A\u043B\u0435\u0442\u043A\u0435 \u2014 \xAB5 4\xBB). Enter / \u2193 \u2014 \u043A \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0435\u043C\u0443 \u0443\u0447\u0435\u043D\u0438\u043A\u0443, Tab \u2014 \u043A \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0435\u0439 \u0434\u0430\u0442\u0435. \u041F\u043E\u043A\u0430 \u043D\u0435 \u043D\u0430\u0436\u0430\u0442\u0430 \u2713 \xAB\u0417\u0430\u043F\u0438\u0441\u0430\u0442\u044C\xBB, \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u043C\u0435\u043D\u044F\u0435\u0442\u0441\u044F."), students.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u0421\u043F\u0438\u0441\u043E\u043A \u043F\u0443\u0441\u0442 \u2014 \u0434\u043E\u0431\u0430\u0432\u044C\u0442\u0435 \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432 \u044D\u0442\u043E\u0433\u043E \u043A\u043B\u0430\u0441\u0441\u0430/\u043A\u0443\u0440\u0441\u0430 \u0432\u043E \u0432\u043A\u043B\u0430\u0434\u043A\u0435 \xAB\u0421\u043F\u0438\u0441\u043A\u0438 \u043A\u043B\u0430\u0441\u0441\u043E\u0432\xBB.") : /*#__PURE__*/React.createElement(React.Fragment, null, quarterTables.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041F\u043E\u043A\u0430 \u043D\u0435\u0442 \u043D\u0438 \u0443\u0440\u043E\u043A\u043E\u0432, \u043D\u0438 \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u2014 \u0442\u0430\u0431\u043B\u0438\u0446\u0430 \u043F\u043E\u044F\u0432\u0438\u0442\u0441\u044F, \u043A\u0430\u043A \u0442\u043E\u043B\u044C\u043A\u043E \u0447\u0442\u043E-\u0442\u043E \u0438\u0437 \u044D\u0442\u043E\u0433\u043E \u0431\u0443\u0434\u0435\u0442.") : quarterTables, /*#__PURE__*/React.createElement("div", {
    style: {
      minWidth: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "var(--font-display)",
      fontWeight: 700,
      fontSize: 15,
      marginBottom: 8
    }
  }, "\u0418\u0442\u043E\u0433\u0438 \u0433\u043E\u0434\u0430"), /*#__PURE__*/React.createElement("div", {
    className: "jwrap"
  }, /*#__PURE__*/React.createElement("table", {
    className: "jtable"
  }, /*#__PURE__*/React.createElement("thead", null, /*#__PURE__*/React.createElement("tr", null, /*#__PURE__*/React.createElement("th", null, "\u2116"), /*#__PURE__*/React.createElement("th", null, "\u0424.\u0418.\u041E."), Array.from({
    length: qc
  }).map((_, qi) => /*#__PURE__*/React.createElement("th", {
    key: qi
  }, semesterMode ? `${Math.floor(qi / 2) + 1} пол.` : `${qi + 1} чет.`)), /*#__PURE__*/React.createElement("th", null, "\u0413\u043E\u0434\u043E\u0432\u0430\u044F"), /*#__PURE__*/React.createElement("th", null, "\u042D\u043A\u0437\u0430\u043C\u0435\u043D"), /*#__PURE__*/React.createElement("th", null, "\u0418\u0442\u043E\u0433\u043E\u0432\u0430\u044F"))), /*#__PURE__*/React.createElement("tbody", null, students.map((s, idx) => {
    const qVals = Array.from({
      length: qc
    }).map((_, qi) => (periodMark(s.id, "quarter", qi) || {}).value);
    const nums = qVals.filter(v => typeof v === "number");
    const suggest = nums.length ? String(Math.round(nums.reduce((a, b) => a + b, 0) / nums.length)) : "";
    return /*#__PURE__*/React.createElement("tr", {
      key: s.id,
      style: {
        background: s.noCall ? "#f6f2fb" : undefined
      }
    }, /*#__PURE__*/React.createElement("td", {
      style: {
        color: "var(--faint)"
      }
    }, idx + 1), /*#__PURE__*/React.createElement("td", {
      className: "jname"
    }, s.name), qVals.map((v, qi) => /*#__PURE__*/React.createElement("td", {
      key: qi,
      style: {
        fontWeight: 800,
        color: v == null ? "#c3ceda" : undefined
      }
    }, v == null ? "·" : v)), /*#__PURE__*/React.createElement("td", null, /*#__PURE__*/React.createElement(PeriodMarkInput, {
      value: (periodMark(s.id, "year") || {}).value,
      placeholder: suggest,
      onCommit: t => setPeriodMark(s.id, "year", null, t)
    })), /*#__PURE__*/React.createElement("td", null, /*#__PURE__*/React.createElement(PeriodMarkInput, {
      value: (periodMark(s.id, "exam") || {}).value,
      placeholder: "",
      onCommit: t => setPeriodMark(s.id, "exam", null, t)
    })), /*#__PURE__*/React.createElement("td", null, /*#__PURE__*/React.createElement(PeriodMarkInput, {
      value: (periodMark(s.id, "final") || {}).value,
      placeholder: "",
      onCommit: t => setPeriodMark(s.id, "final", null, t)
    })));
  })))))), hoverMark && (() => {
    const hm = data.calendar.marks.find(m => m.id === hoverMark.id);
    if (!hm) return null;
    const ls = hm.lessonId ? data.lessons.find(l => l.id === hm.lessonId) : null;
    return /*#__PURE__*/React.createElement("div", {
      onMouseEnter: () => clearTimeout(hoverTimer.current),
      onMouseLeave: hideMarkHover,
      style: {
        position: "fixed",
        left: Math.max(6, Math.min(hoverMark.x, window.innerWidth - 200)),
        top: hoverMark.y,
        transform: hoverMark.below ? "none" : "translateY(-100%)",
        zIndex: 1000,
        background: "#fff",
        border: "1.5px solid var(--line)",
        borderRadius: 9,
        padding: "6px 10px",
        boxShadow: "0 6px 18px rgba(19,44,83,.22)",
        maxWidth: 190
      }
    }, /*#__PURE__*/React.createElement(ToggleSwitch, {
      checked: !!hm.posted,
      onChange: v => setPosted([hm.id], v),
      label: "\u0412\u044B\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u0430",
      small: true
    }), /*#__PURE__*/React.createElement("div", {
      onClick: () => {
        setHoverMark(null);
        setCellModal({
          studentId: hm.studentId,
          date: hm.date,
          markId: hm.id
        });
      },
      style: {
        fontSize: 11,
        color: "#3b82c4",
        marginTop: 4,
        cursor: "pointer",
        textDecoration: "underline"
      },
      title: "\u041A\u043E\u043C\u043C\u0435\u043D\u0442\u0430\u0440\u0438\u0439, \u043F\u0440\u0438\u0447\u0438\u043D\u0430 \xAB\u041D\xBB \u0438 \u0442.\u043F. (\u043D\u0430 \u043F\u043B\u0430\u043D\u0448\u0435\u0442\u0435 \u2014 \u0434\u043E\u043B\u0433\u043E\u0435 \u043D\u0430\u0436\u0430\u0442\u0438\u0435 \u043D\u0430 \u043E\u0442\u043C\u0435\u0442\u043A\u0443)"
    }, "\u041F\u043E\u0434\u0440\u043E\u0431\u043D\u0435\u0435\u2026"), /*#__PURE__*/React.createElement("div", {
      className: "hint",
      style: {
        marginTop: 3,
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis"
      }
    }, formatDateRu(hm.date), ls && ls.topic ? " · " + ls.topic : ""));
  })(), cellModal && cellStudent && /*#__PURE__*/React.createElement(MarkCellModal, {
    key: `${cellModal.studentId}-${cellModal.date}-${cellModal.markId || "new"}`,
    student: cellStudent,
    date: cellModal.date,
    lesson: cellLesson,
    existingMark: cellExisting,
    scale: gridScale,
    onSave: patch => {
      saveCellMark(cellModal, patch);
      setCellModal(null);
    },
    onNavigate: navigateCell,
    onClear: () => {
      update(d => {
        const removedIds = cellModal.markId ? [cellModal.markId] : [];
        d.calendar.marks = d.calendar.marks.filter(m => !removedIds.includes(m.id));
        d.lessons.forEach(l => (l.items || []).forEach(it => {
          if (Array.isArray(it.students)) it.students = it.students.filter(entry => !removedIds.includes(entry.markId));
        }));
        return d;
      });
      setCellModal(null);
    },
    onClose: () => setCellModal(null)
  }), newQuarterOpen && /*#__PURE__*/React.createElement(NewQuarterModal, {
    data: data,
    students: students,
    groupLabel: groupLabel,
    qc: qc,
    defaultQ: defaultCloseQ,
    onClose: () => setNewQuarterOpen(false),
    onConfirm: startNewQuarter
  }), xlsxOpen && /*#__PURE__*/React.createElement(JournalXlsxImportModal, {
    data: data,
    update: update,
    onClose: () => setXlsxOpen(false)
  }), autoTypeOpen && /*#__PURE__*/React.createElement(EntryQueueModal, {
    data: data,
    update: update,
    focus: {
      kind,
      value
    },
    onClose: () => setAutoTypeOpen(false)
  }), reportDate && /*#__PURE__*/React.createElement(JournalDateReportModal, {
    data: data,
    students: students,
    date: reportDate,
    onClose: () => setReportDate(null),
    groupLabel: groupLabel
  }), addMarkOpen && /*#__PURE__*/React.createElement(AddJournalMarkModal, {
    students: students,
    gridScale: gridScale,
    onClose: () => setAddMarkOpen(false),
    onSave: patch => {
      update(d => {
        d.calendar.marks.push({
          id: uid(),
          lessonId: null,
          exerciseId: null,
          questionId: null,
          posted: false,
          ...patch
        });
        return d;
      });
      setAddMarkOpen(false);
    }
  }));
}
function SeatingChart({
  data,
  update,
  kind,
  value,
  classLetterId,
  groupLabel,
  students,
  deleteStudent
}) {
  const [addingAt, setAddingAt] = useState(null);
  const [addingName, setAddingName] = useState("");
  const [renamingId, setRenamingId] = useState(null);
  const [renamingName, setRenamingName] = useState("");
  const [visionStudentId, setVisionStudentId] = useState("");
  const [visionMaxRow, setVisionMaxRow] = useState(1);
  const [avoidDraft, setAvoidDraft] = useState([]);
  const importRef = useRef(null);
  const existingLayout = getSeatingLayout(data, kind, value, classLetterId);
  const layout = existingLayout ? JSON.parse(JSON.stringify(existingLayout)) : {
    files: [[], [], []],
    visionLimits: [],
    seatFurther: [],
    avoidGroups: []
  };
  ensureSeatingCapacity(layout);
  const nameById = id => students.find(s => s.id === id)?.name || "";
  const seatedIds = new Set();
  layout.files.forEach(file => file.forEach(desk => desk.seats.forEach(s => s && seatedIds.add(s))));
  const unseated = students.filter(s => !seatedIds.has(s.id));
  const visionMapNow = new Map((layout.visionLimits || []).map(v => [v.studentId, v.maxRow]));
  const avoidIds = new Set();
  (layout.avoidGroups || []).forEach(g => g.members.forEach(m => avoidIds.add(m)));
  // Цвет места: красноватый — ограничение по зрению, жёлтый — нежелательное соседство (оба — пополам).
  const tintOf = sid => {
    const v = visionMapNow.has(sid),
      a = avoidIds.has(sid);
    if (v && a) return {
      background: "linear-gradient(135deg, #f9cfc9 50%, #fff0a8 50%)",
      border: "1.5px solid #e09a90"
    };
    if (v) return {
      background: "#f9cfc9",
      border: "1.5px solid #e09a90"
    };
    if (a) return {
      background: "#fff0a8",
      border: "1.5px solid #e0c24a"
    };
    return null;
  };
  const seatBlocked = (sid, di) => visionMapNow.has(sid) && di + 1 > visionMapNow.get(sid);
  const blockedMsg = sid => `${nameById(sid)} — ограничение по зрению: не дальше ${visionMapNow.get(sid)} парты от доски. Посадить дальше нельзя.`;
  const withLayout = fn => update(d => {
    const l = getOrCreateSeatingLayout(d, kind, value, classLetterId);
    fn(l, d);
    ensureSeatingCapacity(l);
    return d;
  });
  const moveStudentToSeat = (studentId, fileIdx, deskIdx, seatIdx) => {
    // физически нельзя посадить ученика с ограничением по зрению дальше указанной парты (и вытесняемого тоже)
    if (seatBlocked(studentId, deskIdx)) {
      showToast(blockedMsg(studentId));
      return;
    }
    {
      let src = null;
      layout.files.forEach((file, fi) => file.forEach((desk, di) => desk.seats.forEach((x, si) => {
        if (x === studentId) src = {
          fi,
          di,
          si
        };
      })));
      const target = layout.files[fileIdx] && layout.files[fileIdx][deskIdx];
      const disp = target ? target.seats[seatIdx] : null;
      if (disp && disp !== studentId && visionMapNow.has(disp)) {
        const fitsSource = src && !seatBlocked(disp, src.di);
        const otherFree = layout.files.some(file => file.some((desk, di) => !seatBlocked(disp, di) && desk.seats.some(x => !x)));
        if (!fitsSource && !otherFree) {
          showToast(blockedMsg(disp));
          return;
        }
      }
    }
    withLayout(l => {
      let sourcePos = null;
      l.files.forEach((file, fi) => file.forEach((desk, di) => desk.seats.forEach((s, si) => {
        if (s === studentId) sourcePos = {
          fi,
          di,
          si
        };
      })));
      const desk = l.files[fileIdx][deskIdx];
      if (!desk) return;
      const displaced = desk.seats[seatIdx];
      desk.seats[seatIdx] = studentId;
      if (sourcePos) {
        const same = sourcePos.fi === fileIdx && sourcePos.di === deskIdx && sourcePos.si === seatIdx;
        if (!same) {
          l.files[sourcePos.fi][sourcePos.di].seats[sourcePos.si] = null;
          if (displaced && displaced !== studentId) {
            // вытесненного сажаем на освободившееся место, если оно допустимо, иначе — на первое подходящее
            const lim = (l.visionLimits || []).find(v => v.studentId === displaced);
            if (!lim || sourcePos.di + 1 <= lim.maxRow) l.files[sourcePos.fi][sourcePos.di].seats[sourcePos.si] = displaced;else placeInFirstEmptySeat(l, displaced, sourcePos.fi);
          }
        }
      } else if (displaced && displaced !== studentId) placeInFirstEmptySeat(l, displaced, fileIdx);
    });
  };
  const handleDrop = (e, fi, di, si) => {
    e.preventDefault();
    const sid = e.dataTransfer.getData("text/plain");
    if (sid) moveStudentToSeat(sid, fi, di, si);
  };
  const seatUnseated = studentId => {
    const free = layout.files.some(file => file.some((desk, di) => !seatBlocked(studentId, di) && desk.seats.some(x => !x)));
    if (!free && visionMapNow.has(studentId)) {
      showToast(blockedMsg(studentId).replace("Посадить дальше нельзя.", "Свободных мест ближе нет — освободите место на первых партах."));
      return;
    }
    withLayout(l => placeInFirstEmptySeat(l, studentId, 0));
  };
  const addStudentAtSeat = (fi, di, si, name) => {
    if (!name.trim()) return;
    const newId = uid();
    update(d => {
      d.calendar.students.push({
        id: newId,
        name: name.trim(),
        class: kind === "class" ? value : null,
        courseId: kind === "course" ? value : null,
        classLetterId: kind === "class" && classLetterId ? classLetterId : null,
        courseSubgroupId: kind === "course" && classLetterId ? classLetterId : null
      });
      const l = getOrCreateSeatingLayout(d, kind, value, classLetterId);
      const desk = l.files[fi][di];
      if (desk) desk.seats[si] = newId;
      ensureSeatingCapacity(l);
      return d;
    });
    setAddingAt(null);
    setAddingName("");
  };
  const renameStudent = (id, name) => update(d => {
    const s = d.calendar.students.find(x => x.id === id);
    if (s) s.name = name.trim() || "Ученик";
    return d;
  });
  const addVisionLimit = () => {
    if (!visionStudentId) return;
    const sid = visionStudentId,
      row = Math.max(1, +visionMaxRow || 1);
    withLayout(l => {
      l.visionLimits = (l.visionLimits || []).filter(v => v.studentId !== sid);
      l.visionLimits.push({
        studentId: sid,
        maxRow: row
      });
      // если ученик сидит дальше разрешённого — пересаживаем на допустимое место
      let pos = null;
      l.files.forEach((file, fi) => file.forEach((desk, di) => desk.seats.forEach((x, si) => {
        if (x === sid) pos = {
          fi,
          di,
          si
        };
      })));
      if (pos && pos.di + 1 > row) {
        l.files[pos.fi][pos.di].seats[pos.si] = null;
        if (!placeInFirstEmptySeat(l, sid, pos.fi)) l.files[pos.fi][pos.di].seats[pos.si] = null;
      }
    });
    setVisionStudentId("");
    setVisionMaxRow(1);
  };
  const removeVisionLimit = sid => withLayout(l => {
    l.visionLimits = (l.visionLimits || []).filter(v => v.studentId !== sid);
  });
  const addAvoidGroup = () => {
    if (avoidDraft.length < 2) return;
    const members = [...avoidDraft];
    withLayout(l => {
      l.avoidGroups = l.avoidGroups || [];
      l.avoidGroups.push({
        id: uid(),
        members
      });
    });
    setAvoidDraft([]);
  };
  const removeAvoidGroup = gid => withLayout(l => {
    l.avoidGroups = (l.avoidGroups || []).filter(g => g.id !== gid);
  });
  const randomize = async () => {
    if (students.length === 0) return;
    if (!(await Confirm("Рассадить учеников случайным образом? Текущая рассадка будет заменена."))) return;
    // расстановка считается заранее (по текущим ограничениям), затем записывается в проект
    const allIds = students.map(st => st.id);
    const desksPerFile = Math.max(1, Math.ceil(allIds.length / 3 / 2));
    const visionMap = new Map((layout.visionLimits || []).map(v => [v.studentId, v.maxRow]));
    const groups = (layout.avoidGroups || []).map(g => ({
      ...g,
      members: g.members.filter(m => allIds.includes(m))
    })).filter(g => g.members.length >= 2);
    const attempt = () => {
      const files = [0, 1, 2].map(() => Array.from({
        length: desksPerFile
      }, () => ({
        id: uid(),
        seats: [null, null]
      })));
      const failed = [];
      const free = () => {
        const arr = [];
        files.forEach((file, fi) => file.forEach((desk, di) => desk.seats.forEach((x, si) => {
          if (!x) arr.push({
            fi,
            di,
            si
          });
        })));
        return arr;
      };
      shuffleArray(allIds.filter(id => visionMap.has(id))).forEach(id => {
        const cand = free().filter(p => p.di + 1 <= visionMap.get(id));
        const pick = cand[Math.floor(Math.random() * cand.length)];
        if (pick) files[pick.fi][pick.di].seats[pick.si] = id;else failed.push(id);
      });
      shuffleArray(allIds.filter(id => !visionMap.has(id))).forEach(id => {
        const f = free();
        const pick = f[Math.floor(Math.random() * f.length)];
        if (pick) files[pick.fi][pick.di].seats[pick.si] = id;
      });
      return {
        files,
        failed,
        bad: countNearInGroups(files, groups)
      };
    };
    let best = null;
    for (let i = 0; i < 300; i++) {
      const t = attempt();
      if (!best || t.failed.length < best.failed.length || t.failed.length === best.failed.length && t.bad < best.bad) best = t;
      if (!t.failed.length && !t.bad) break;
    }
    withLayout(l => {
      l.files = best.files;
    });
    if (best.failed.length) showToast("Не хватило мест достаточно близко к доске для: " + best.failed.map(nameById).join(", ") + ". Они остались в списке «Не рассажены».");else if (best.bad) showToast("Удалось рассадить, но не все «нежелательные» соседи разведены (" + best.bad + ") — они выделены жёлтым.");
  };
  const exportLayout = () => {
    const big = `<div style="page-break-after:always;">${buildSeatingLayoutHtml(layout, nameById, false)}</div>`;
    const small = Array.from({
      length: 6
    }, () => buildSeatingLayoutHtml(layout, nameById, true)).join("");
    downloadWord(`Рассадка ${groupLabel}.doc`, `Рассадка — ${groupLabel}`, big + small);
  };
  const exportJson = () => {
    const nameOf = id => id ? nameById(id) : null;
    const obj = {
      type: "literature-constructor-seating",
      version: 1,
      exportedAt: new Date().toISOString(),
      groupLabel,
      files: layout.files.map((file, fi) => ({
        label: SEATING_FILE_LABELS[fi],
        desks: file.map(desk => ({
          seats: desk.seats.map(nameOf)
        }))
      })),
      visionLimits: (layout.visionLimits || []).map(v => ({
        name: nameById(v.studentId),
        maxRow: v.maxRow
      })).filter(v => v.name),
      seatFurther: (layout.seatFurther || []).map(nameById).filter(Boolean),
      avoidGroups: (layout.avoidGroups || []).map(g => g.members.map(nameById).filter(Boolean)).filter(g => g.length >= 2)
    };
    const blob = new Blob([JSON.stringify(obj, null, 2)], {
      type: "application/json"
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `Рассадка ${groupLabel}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };
  const importJson = file => {
    const reader = new FileReader();
    reader.onload = async () => {
      let obj;
      try {
        obj = JSON.parse(reader.result);
      } catch (e) {
        showToast("Не удалось прочитать файл — это не корректный JSON.");
        return;
      }
      if (!obj || obj.type !== "literature-constructor-seating" || !Array.isArray(obj.files)) {
        showToast("Файл не похож на экспортированный список рассадки");
        return;
      }
      if (!(await Confirm("Импортировать рассадку? Текущая рассадка для этой литеры/подгруппы будет заменена."))) return;
      const desksOf = src => Array.isArray(src) ? src : src && Array.isArray(src.desks) ? src.desks : [];
      update(d => {
        const findOrCreate = name => {
          if (!name) return null;
          const trimmed = String(name).trim();
          if (!trimmed) return null;
          let s = d.calendar.students.find(x => x.name.trim().toLowerCase() === trimmed.toLowerCase() && (kind === "class" ? x.class === value && !x.courseId && (x.classLetterId || null) === (classLetterId || null) : x.courseId === value && (x.courseSubgroupId || null) === (classLetterId || null)));
          if (!s) {
            s = {
              id: uid(),
              name: trimmed,
              class: kind === "class" ? value : null,
              courseId: kind === "course" ? value : null,
              classLetterId: kind === "class" && classLetterId ? classLetterId : null,
              courseSubgroupId: kind === "course" && classLetterId ? classLetterId : null
            };
            d.calendar.students.push(s);
          }
          return s.id;
        };
        const l = getOrCreateSeatingLayout(d, kind, value, classLetterId);
        l.files = [0, 1, 2].map(fi => desksOf(obj.files[fi]).map(desk => ({
          id: uid(),
          seats: [findOrCreate(desk.seats?.[0]), findOrCreate(desk.seats?.[1])]
        })));
        l.visionLimits = (obj.visionLimits || []).map(v => ({
          studentId: findOrCreate(v.name),
          maxRow: Math.max(1, +v.maxRow || 1)
        })).filter(v => v.studentId);
        l.seatFurther = (obj.seatFurther || []).map(findOrCreate).filter(Boolean);
        l.avoidGroups = [...(obj.avoidGroups || []).map(g => ({
          id: uid(),
          members: g.map(findOrCreate).filter(Boolean)
        })), ...(obj.avoidPairs || []).map(p => ({
          id: uid(),
          members: [findOrCreate(p.a), findOrCreate(p.b)].filter(Boolean)
        }))].filter(g => g.members.length >= 2);
        ensureSeatingCapacity(l);
        return d;
      });
      showToast("Рассадка импортирована");
    };
    reader.readAsText(file, "utf-8");
  };
  const visionCandidates = students.filter(s => !(layout.visionLimits || []).some(v => v.studentId === s.id));
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "panel",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      marginBottom: 12
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontWeight: 700,
      marginBottom: 6
    }
  }, "\uD83D\uDC41 \u041E\u0433\u0440\u0430\u043D\u0438\u0447\u0435\u043D\u0438\u044F \u043F\u043E \u0437\u0440\u0435\u043D\u0438\u044E (\u043D\u0435 \u0434\u0430\u043B\u0435\u0435 \u0443\u043A\u0430\u0437\u0430\u043D\u043D\u043E\u0439 \u043F\u0430\u0440\u0442\u044B \u043E\u0442 \u0434\u043E\u0441\u043A\u0438)"), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 6
    }
  }, (layout.visionLimits || []).map(v => /*#__PURE__*/React.createElement("span", {
    key: v.studentId,
    className: "chip"
  }, nameById(v.studentId), " \xB7 \u043D\u0435 \u0434\u0430\u043B\u0435\u0435 ", v.maxRow, " \u043F\u0430\u0440\u0442\u044B", /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => removeVisionLimit(v.studentId)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 10
  }))))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: visionStudentId,
    onChange: e => setVisionStudentId(e.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0412\u044B\u0431\u0440\u0430\u0442\u044C \u0443\u0447\u0435\u043D\u0438\u043A\u0430\u2026"), visionCandidates.map(s => /*#__PURE__*/React.createElement("option", {
    key: s.id,
    value: s.id
  }, s.name))), /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u043D\u0435 \u0434\u0430\u043B\u0435\u0435 \u043F\u0430\u0440\u0442\u044B \u2116"), /*#__PURE__*/React.createElement("input", {
    type: "number",
    min: "1",
    value: visionMaxRow,
    onChange: e => setVisionMaxRow(e.target.value),
    style: {
      width: 60,
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "5px 8px"
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: addVisionLimit
  }, "\u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C"))), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      fontWeight: 700,
      marginBottom: 6
    }
  }, "\uD83D\uDFE8 \u041D\u0435\u0436\u0435\u043B\u0430\u0442\u0435\u043B\u044C\u043D\u043E \u0441\u0430\u0436\u0430\u0442\u044C \u0440\u044F\u0434\u043E\u043C (\u0433\u0440\u0443\u043F\u043F\u044B: \u043D\u0430 \u043E\u0434\u043D\u043E\u0439 \u043F\u0430\u0440\u0442\u0435, \u043D\u0430 \u0441\u043E\u0441\u0435\u0434\u043D\u0438\u0445 \u0438 \u0431\u043B\u0438\u0437\u043A\u0438\u0445). \u041F\u043E\u0441\u0430\u0434\u0438\u0442\u044C \u0440\u044F\u0434\u043E\u043C \u043C\u043E\u0436\u043D\u043E \u2014 \u0442\u0430\u043A\u0438\u0435 \u0443\u0447\u0435\u043D\u0438\u043A\u0438 \u043F\u0440\u043E\u0441\u0442\u043E \u0432\u044B\u0434\u0435\u043B\u0435\u043D\u044B \u0436\u0451\u043B\u0442\u044B\u043C"), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 6
    }
  }, (layout.avoidGroups || []).map(g => /*#__PURE__*/React.createElement("span", {
    key: g.id,
    className: "chip"
  }, g.members.map(nameById).filter(Boolean).join(", "), /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => removeAvoidGroup(g.id)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 10
  }))))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap"
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041D\u043E\u0432\u0430\u044F \u0433\u0440\u0443\u043F\u043F\u0430:"), avoidDraft.map(id => /*#__PURE__*/React.createElement("span", {
    key: id,
    className: "chip"
  }, nameById(id), /*#__PURE__*/React.createElement("span", {
    className: "x",
    onClick: () => setAvoidDraft(avoidDraft.filter(x => x !== id))
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 10
  })))), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: "",
    onChange: e => e.target.value && setAvoidDraft([...avoidDraft, e.target.value])
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, avoidDraft.length ? "Ещё ученик…" : "Выбрать ученика…"), students.filter(st => !avoidDraft.includes(st.id)).map(st => /*#__PURE__*/React.createElement("option", {
    key: st.id,
    value: st.id
  }, st.name))), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: addAvoidGroup,
    disabled: avoidDraft.length < 2
  }, "\u0421\u043E\u0437\u0434\u0430\u0442\u044C \u0433\u0440\u0443\u043F\u043F\u0443"))))), unseated.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "unseated-strip"
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041D\u0435 \u0440\u0430\u0441\u0441\u0430\u0436\u0435\u043D\u044B (\u043D\u0430\u0436\u043C\u0438\u0442\u0435, \u0447\u0442\u043E\u0431\u044B \u043F\u043E\u0441\u0430\u0434\u0438\u0442\u044C \u043D\u0430 \u0441\u0432\u043E\u0431\u043E\u0434\u043D\u043E\u0435 \u043C\u0435\u0441\u0442\u043E, \u0438\u043B\u0438 \u043F\u0435\u0440\u0435\u0442\u0430\u0449\u0438\u0442\u0435 \u043D\u0430 \u043F\u0430\u0440\u0442\u0443):"), unseated.map(s => /*#__PURE__*/React.createElement("span", {
    key: s.id,
    className: "unseated-chip",
    style: tintOf(s.id) || undefined,
    draggable: true,
    onDragStart: e => e.dataTransfer.setData("text/plain", s.id),
    onClick: () => seatUnseated(s.id)
  }, s.name))), /*#__PURE__*/React.createElement("div", {
    className: "panel",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "seatboard"
  }, "\u0414\u043E\u0441\u043A\u0430"), /*#__PURE__*/React.createElement("div", {
    className: "seatcols"
  }, layout.files.map((file, fi) => /*#__PURE__*/React.createElement("div", {
    key: fi,
    className: "seatcol"
  }, /*#__PURE__*/React.createElement("div", {
    className: "seatrow-label",
    style: {
      textAlign: "center"
    }
  }, SEATING_FILE_LABELS[fi]), file.map((desk, di) => /*#__PURE__*/React.createElement("div", {
    key: desk.id,
    className: "deskrow"
  }, /*#__PURE__*/React.createElement("span", {
    className: "deskno"
  }, di + 1), /*#__PURE__*/React.createElement("div", {
    className: "deskseats"
  }, desk.seats.map((sid, si) => {
    const isAdding = addingAt && addingAt.fi === fi && addingAt.di === di && addingAt.si === si;
    const isRenaming = sid && renamingId === sid;
    if (isAdding) {
      return /*#__PURE__*/React.createElement("input", {
        key: si,
        autoFocus: true,
        className: "seatinput",
        value: addingName,
        placeholder: "\u0418\u043C\u044F",
        onChange: e => setAddingName(e.target.value),
        onKeyDown: e => {
          if (e.key === "Enter") addStudentAtSeat(fi, di, si, addingName);
          if (e.key === "Escape") {
            setAddingAt(null);
            setAddingName("");
          }
        },
        onBlur: () => addingName.trim() ? addStudentAtSeat(fi, di, si, addingName) : (setAddingAt(null), setAddingName(""))
      });
    }
    if (isRenaming) {
      return /*#__PURE__*/React.createElement("input", {
        key: si,
        autoFocus: true,
        className: "seatinput",
        value: renamingName,
        onChange: e => setRenamingName(e.target.value),
        onKeyDown: e => {
          if (e.key === "Enter") {
            renameStudent(sid, renamingName);
            setRenamingId(null);
          }
          if (e.key === "Escape") setRenamingId(null);
        },
        onBlur: () => {
          if (renamingName.trim()) renameStudent(sid, renamingName);
          setRenamingId(null);
        }
      });
    }
    return /*#__PURE__*/React.createElement("div", {
      key: si,
      className: "seat" + (sid ? " filled" : ""),
      style: sid ? tintOf(sid) || undefined : undefined,
      draggable: !!sid,
      onDragStart: e => sid && e.dataTransfer.setData("text/plain", sid),
      onDragOver: e => e.preventDefault(),
      onDrop: e => handleDrop(e, fi, di, si),
      onClick: () => {
        if (!sid) {
          setAddingAt({
            fi,
            di,
            si
          });
          setAddingName("");
        }
      },
      onDoubleClick: () => {
        if (sid) {
          setRenamingId(sid);
          setRenamingName(nameById(sid));
        }
      },
      title: sid ? "Перетащите, чтобы пересадить · двойной клик — переименовать" : "Нажмите, чтобы добавить ученика"
    }, sid ? /*#__PURE__*/React.createElement(React.Fragment, null, /*#__PURE__*/React.createElement("span", {
      className: "seatname"
    }, nameById(sid)), /*#__PURE__*/React.createElement("span", {
      className: "x",
      onClick: e => {
        e.stopPropagation();
        deleteStudent(sid);
      },
      title: "\u0423\u0434\u0430\u043B\u0438\u0442\u044C \u0443\u0447\u0435\u043D\u0438\u043A\u0430 (\u0441\u043E \u0432\u0441\u0435\u043C\u0438 \u043E\u0442\u043C\u0435\u0442\u043A\u0430\u043C\u0438)"
    }, "\xD7")) : "+");
  }))))))))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    sm: true,
    onClick: randomize
  }, "\uD83C\uDFB2 \u041F\u0440\u043E\u0438\u0437\u0432\u043E\u043B\u044C\u043D\u0430\u044F \u0440\u0430\u0441\u0441\u0430\u0434\u043A\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: exportLayout
  }, "\uD83D\uDDA8 \u042D\u043A\u0441\u043F\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u043C\u0430\u043A\u0435\u0442"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: exportJson
  }, "\u2B07 \u042D\u043A\u0441\u043F\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0441\u043F\u0438\u0441\u043E\u043A"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => importRef.current?.click()
  }, "\u2B06 \u0418\u043C\u043F\u043E\u0440\u0442\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0441\u043F\u0438\u0441\u043E\u043A"), /*#__PURE__*/React.createElement("input", {
    ref: importRef,
    type: "file",
    accept: "application/json,.json",
    style: {
      display: "none"
    },
    onChange: e => {
      if (e.target.files[0]) importJson(e.target.files[0]);
      e.target.value = "";
    }
  })));
}
function ClassListsSection({
  data,
  update,
  trackSel,
  setTrackSel,
  subSel,
  setSubSel,
  listView,
  setListView
}) {
  const [newSubName, setNewSubName] = useState("");
  const [newStudentName, setNewStudentName] = useState("");
  const isClass = trackSel.kind === "class";
  const subs = isClass ? lettersForGrade(data, trackSel.value) : subgroupsForCourse(data, trackSel.value);
  const subIdForSettings = subSel === "none" ? null : subSel;
  const noLetterHere = subs.length === 0 || (isClass ? hasNoLetterItems(data, trackSel.value) : data.calendar.students.some(st => st.courseId === trackSel.value && !st.courseSubgroupId));
  // «Без литеры» показывается, только если такие ученики/уроки есть; иначе открываем первую литеру
  useEffect(() => {
    if (subSel === "none" && !noLetterHere && subs.length) setSubSel(subs[0].id);
  }, [trackSel.kind, trackSel.value, subSel, noLetterHere, subs.length]);
  const addSub = () => {
    if (!newSubName.trim()) return;
    const id = uid();
    update(d => {
      if (isClass) d.classLetters.push({
        id,
        grade: trackSel.value,
        name: newSubName.trim()
      });else d.courseSubgroups.push({
        id,
        courseId: trackSel.value,
        name: newSubName.trim()
      });
      return d;
    });
    setNewSubName("");
  };
  const deleteSub = async id => {
    if (!(await Confirm(isClass ? "Удалить эту литеру? У уроков и учеников, где она была указана, литера просто снимется." : "Удалить эту подгруппу? У учеников, где она была указана, подгруппа просто снимется."))) return;
    update(d => {
      if (isClass) {
        d.classLetters = d.classLetters.filter(cl => cl.id !== id);
        d.lessons.forEach(l => {
          if (l.classLetterId === id) l.classLetterId = null;
        });
        d.calendar.students.forEach(s => {
          if (s.classLetterId === id) s.classLetterId = null;
        });
      } else {
        d.courseSubgroups = d.courseSubgroups.filter(g => g.id !== id);
        d.calendar.students.forEach(s => {
          if (s.courseSubgroupId === id) s.courseSubgroupId = null;
        });
      }
      return d;
    });
    if (subSel === id) setSubSel("none");
  };
  const studentsHere = data.calendar.students.filter(s => isClass ? s.class === trackSel.value && !s.courseId && (subSel === "none" ? !s.classLetterId : s.classLetterId === subSel) : s.courseId === trackSel.value && (subSel === "none" ? !s.courseSubgroupId : s.courseSubgroupId === subSel));
  const addStudent = () => {
    if (!newStudentName.trim()) return;
    update(d => {
      d.calendar.students.push({
        id: uid(),
        name: newStudentName.trim(),
        class: isClass ? trackSel.value : null,
        courseId: isClass ? null : trackSel.value,
        classLetterId: isClass ? subSel === "none" ? null : subSel : null,
        courseSubgroupId: isClass ? null : subSel === "none" ? null : subSel
      });
      return d;
    });
    setNewStudentName("");
  };
  const deleteStudent = async id => {
    const s = studentsHere.find(x => x.id === id);
    if (!s || !(await Confirm(`Удалить ученика «${s.name}» и все его отметки?`))) return;
    update(d => {
      removeStudentCascade(d, id);
      return d;
    });
  };
  // «Не вызывать» (домашнее обучение): остаётся в списке и журнале, но убирается из рассадки и рекомендаций «кого вызвать».
  const toggleNoCall = id => update(d => {
    const st = d.calendar.students.find(x => x.id === id);
    if (!st) return d;
    st.noCall = !st.noCall;
    if (st.noCall) removeStudentFromSeating(d, id);
    return d;
  });
  const renameStudent = (id, name) => update(d => {
    const s = d.calendar.students.find(x => x.id === id);
    if (s) s.name = name.trim() || "Ученик";
    return d;
  });
  const settings = getClassSettings(data, trackSel.kind, trackSel.value, subIdForSettings);
  const patchClassSettings = patch => update(d => {
    d.classSettings = d.classSettings || [];
    const key = classSettingsKey(trackSel.kind, trackSel.value, subIdForSettings);
    let rec = d.classSettings.find(s => classSettingsKey(s.kind, s.value, s.classLetterId) === key);
    if (!rec) {
      rec = {
        id: uid(),
        kind: trackSel.kind,
        value: trackSel.value,
        classLetterId: subIdForSettings,
        gradingSystem: "5",
        periodMode: "quarter",
        requiredMarks: null
      };
      d.classSettings.push(rec);
    }
    Object.assign(rec, patch);
    return d;
  });
  const target = isClass ? nextGrade(trackSel.value) : null;
  const promoteClass = async () => {
    if (!target) return;
    if (studentsHere.length === 0) {
      showToast("В этом классе пока нет учеников для перевода.");
      return;
    }
    const letterName = subSel === "none" ? null : classLetterName(data, subSel);
    const label = subSel === "none" ? `${trackSel.value} класс` : `${trackSel.value}${letterName}`;
    if (!(await Confirm(`Перевести «${label}» (${studentsHere.length} уч.) в ${target} класс? Список учеников переедет в новый класс, а уже прошедшие уроки и отметки останутся как есть — они относятся к ${trackSel.value} классу.`))) return;
    const existingTargetLetter = letterName ? lettersForGrade(data, target).find(cl => cl.name === letterName) : null;
    const targetLetterId = letterName ? existingTargetLetter?.id || uid() : null;
    const shouldCreateLetter = letterName && !existingTargetLetter;
    const movingIds = new Set(studentsHere.map(s => s.id));
    update(d => {
      if (shouldCreateLetter) d.classLetters.push({
        id: targetLetterId,
        grade: target,
        name: letterName
      });
      d.calendar.students.forEach(s => {
        if (movingIds.has(s.id)) {
          s.class = target;
          s.classLetterId = targetLetterId;
        }
      });
      return d;
    });
    setTrackSel({
      kind: "class",
      value: target
    });
    setSubSel(targetLetterId || "none");
  };
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10
    }
  }, allTracks(data).map(t => /*#__PURE__*/React.createElement(TrackChip, {
    key: trackKeyOf(t.kind, t.value),
    label: t.label,
    color: t.color,
    active: trackSel.kind === t.kind && trackSel.value === t.value,
    onClick: () => {
      setTrackSel({
        kind: t.kind,
        value: t.value
      });
      setSubSel("none");
    }
  }))), /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, "\u0417\u0434\u0435\u0441\u044C \u043C\u043E\u0436\u043D\u043E \u0432\u0435\u0441\u0442\u0438 \u0441\u043F\u0438\u0441\u043A\u0438 \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432 \u043F\u043E \u043A\u043B\u0430\u0441\u0441\u0430\u043C \u0438 \u043A\u0443\u0440\u0441\u0430\u043C \u043E\u0442\u0434\u0435\u043B\u044C\u043D\u043E \u043E\u0442 \u0436\u0443\u0440\u043D\u0430\u043B\u0430: \u0441\u043E\u0437\u0434\u0430\u0432\u0430\u0442\u044C \u043B\u0438\u0442\u0435\u0440\u044B/\u043F\u043E\u0434\u0433\u0440\u0443\u043F\u043F\u044B, \u0434\u043E\u0431\u0430\u0432\u043B\u044F\u0442\u044C \u0438 \u0443\u0434\u0430\u043B\u044F\u0442\u044C \u0443\u0447\u0435\u043D\u0438\u043A\u043E\u0432, \u043F\u0435\u0440\u0435\u0432\u043E\u0434\u0438\u0442\u044C \u043A\u043B\u0430\u0441\u0441 \u043D\u0430 \u0441\u043B\u0435\u0434\u0443\u044E\u0449\u0438\u0439 \u0433\u043E\u0434."), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 10
    }
  }, (subs.length === 0 || (isClass ? hasNoLetterItems(data, trackSel.value) : data.calendar.students.some(st => st.courseId === trackSel.value && !st.courseSubgroupId))) && /*#__PURE__*/React.createElement(TrackChip, {
    label: isClass ? "Без литеры" : "Без подгруппы",
    active: subSel === "none",
    color: "#8593a3",
    onClick: () => setSubSel("none")
  }), subs.map(s => /*#__PURE__*/React.createElement(TrackChip, {
    key: s.id,
    label: isClass ? `${trackSel.value}${s.name}` : s.name,
    active: subSel === s.id,
    color: "#8593a3",
    onClick: () => setSubSel(s.id)
  }))), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: newSubName,
    onChange: e => setNewSubName(e.target.value),
    placeholder: isClass ? "Буква литеры (А, Б…)" : "Название подгруппы",
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      maxWidth: 180
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: addSub
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " ", isClass ? "литеру" : "подгруппу"), subSel !== "none" && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => deleteSub(subSel)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 12
  }), " \u0423\u0434\u0430\u043B\u0438\u0442\u044C ", isClass ? "литеру" : "подгруппу")), /*#__PURE__*/React.createElement("div", {
    className: "panel",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-head"
  }, /*#__PURE__*/React.createElement("h2", null, "\u0421\u0438\u0441\u0442\u0435\u043C\u0430 \u043E\u0446\u0435\u043D\u0438\u0432\u0430\u043D\u0438\u044F")), /*#__PURE__*/React.createElement("div", {
    className: "panel-body"
  }, /*#__PURE__*/React.createElement("div", {
    className: "row wrap"
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u0428\u043A\u0430\u043B\u0430"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: settings.gradingSystem,
    onChange: e => patchClassSettings({
      gradingSystem: e.target.value
    })
  }, Object.entries(GRADING_SYSTEM_LABELS).map(([k, l]) => /*#__PURE__*/React.createElement("option", {
    key: k,
    value: k
  }, l)))), /*#__PURE__*/React.createElement(Field, {
    label: "\u041F\u0435\u0440\u0438\u043E\u0434"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: settings.periodMode,
    onChange: e => patchClassSettings({
      periodMode: e.target.value
    })
  }, /*#__PURE__*/React.createElement("option", {
    value: "quarter"
  }, "\u0427\u0435\u0442\u0432\u0435\u0440\u0442\u0438"), /*#__PURE__*/React.createElement("option", {
    value: "semester"
  }, "\u041F\u043E\u043B\u0443\u0433\u043E\u0434\u0438\u044F"))), /*#__PURE__*/React.createElement(Field, {
    label: "\u041D\u043E\u0440\u043C\u0430 \u043E\u0442\u043C\u0435\u0442\u043E\u043A \u0437\u0430 \u043F\u0435\u0440\u0438\u043E\u0434"
  }, /*#__PURE__*/React.createElement("input", {
    type: "number",
    value: settings.requiredMarks || "",
    onChange: e => patchClassSettings({
      requiredMarks: e.target.value ? +e.target.value : null
    }),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      width: 80
    }
  }))))), /*#__PURE__*/React.createElement("div", {
    className: "viewtoggle",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": listView === "alphabet" ? "true" : undefined,
    onClick: () => setListView("alphabet")
  }, "\u041F\u043E \u0430\u043B\u0444\u0430\u0432\u0438\u0442\u0443"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": listView !== "alphabet" ? "true" : undefined,
    onClick: () => setListView("seating")
  }, "\u041F\u043E \u0440\u0430\u0441\u0441\u0430\u0434\u043A\u0435")), /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("input", {
    value: newStudentName,
    onChange: e => setNewStudentName(e.target.value),
    placeholder: "\u0418\u043C\u044F \u0443\u0447\u0435\u043D\u0438\u043A\u0430",
    onKeyDown: e => e.key === "Enter" && addStudent(),
    style: {
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "7px 9px",
      maxWidth: 240
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    sm: true,
    onClick: addStudent
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 12
  }), " \u0414\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0443\u0447\u0435\u043D\u0438\u043A\u0430")), studentsHere.length === 0 && /*#__PURE__*/React.createElement("div", {
    className: "hint",
    style: {
      marginBottom: 10
    }
  }, "\u0421\u043F\u0438\u0441\u043E\u043A \u043F\u0443\u0441\u0442 \u2014 \u0434\u043E\u0431\u0430\u0432\u044C\u0442\u0435 \u0443\u0447\u0435\u043D\u0438\u043A\u0430 \u0432\u044B\u0448\u0435, \u043A\u043B\u0438\u043A\u043D\u0438\u0442\u0435 \u043F\u043E \u043F\u0430\u0440\u0442\u0435 (\u0432 \u0440\u0435\u0436\u0438\u043C\u0435 \xAB\u041F\u043E \u0440\u0430\u0441\u0441\u0430\u0434\u043A\u0435\xBB) \u0438\u043B\u0438 \u0438\u043C\u043F\u043E\u0440\u0442\u0438\u0440\u0443\u0439\u0442\u0435 \u0440\u0430\u0441\u0441\u0430\u0434\u043A\u0443 \u0438\u0437 \u0444\u0430\u0439\u043B\u0430."), listView === "alphabet" ? studentsHere.length > 0 && /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gap: 8
    }
  }, studentsHere.slice().sort((a, b) => a.name.localeCompare(b.name, "ru")).map((s, i) => /*#__PURE__*/React.createElement("div", {
    key: s.id,
    style: {
      background: s.noCall ? "#f3eefa" : "var(--surface)",
      border: "1px solid var(--line)",
      borderRadius: "var(--radius)",
      padding: "8px 12px",
      display: "flex",
      alignItems: "center",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint",
    style: {
      width: 20,
      flex: "none"
    }
  }, i + 1, "."), /*#__PURE__*/React.createElement("input", {
    value: s.name,
    onChange: e => renameStudent(s.id, e.target.value),
    style: {
      flex: 1,
      border: "1px solid transparent",
      background: "transparent",
      borderRadius: 7,
      padding: "5px 7px",
      fontWeight: 600,
      fontSize: 14
    }
  }), /*#__PURE__*/React.createElement("label", {
    title: "\u0414\u043E\u043C\u0430\u0448\u043D\u0435\u0435 \u043E\u0431\u0443\u0447\u0435\u043D\u0438\u0435: \u043D\u0435 \u0432\u044B\u0437\u044B\u0432\u0430\u0442\u044C \u043A \u0434\u043E\u0441\u043A\u0435, \u043D\u0435 \u0441\u0430\u0436\u0430\u0442\u044C \u0432 \u0440\u0430\u0441\u0441\u0430\u0434\u043A\u0435",
    style: {
      display: "flex",
      alignItems: "center",
      gap: 4,
      fontSize: 12,
      color: "#6b5a8a",
      cursor: "pointer",
      flex: "none"
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: !!s.noCall,
    onChange: () => toggleNoCall(s.id),
    style: {
      margin: 0
    }
  }), "\u043D\u0435 \u0432\u044B\u0437\u044B\u0432\u0430\u0442\u044C"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => deleteStudent(s.id)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 12
  }))))) : /*#__PURE__*/React.createElement(SeatingChart, {
    data: data,
    update: update,
    kind: trackSel.kind,
    value: trackSel.value,
    classLetterId: subIdForSettings,
    groupLabel: isClass ? `${trackSel.value}${subSel === "none" ? "" : classLetterName(data, subSel)} класс` : `${data.courses.find(c => c.id === trackSel.value)?.name || "Курс"}${subSel === "none" ? "" : " · " + (subgroupsForCourse(data, trackSel.value).find(g => g.id === subSel)?.name || "")}`,
    students: studentsHere.filter(st => !st.noCall),
    deleteStudent: deleteStudent,
    renameStudent: renameStudent
  }), isClass && /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 16
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    onClick: promoteClass,
    disabled: !target
  }, target ? `Перевести в ${target} класс →` : "11 класс — переводить некуда")));
}
function PlanningTab({
  data,
  update,
  onOpenLesson,
  nav,
  setNav,
  jui,
  setJui
}) {
  const p = nav.planning;
  const setSubTab = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      subTab: v
    }
  }));
  const setJournalTrackKey = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      journalTrackKey: v
    }
  }));
  const setJournalLetterFilter = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      journalLetterFilter: v
    }
  }));
  const setTrackSel = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      clTrackSel: v
    }
  }));
  const setSubSel = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      clSubSel: v
    }
  }));
  const setListView = v => setNav(n => ({
    ...n,
    planning: {
      ...n.planning,
      clListView: v
    }
  }));
  const tracks = allTracks(data);
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "viewtoggle",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": p.subTab === "journal" ? "true" : undefined,
    onClick: () => setSubTab("journal")
  }, "\u0416\u0443\u0440\u043D\u0430\u043B"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": p.subTab === "calendar" ? "true" : undefined,
    onClick: () => setSubTab("calendar")
  }, "\u041A\u0430\u043B\u0435\u043D\u0434\u0430\u0440\u044C"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": p.subTab === "classlists" ? "true" : undefined,
    onClick: () => setSubTab("classlists")
  }, "\u0421\u043F\u0438\u0441\u043A\u0438 \u043A\u043B\u0430\u0441\u0441\u043E\u0432")), p.subTab === "journal" && /*#__PURE__*/React.createElement(JournalSection, {
    data: data,
    update: update,
    trackKey: p.journalTrackKey,
    setTrackKey: setJournalTrackKey,
    tracks: tracks,
    letterFilter: p.journalLetterFilter,
    setLetterFilter: setJournalLetterFilter,
    jui: jui,
    setJui: setJui
  }), p.subTab === "calendar" && /*#__PURE__*/React.createElement(CalendarTab, {
    data: data,
    update: update,
    onOpenLesson: onOpenLesson,
    nav: nav,
    setNav: setNav
  }), p.subTab === "classlists" && /*#__PURE__*/React.createElement(ClassListsSection, {
    data: data,
    update: update,
    trackSel: p.clTrackSel,
    setTrackSel: setTrackSel,
    subSel: p.clSubSel,
    setSubSel: setSubSel,
    listView: p.clListView || "alphabet",
    setListView: setListView
  }));
}
;
// ============================================================
// materials.js — Материалы для статей: собранные отрывки со всего
// приложения, с поиском по тегам/автору/произведению.
// ============================================================

function MaterialCard({
  data,
  m,
  onOpen
}) {
  const author = m.authorId ? data.personalities.find(p => p.id === m.authorId) : null;
  const work = m.workId ? data.works.find(w => w.id === m.workId) : null;
  const tags = (m.tagIds || []).map(id => findTagAnywhere(data, id)).filter(Boolean);
  return /*#__PURE__*/React.createElement("div", {
    className: "card",
    onClick: onOpen
  }, /*#__PURE__*/React.createElement("div", {
    className: "ct-title"
  }, m.articleName || "Без названия статьи"), /*#__PURE__*/React.createElement("div", {
    className: "ct-preview"
  }, stripHtml(m.excerptHtml).slice(0, 140)), /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, m.sourceSection, m.sourceEntityTitle ? ` · ${m.sourceEntityTitle}` : ""), (author || work) && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, author?.title, author && work ? " — " : "", work?.title), tags.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "ct-tags"
  }, tags.slice(0, 3).map(t => /*#__PURE__*/React.createElement(Badge, {
    key: t.id,
    tag: t
  }))));
}
function patchMaterial(update, id, patch) {
  update(d => {
    const idx = d.materialsForArticles.findIndex(x => x.id === id);
    if (idx >= 0) d.materialsForArticles[idx] = {
      ...d.materialsForArticles[idx],
      ...patch
    };
    return d;
  });
}
function MaterialModal({
  data,
  update,
  id,
  onClose
}) {
  const m = data.materialsForArticles.find(x => x.id === id);
  const [articleName, setArticleName] = useState(m?.articleName || "");
  const [note, setNote] = useState(m?.note || "");
  if (!m) return null;
  const remove = async () => {
    if (!(await Confirm("Удалить материал?"))) return;
    update(d => {
      d.materialsForArticles = d.materialsForArticles.filter(x => x.id !== id);
      return d;
    });
    onClose();
  };
  const author = m.authorId ? data.personalities.find(p => p.id === m.authorId) : null;
  const work = m.workId ? data.works.find(w => w.id === m.workId) : null;
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B \u0434\u043B\u044F \u0441\u0442\u0430\u0442\u044C\u0438"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u0441\u0442\u0430\u0442\u044C\u0438"
  }, /*#__PURE__*/React.createElement("input", {
    value: articleName,
    onChange: e => setArticleName(e.target.value),
    onBlur: () => patchMaterial(update, id, {
      articleName
    }),
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u041E\u0442\u0440\u044B\u0432\u043E\u043A"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel"
  }, /*#__PURE__*/React.createElement("div", {
    className: "panel-body",
    dangerouslySetInnerHTML: {
      __html: m.excerptHtml
    }
  }))), /*#__PURE__*/React.createElement(Field, {
    label: "\u0417\u0430\u043C\u0435\u0442\u043A\u0430"
  }, /*#__PURE__*/React.createElement("textarea", {
    value: note,
    onChange: e => setNote(e.target.value),
    onBlur: () => patchMaterial(update, id, {
      note
    }),
    rows: 3,
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: 8
    }
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u0418\u0441\u0442\u043E\u0447\u043D\u0438\u043A"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, m.sourceSection, m.sourceEntityTitle ? ` · ${m.sourceEntityTitle}` : "", author ? ` · автор: ${author.title}` : "", work ? ` · произведение: ${work.title}` : "")), /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.genre.title
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "genre",
    selectedIds: nsTagIds(m.tagIds, "genre", data),
    onChange: ids => patchMaterial(update, id, {
      tagIds: withNsTagIds(m.tagIds, "genre", data, ids)
    })
  })), /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.literaryGenre.title
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "literaryGenre",
    selectedIds: nsTagIds(m.tagIds, "literaryGenre", data),
    onChange: ids => patchMaterial(update, id, {
      tagIds: withNsTagIds(m.tagIds, "literaryGenre", data, ids)
    })
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "danger",
    onClick: remove
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  }), " \u0423\u0434\u0430\u043B\u0438\u0442\u044C"))));
}
function MaterialCreateModal({
  data,
  update,
  projectPath,
  onClose
}) {
  const [articleName, setArticleName] = useState("");
  const [excerptHtml, setExcerptHtml] = useState("");
  const [note, setNote] = useState("");
  const [authorId, setAuthorId] = useState("");
  const [workId, setWorkId] = useState("");
  const [tagIds, setTagIds] = useState([]);
  const save = () => {
    if (!stripHtml(excerptHtml).trim() && !articleName.trim()) {
      showToast("Впишите название статьи или текст отрывка");
      return;
    }
    update(d => {
      d.materialsForArticles.push({
        id: uid(),
        excerptHtml,
        note,
        articleName: articleName.trim(),
        tagIds: [...tagIds],
        sourceSection: "Материалы для статей",
        sourceEntityId: null,
        sourceEntityTitle: "",
        authorId: authorId || null,
        workId: workId || null,
        createdAt: nowISO()
      });
      return d;
    });
    onClose();
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u041D\u043E\u0432\u044B\u0439 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B \u0434\u043B\u044F \u0441\u0442\u0430\u0442\u044C\u0438"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body"
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u041D\u0430\u0437\u0432\u0430\u043D\u0438\u0435 \u0441\u0442\u0430\u0442\u044C\u0438"
  }, /*#__PURE__*/React.createElement("input", {
    value: articleName,
    onChange: e => setArticleName(e.target.value),
    placeholder: "\u041D\u0430\u043F\u0440\u0438\u043C\u0435\u0440: \xAB\u041C\u043E\u0442\u0438\u0432 \u0434\u043E\u0440\u043E\u0433\u0438 \u0443 \u041F\u0443\u0448\u043A\u0438\u043D\u0430\xBB",
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: "8px 10px"
    }
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u0422\u0435\u043A\u0441\u0442 \u043E\u0442\u0440\u044B\u0432\u043A\u0430"
  }, /*#__PURE__*/React.createElement(RichEditor, {
    html: excerptHtml,
    onChange: setExcerptHtml,
    projectPath: projectPath,
    compact: true
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u0417\u0430\u043C\u0435\u0442\u043A\u0430"
  }, /*#__PURE__*/React.createElement("textarea", {
    value: note,
    onChange: e => setNote(e.target.value),
    rows: 2,
    style: {
      width: "100%",
      border: "1px solid var(--line)",
      borderRadius: 8,
      padding: 8
    }
  })), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 18
    }
  }, /*#__PURE__*/React.createElement(Field, {
    label: "\u0410\u0432\u0442\u043E\u0440 (\u043D\u0435\u043E\u0431\u044F\u0437\u0430\u0442\u0435\u043B\u044C\u043D\u043E)"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: authorId,
    onChange: e => setAuthorId(e.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u041D\u0435 \u0443\u043A\u0430\u0437\u0430\u043D"), data.personalities.map(p => /*#__PURE__*/React.createElement("option", {
    key: p.id,
    value: p.id
  }, p.title)))), /*#__PURE__*/React.createElement(Field, {
    label: "\u041F\u0440\u043E\u0438\u0437\u0432\u0435\u0434\u0435\u043D\u0438\u0435 (\u043D\u0435\u043E\u0431\u044F\u0437\u0430\u0442\u0435\u043B\u044C\u043D\u043E)"
  }, /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: workId,
    onChange: e => setWorkId(e.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u041D\u0435 \u0443\u043A\u0430\u0437\u0430\u043D\u043E"), data.works.map(w => /*#__PURE__*/React.createElement("option", {
    key: w.id,
    value: w.id
  }, w.title))))), /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.genre.title
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "genre",
    selectedIds: nsTagIds(tagIds, "genre", data),
    onChange: ids => setTagIds(withNsTagIds(tagIds, "genre", data, ids))
  })), /*#__PURE__*/React.createElement(Field, {
    label: TAG_NS_LABELS.literaryGenre.title
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: "literaryGenre",
    selectedIds: nsTagIds(tagIds, "literaryGenre", data),
    onChange: ids => setTagIds(withNsTagIds(tagIds, "literaryGenre", data, ids))
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-foot"
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, "\u041E\u0442\u043C\u0435\u043D\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: save
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C"))));
}
function MaterialsForArticlesTab({
  data,
  update,
  projectPath,
  nav,
  setNav
}) {
  const q = nav.materials.q;
  const setQ = v => setNav(n => ({
    ...n,
    materials: {
      ...n.materials,
      q: v
    }
  }));
  const [openId, setOpenId] = useState(null);
  const [creating, setCreating] = useState(false);
  const matches = m => {
    if (!q.trim()) return true;
    const query = q.toLowerCase();
    if ((m.articleName || "").toLowerCase().includes(query)) return true;
    if (stripHtml(m.excerptHtml).toLowerCase().includes(query)) return true;
    const author = m.authorId ? data.personalities.find(p => p.id === m.authorId) : null;
    const work = m.workId ? data.works.find(w => w.id === m.workId) : null;
    if (author && author.title.toLowerCase().includes(query)) return true;
    if (work && work.title.toLowerCase().includes(query)) return true;
    const tagNames = (m.tagIds || []).map(id => findTagAnywhere(data, id)?.name || "").join(" ").toLowerCase();
    return tagNames.includes(query);
  };
  const filtered = [...data.materialsForArticles].filter(matches).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "toolbar"
  }, /*#__PURE__*/React.createElement(TagSearchBar, {
    value: q,
    onChange: setQ,
    placeholder: "\u041F\u043E\u0438\u0441\u043A \u043F\u043E \u0441\u0442\u0430\u0442\u044C\u0435, \u0442\u0435\u043A\u0441\u0442\u0443, \u0430\u0432\u0442\u043E\u0440\u0443, \u043F\u0440\u043E\u0438\u0437\u0432\u0435\u0434\u0435\u043D\u0438\u044E, \u0442\u0435\u0433\u0430\u043C\u2026"
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: () => setCreating(true)
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B")), filtered.length === 0 ? /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u043E\u0432 \u043F\u043E\u043A\u0430 \u043D\u0435\u0442. \u041D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB+ \u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\xBB, \u0447\u0442\u043E\u0431\u044B \u0441\u043E\u0437\u0434\u0430\u0442\u044C \u0432\u0440\u0443\u0447\u043D\u0443\u044E, \u0438\u043B\u0438 \u0432\u044B\u0434\u0435\u043B\u0438\u0442\u0435 \u0444\u0440\u0430\u0433\u043C\u0435\u043D\u0442 \u0442\u0435\u043A\u0441\u0442\u0430 \u0432 \u043B\u044E\u0431\u043E\u043C \u0440\u0430\u0437\u0434\u0435\u043B\u0435 \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB\uD83D\uDCF0 \u0412 \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B\xBB \u0432 \u043F\u0430\u043D\u0435\u043B\u0438 \u0438\u043D\u0441\u0442\u0440\u0443\u043C\u0435\u043D\u0442\u043E\u0432 \u0440\u0435\u0434\u0430\u043A\u0442\u043E\u0440\u0430.") : /*#__PURE__*/React.createElement("div", {
    className: "grid"
  }, filtered.map(m => /*#__PURE__*/React.createElement(MaterialCard, {
    key: m.id,
    data: data,
    m: m,
    onOpen: () => setOpenId(m.id)
  }))), openId && /*#__PURE__*/React.createElement(MaterialModal, {
    data: data,
    update: update,
    id: openId,
    onClose: () => setOpenId(null)
  }), creating && /*#__PURE__*/React.createElement(MaterialCreateModal, {
    data: data,
    update: update,
    projectPath: projectPath,
    onClose: () => setCreating(false)
  }));
}
;
// ============================================================
// sections.js + app.js — Персоналии / История и теория / Произведения,
// оболочка приложения, работа с файлом проекта.
// ============================================================

const SECTION_CONFIGS = {
  personalities: {
    key: "personalities",
    title: "Персоналии",
    addLabel: "+ Автор",
    titlePlaceholder: "Имя автора",
    tagPickers: ["classTag"],
    hasKind: false,
    showAuthorLink: false,
    showAddWork: true,
    hasColorPicker: true,
    viewModes: [{
      id: "new",
      label: "Сначала новые"
    }, {
      id: "old",
      label: "Сначала старые"
    }, {
      id: "classes",
      label: "По классам",
      groupBy: "classTag"
    }]
  },
  theoryHistory: {
    key: "theoryHistory",
    title: "История и теория",
    addLabel: "+ Материал",
    titlePlaceholder: "Название материала",
    tagPickers: ["literaryGenre", "classTag"],
    hasKind: true,
    showAuthorLink: false,
    showAddWork: false,
    hasColorPicker: true,
    viewModes: [{
      id: "new",
      label: "Сначала новые"
    }, {
      id: "old",
      label: "Сначала старые"
    }, {
      id: "topic",
      label: "По темам",
      groupBy: "literaryGenre"
    }, {
      id: "classes",
      label: "По классам",
      groupBy: "classTag"
    }]
  },
  works: {
    key: "works",
    title: "Произведения",
    addLabel: "+ Произведение",
    titlePlaceholder: "Название произведения",
    tagPickers: ["genre", "literaryGenre", "literaryDirection", "classTag"],
    hasKind: false,
    showAuthorLink: true,
    showAddWork: false,
    hasColorPicker: false,
    viewModes: [{
      id: "new",
      label: "Сначала новые"
    }, {
      id: "old",
      label: "Сначала старые"
    }, {
      id: "topic",
      label: "По темам",
      groupBy: "literaryGenre"
    }, {
      id: "direction",
      label: "По литературному направлению",
      groupBy: "literaryDirection"
    }, {
      id: "classes",
      label: "По классам",
      groupBy: "classTag"
    }]
  }
};
function entityTagObjs(data, cfg, e) {
  return (e.tagIds || []).map(id => (cfg.tagPickers || []).map(ns => data.tags[ns].find(t => t.id === id)).find(Boolean)).filter(Boolean);
}
function EntityCard({
  data,
  cfg,
  e,
  onOpen,
  selected,
  onToggle,
  dim
}) {
  const tags = entityTagObjs(data, cfg, e);
  const author = cfg.key === "works" && e.authorId ? data.personalities.find(p => p.id === e.authorId) : null;
  const kindTag = cfg.hasKind ? data.tags.historyTheoryKind.find(k => k.id === e.kind) : null;
  return /*#__PURE__*/React.createElement("div", {
    className: "card" + (onToggle ? " hasbox" : "") + (selected ? " sel" : ""),
    onClick: () => onOpen(e.id),
    style: {
      ...(e.color ? {
        borderLeft: `4px solid ${e.color}`
      } : {}),
      ...(dim ? {
        opacity: 0.5
      } : {})
    }
  }, onToggle && /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    className: "selbox",
    checked: !!selected,
    readOnly: true,
    title: "\u0412\u044B\u0431\u0440\u0430\u0442\u044C",
    onClick: ev => {
      ev.stopPropagation();
      onToggle();
    }
  }), /*#__PURE__*/React.createElement("div", {
    className: "ct-title"
  }, e.title || "Без названия"), author && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\u270E ", author.title), kindTag && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, kindTag.name), /*#__PURE__*/React.createElement("div", {
    className: "ct-preview"
  }, stripHtml(e.course?.lectureHtml).slice(0, 140) || "Пока пусто…"), tags.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "ct-tags"
  }, tags.slice(0, 3).map(t => /*#__PURE__*/React.createElement(Badge, {
    key: t.id,
    tag: t
  }))), ((e.forClasses || []).length > 0 || (e.forCourseIds || []).length > 0) && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\u0414\u043B\u044F: ", [...(e.forClasses || []).map(c => c + " кл."), ...(e.forCourseIds || []).map(id => (data.courses.find(c => c.id === id) || {}).name).filter(Boolean)].join(", ")), (e.hiddenFor || []).length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\uD83D\uDEAB \u041D\u0435 \u043E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0435\u0442\u0441\u044F: ", (e.hiddenFor || []).map(x => scopeLabelOf(data, x)).join(", ")), (e.alwaysFor || []).length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "ct-sub"
  }, "\u2605 \u0412\u0441\u0435\u0433\u0434\u0430: ", (e.alwaysFor || []).map(x => scopeLabelOf(data, x)).join(", ")));
}
function EntityList({
  data,
  update,
  cfg,
  onOpen,
  nav,
  setNav
}) {
  const q = nav[cfg.key].q;
  const view = nav[cfg.key].view;
  const setQ = v => setNav(n => ({
    ...n,
    [cfg.key]: {
      ...n[cfg.key],
      q: v
    }
  }));
  const setView = v => setNav(n => ({
    ...n,
    [cfg.key]: {
      ...n[cfg.key],
      view: v
    }
  }));
  const items = data[cfg.key];
  const matType = {
    personalities: "personality",
    theoryHistory: "theory",
    works: "work"
  }[cfg.key];
  const [sel, setSel] = useState([]);
  const [hideOpen, setHideOpen] = useState(false);
  const [ctxFilter, setCtxFilter] = useState("");
  const [showHiddenToo, setShowHiddenToo] = useState(false);
  const filterCtx = ctxFilter ? scopeFromKey(ctxFilter) : null;
  const hiddenReasonOf = e => filterCtx ? materialHiddenReason(data, matType, e, filterCtx, null) : null;
  const toggleSel = id => setSel(prev => prev.includes(id) ? prev.filter(x => x !== id) : [...prev, id]);
  const applyHide = (keys, mode) => {
    const ids = sel.slice();
    update(d => {
      d[cfg.key].forEach(e => {
        if (!ids.includes(e.id)) return;
        applyScopeMode(e, keys, mode);
      });
      return d;
    });
    setHideOpen(false);
    showToast(scopeModeToast(mode) + keys.map(k => scopeLabelOf(data, k)).join(", ") + " · карточек: " + ids.length);
  };
  const matches = e => {
    if (filterCtx && !showHiddenToo && hiddenReasonOf(e)) return false;
    if (!q.trim()) return true;
    const query = q.toLowerCase();
    if ((e.title || "").toLowerCase().includes(query)) return true;
    const tagNames = entityTagObjs(data, cfg, e).map(t => t.name).join(" ").toLowerCase();
    if (tagNames.includes(query)) return true;
    if (cfg.key === "works" && e.authorId) {
      const author = data.personalities.find(p => p.id === e.authorId);
      if (author && author.title.toLowerCase().includes(query)) return true;
    }
    return false;
  };
  const filtered = items.filter(matches);
  const addNew = async () => {
    const title = await PromptStr(cfg.titlePlaceholder + ":");
    if (!title || !title.trim()) return;
    const e = migrateEntity({
      title: title.trim(),
      kind: cfg.hasKind ? "theory" : null
    }, {});
    update(d => {
      d[cfg.key] = [...d[cfg.key], e];
      return d;
    });
    onOpen(e.id);
  };
  const activeViewMode = cfg.viewModes.find(v => v.id === view);
  let body;
  if (filtered.length === 0) {
    body = /*#__PURE__*/React.createElement("div", {
      className: "empty"
    }, "\u041F\u043E\u043A\u0430 \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435\u0442. \u041D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB", cfg.addLabel, "\xBB, \u0447\u0442\u043E\u0431\u044B \u0441\u043E\u0437\u0434\u0430\u0442\u044C \u043F\u0435\u0440\u0432\u0443\u044E \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0443.");
  } else if (activeViewMode && activeViewMode.groupBy) {
    const ns = activeViewMode.groupBy;
    const groups = {};
    filtered.forEach(e => {
      const idsOfNs = (e.tagIds || []).filter(id => (data.tags[ns] || []).some(t => t.id === id));
      if (idsOfNs.length === 0) (groups.__none = groups.__none || []).push(e);else idsOfNs.forEach(id => (groups[id] = groups[id] || []).push(e));
    });
    const ids = Object.keys(groups).filter(k => k !== "__none").sort((a, b) => {
      const na = data.tags[ns].find(t => t.id === a)?.name || "";
      const nb = data.tags[ns].find(t => t.id === b)?.name || "";
      return na.localeCompare(nb, "ru");
    });
    if (groups.__none) ids.push("__none");
    body = ids.map(gid => /*#__PURE__*/React.createElement("div", {
      key: gid
    }, /*#__PURE__*/React.createElement("div", {
      className: "group-h"
    }, gid === "__none" ? "Без темы" : data.tags[ns].find(t => t.id === gid)?.name || "—"), /*#__PURE__*/React.createElement("div", {
      className: "grid"
    }, groups[gid].map(e => /*#__PURE__*/React.createElement(EntityCard, {
      key: e.id,
      data: data,
      cfg: cfg,
      e: e,
      onOpen: onOpen,
      selected: sel.includes(e.id),
      onToggle: () => toggleSel(e.id),
      dim: !!hiddenReasonOf(e)
    })))));
  } else {
    const sorted = [...filtered].sort((a, b) => view === "old" ? new Date(a.createdAt) - new Date(b.createdAt) : new Date(b.createdAt) - new Date(a.createdAt));
    body = /*#__PURE__*/React.createElement("div", {
      className: "grid"
    }, sorted.map(e => /*#__PURE__*/React.createElement(EntityCard, {
      key: e.id,
      data: data,
      cfg: cfg,
      e: e,
      onOpen: onOpen,
      selected: sel.includes(e.id),
      onToggle: () => toggleSel(e.id),
      dim: !!hiddenReasonOf(e)
    })));
  }
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "toolbar"
  }, /*#__PURE__*/React.createElement(TagSearchBar, {
    value: q,
    onChange: setQ
  }), /*#__PURE__*/React.createElement(ViewModeToggle, {
    value: view,
    onChange: setView,
    options: cfg.viewModes
  }), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    onClick: addNew
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 14
  }), " ", cfg.addLabel)), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8,
      margin: "0 0 8px",
      alignItems: "center"
    }
  }, /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u041F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u0442\u044C \u0434\u043B\u044F:"), /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: ctxFilter,
    onChange: ev => setCtxFilter(ev.target.value)
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0412\u0441\u0435 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438"), CLASSES.map(c => /*#__PURE__*/React.createElement(React.Fragment, {
    key: c
  }, /*#__PURE__*/React.createElement("option", {
    value: "c:" + c
  }, c, " \u043A\u043B\u0430\u0441\u0441 (\u0431\u0435\u0437 \u043B\u0438\u0442\u0435\u0440\u044B)"), lettersForGrade(data, c).map(cl => /*#__PURE__*/React.createElement("option", {
    key: cl.id,
    value: "l:" + c + ":" + cl.id
  }, c, cl.name)))), data.courses.map(c => /*#__PURE__*/React.createElement("option", {
    key: c.id,
    value: "k:" + c.id
  }, "\u041A\u0443\u0440\u0441 \xAB", c.name, "\xBB"))), filterCtx && /*#__PURE__*/React.createElement("label", {
    className: "row",
    style: {
      gap: 5,
      fontSize: 12.5
    }
  }, /*#__PURE__*/React.createElement("input", {
    type: "checkbox",
    checked: showHiddenToo,
    onChange: ev => setShowHiddenToo(ev.target.checked)
  }), " \u043F\u043E\u043A\u0430\u0437\u044B\u0432\u0430\u0442\u044C \u0438 \u0441\u043A\u0440\u044B\u0442\u044B\u0435 (\u0437\u0430\u0442\u0435\u043C\u043D\u0435\u043D\u044B)"), filterCtx && /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\u0421\u043A\u0440\u044B\u0442\u043E \u0442\u0430\u043C: ", items.filter(e => hiddenReasonOf(e)).length, " \u0438\u0437 ", items.length), sel.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    sm: true,
    onClick: () => setHideOpen("hide")
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 12
  }), " \u041D\u0435 \u043E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0442\u044C \u0432\u2026 (", sel.length, ")"), sel.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setHideOpen("always")
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "check",
    size: 12
  }), " \u041E\u0442\u043E\u0431\u0440\u0430\u0436\u0430\u0442\u044C \u0432\u0441\u0435\u0433\u0434\u0430\u2026 (", sel.length, ")"), sel.length > 0 && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setSel([])
  }, "\u0421\u043D\u044F\u0442\u044C \u0432\u044B\u0431\u043E\u0440")), body, hideOpen && /*#__PURE__*/React.createElement(HidePoemsModal, {
    data: data,
    count: sel.length,
    variant: hideOpen,
    onApply: applyHide,
    onClose: () => setHideOpen(false)
  }));
}
function EntityPage({
  data,
  update,
  cfg,
  id,
  onBack,
  projectPath
}) {
  const entity = data[cfg.key].find(e => e.id === id);
  const editorRef = useRef(null);
  const [printOpen, setPrintOpen] = useState(false);
  if (!entity) return /*#__PURE__*/React.createElement("div", {
    className: "empty"
  }, "\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B \u043D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D. ", /*#__PURE__*/React.createElement(Btn, {
    onClick: onBack
  }, "\u041D\u0430\u0437\u0430\u0434"));
  const patchEntity = patch => update(d => {
    const idx = d[cfg.key].findIndex(e => e.id === id);
    if (idx >= 0) {
      d[cfg.key][idx] = {
        ...d[cfg.key][idx],
        ...patch,
        updatedAt: nowISO()
      };
      if (patch.color !== undefined) syncColorToLessonSystems(d, cfg.key, id, patch.color);
    }
    return d;
  });
  const setCourse = fnOrObj => update(d => {
    const idx = d[cfg.key].findIndex(e => e.id === id);
    if (idx >= 0) {
      const prevCourse = d[cfg.key][idx].course;
      d[cfg.key][idx] = {
        ...d[cfg.key][idx],
        course: typeof fnOrObj === "function" ? fnOrObj(prevCourse) : fnOrObj,
        updatedAt: nowISO()
      };
    }
    return d;
  });
  const removeEntity = async () => {
    if (!(await Confirm(`Удалить «${entity.title || "без названия"}»? Действие необратимо.`))) return;
    update(d => {
      d[cfg.key] = d[cfg.key].filter(e => e.id !== id);
      if (cfg.key === "personalities") d.works.forEach(w => {
        if (w.authorId === id) w.authorId = null;
      });
      return d;
    });
    onBack();
  };
  const addWork = async () => {
    const title = await PromptStr("Название произведения:");
    if (!title || !title.trim()) return;
    const w = migrateEntity({
      title: title.trim(),
      authorId: id
    }, {});
    update(d => {
      d.works = [...d.works, w];
      return d;
    });
    showToast("Произведение создано — оно появилось в разделе «Произведения»");
  };
  const linkedWorks = cfg.key === "personalities" ? data.works.filter(w => w.authorId === id) : [];
  return /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      marginBottom: 14
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onBack
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "left",
    size: 15
  }), " \u041D\u0430\u0437\u0430\u0434"), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: () => setPrintOpen(true),
    title: "\u0420\u0430\u0441\u043F\u0435\u0447\u0430\u0442\u0430\u0442\u044C \u0442\u0435\u043A\u0441\u0442 \u0438 \u0432\u043E\u043F\u0440\u043E\u0441\u044B \u044D\u0442\u043E\u0439 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  }), " \u041F\u0435\u0447\u0430\u0442\u044C"), cfg.showAddWork && /*#__PURE__*/React.createElement(Btn, {
    kind: "gold",
    onClick: addWork
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "plus",
    size: 13
  }), " \u041F\u0440\u043E\u0438\u0437\u0432\u0435\u0434\u0435\u043D\u0438\u0435"), /*#__PURE__*/React.createElement(Btn, {
    kind: "danger",
    onClick: removeEntity
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "trash",
    size: 13
  }), " \u0423\u0434\u0430\u043B\u0438\u0442\u044C")), /*#__PURE__*/React.createElement("input", {
    className: "titleinput",
    value: entity.title,
    placeholder: cfg.titlePlaceholder,
    onChange: e => patchEntity({
      title: e.target.value
    })
  }), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      margin: "10px 0 16px"
    }
  }, cfg.hasKind && /*#__PURE__*/React.createElement("div", {
    className: "seg"
  }, data.tags.historyTheoryKind.map(k => /*#__PURE__*/React.createElement("button", {
    key: k.id,
    "aria-pressed": entity.kind === k.id ? "true" : undefined,
    onClick: () => patchEntity({
      kind: k.id
    })
  }, k.name))), cfg.showAuthorLink && /*#__PURE__*/React.createElement("select", {
    className: "toolselect",
    value: entity.authorId || "",
    onChange: e => patchEntity({
      authorId: e.target.value || null
    })
  }, /*#__PURE__*/React.createElement("option", {
    value: ""
  }, "\u0410\u0432\u0442\u043E\u0440 \u043D\u0435 \u0443\u043A\u0430\u0437\u0430\u043D"), data.personalities.map(p => /*#__PURE__*/React.createElement("option", {
    key: p.id,
    value: p.id
  }, p.title)))), cfg.hasColorPicker && /*#__PURE__*/React.createElement(Field, {
    label: "\u0426\u0432\u0435\u0442 \u043A\u0430\u0440\u0442\u043E\u0447\u043A\u0438 (\u043F\u0435\u0440\u0435\u043D\u043E\u0441\u0438\u0442\u0441\u044F \u0432 \u0431\u043B\u043E\u043A \u0441\u0438\u0441\u0442\u0435\u043C\u044B \u0443\u0440\u043E\u043A\u043E\u0432)"
  }, /*#__PURE__*/React.createElement(ColorPicker, {
    value: entity.color || "",
    onChange: c => patchEntity({
      color: c
    })
  })), cfg.tagPickers.map(ns => /*#__PURE__*/React.createElement(Field, {
    key: ns,
    label: TAG_NS_LABELS[ns].title + (ns === "classTag" ? " (для каких классов предлагается при составлении уроков)" : "")
  }, /*#__PURE__*/React.createElement(TagPicker, {
    data: data,
    update: update,
    ns: ns,
    selectedIds: entity.tagIds,
    onChange: ids => patchEntity({
      tagIds: ids
    })
  }))), /*#__PURE__*/React.createElement(Field, {
    label: "\u0414\u043B\u044F \u043A\u0430\u043A\u0438\u0445 \u043A\u0443\u0440\u0441\u043E\u0432 \u0434\u043E\u0441\u0442\u0443\u043F\u043D\u043E (\u043A\u043B\u0430\u0441\u0441\u044B \u0437\u0430\u0434\u0430\u044E\u0442\u0441\u044F \u043C\u0435\u0442\u043A\u0430\u043C\u0438 \xAB\u041A\u043B\u0430\u0441\u0441\u044B\xBB \u0432\u044B\u0448\u0435: \u043C\u0435\u0442\u043A\u0430 \u043A\u043B\u0430\u0441\u0441\u0430 = \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0435\u0442\u0441\u044F \u0442\u043E\u043B\u044C\u043A\u043E \u0432 \u044D\u0442\u043E\u043C \u043A\u043B\u0430\u0441\u0441\u0435; \u043F\u0443\u0441\u0442\u043E \u2014 \u0432\u0441\u0435\u043C)"
  }, /*#__PURE__*/React.createElement(AudiencePicker, {
    data: data,
    classes: entity.forClasses || [],
    courseIds: entity.forCourseIds || [],
    showClasses: false,
    onChange: v => patchEntity({
      forClasses: v.classes,
      forCourseIds: v.courseIds
    })
  })), /*#__PURE__*/React.createElement(Field, {
    label: "\u041D\u0435 \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0442\u044C (\u043D\u0430\u043F\u0440\u0438\u043C\u0435\u0440, \u0434\u043B\u044F \u0432\u043D\u0435\u043A\u043B\u0430\u0441\u0441\u043D\u043E\u0433\u043E \u0447\u0442\u0435\u043D\u0438\u044F) / \u043F\u0440\u0435\u0434\u043B\u0430\u0433\u0430\u0442\u044C \u0432\u0441\u0435\u0433\u0434\u0430 \u043F\u0440\u0438 \u0441\u043E\u0441\u0442\u0430\u0432\u043B\u0435\u043D\u0438\u0438 \u0443\u0440\u043E\u043A\u043E\u0432"
  }, /*#__PURE__*/React.createElement(HiddenScopesRow, {
    data: data,
    hiddenFor: entity.hiddenFor,
    alwaysFor: entity.alwaysFor,
    onChange: v => patchEntity({
      hiddenFor: v.hiddenFor,
      alwaysFor: v.alwaysFor
    })
  })), /*#__PURE__*/React.createElement(AddToLessonPicker, {
    data: data,
    update: update,
    itemType: {
      personalities: "personality",
      theoryHistory: "theory",
      works: "work"
    }[cfg.key],
    refId: entity.id
  }), printOpen && /*#__PURE__*/React.createElement(PrintMaterialModal, {
    data: data,
    type: {
      personalities: "personality",
      theoryHistory: "theory",
      works: "work"
    }[cfg.key],
    entityId: entity.id,
    lesson: null,
    onClose: () => setPrintOpen(false)
  }), linkedWorks.length > 0 && /*#__PURE__*/React.createElement("div", {
    className: "field"
  }, /*#__PURE__*/React.createElement("label", null, "\u041F\u0440\u043E\u0438\u0437\u0432\u0435\u0434\u0435\u043D\u0438\u044F \u044D\u0442\u043E\u0433\u043E \u0430\u0432\u0442\u043E\u0440\u0430"), /*#__PURE__*/React.createElement("div", {
    className: "row wrap"
  }, linkedWorks.map(w => /*#__PURE__*/React.createElement("span", {
    key: w.id,
    className: "chip"
  }, w.title)))), /*#__PURE__*/React.createElement(CourseOfLesson, {
    data: data,
    update: update,
    course: entity.course,
    onChange: setCourse,
    projectPath: projectPath,
    editorRef: editorRef,
    sourceInfo: {
      section: cfg.title,
      entityId: entity.id,
      entityTitle: entity.title,
      authorId: cfg.key === "personalities" ? entity.id : cfg.key === "works" ? entity.authorId : null,
      workId: cfg.key === "works" ? entity.id : null,
      genreTagIds: nsTagIds(entity.tagIds, "genre", data)
    }
  }));
}
class ErrorBoundary extends React.Component {
  constructor(props) {
    super(props);
    this.state = {
      hasError: false
    };
  }
  static getDerivedStateFromError() {
    return {
      hasError: true
    };
  }
  componentDidCatch(err, info) {
    console.error(err, info);
  }
  render() {
    if (this.state.hasError) {
      return /*#__PURE__*/React.createElement("div", {
        className: "gate"
      }, /*#__PURE__*/React.createElement("h1", null, "\u0427\u0442\u043E-\u0442\u043E \u043F\u043E\u0448\u043B\u043E \u043D\u0435 \u0442\u0430\u043A"), /*#__PURE__*/React.createElement("p", {
        className: "muted"
      }, "\u041D\u0435 \u043F\u0435\u0440\u0435\u0436\u0438\u0432\u0430\u0439\u0442\u0435 \u2014 \u0434\u0430\u043D\u043D\u044B\u0435 \u0443\u0436\u0435 \u0441\u043E\u0445\u0440\u0430\u043D\u044F\u044E\u0442\u0441\u044F \u0430\u0432\u0442\u043E\u043C\u0430\u0442\u0438\u0447\u0435\u0441\u043A\u0438. \u041F\u0435\u0440\u0435\u0437\u0430\u043F\u0443\u0441\u0442\u0438\u0442\u0435 \u043F\u0440\u0438\u043B\u043E\u0436\u0435\u043D\u0438\u0435."));
    }
    return this.props.children;
  }
}
function App() {
  const [data, setData] = useState(null);
  const [projectPath, setProjectPath] = useState(null);
  const [tab, setTab] = useState("lessons");
  const [hubSub, setHubSub] = useState("personalities"); // раздел внутри «Материалов»
  const [openId, setOpenId] = useState(null);
  const [pendingLessonId, setPendingLessonId] = useState(null);
  // Откуда открыли урок: если из календаря — «Назад» вернёт в календарь (с тем же месяцем и режимом,
  // они хранятся в nav), а если из списка «Уроки» — в список. Ручной клик по вкладке сбрасывает.
  const [lessonOrigin, setLessonOrigin] = useState(null);
  const goTab = t => {
    setTab(t);
    setLessonOrigin(null);
  }; // открытая карточка/урок при смене раздела не сбрасываются
  const goHub = k => {
    setHubSub(k);
    setOpenId(null);
  };
  const openJournalFor = ({
    trackKey,
    letter
  }) => {
    setNav(n => ({
      ...n,
      planning: {
        ...n.planning,
        subTab: "journal",
        journalTrackKey: trackKey,
        journalLetterFilter: letter || "all"
      }
    }));
    setLessonOrigin(null);
    setTab("calendar");
    window.scrollTo(0, 0);
  };
  const openLessonFrom = (originTab, lesson) => {
    setLessonOrigin({
      tab: originTab,
      scroll: window.scrollY
    });
    setNav(n => ({
      ...n,
      lessons: {
        ...n.lessons,
        openId: lesson.id
      }
    }));
    setTab("lessons");
  };
  const lessonOriginBack = lessonOrigin ? () => {
    const o = lessonOrigin;
    setLessonOrigin(null);
    setTab(o.tab);
    setTimeout(() => window.scrollTo(0, o.scroll || 0), 60);
  } : null;
  const [loaded, setLoaded] = useState(false);
  const [missingAtt, setMissingAtt] = useState([]); // вложения, на которые есть ссылки, но файлов нет в папке «_files»
  const [attHidden, setAttHidden] = useState(false);
  const [attTick, setAttTick] = useState(0);
  // Состояние навигации (открытая литера, вид, поиск, вкладка) — живёт здесь, на уровне
  // всего приложения, а не внутри вкладок, поэтому не сбрасывается при переключении между ними.
  const [nav, setNav] = useState({
    personalities: {
      q: "",
      view: "new"
    },
    theoryHistory: {
      q: "",
      view: "new"
    },
    works: {
      q: "",
      view: "new"
    },
    exercises: {
      q: "",
      openTypeId: null
    },
    lessons: {
      target: {
        kind: "class",
        value: CLASSES[0]
      },
      letterFilter: "all",
      openId: null
    },
    planning: {
      subTab: "journal",
      journalTrackKey: null,
      journalLetterFilter: "all",
      calMonth: curMonthYm(),
      calView: "weeks",
      calFilterKey: "all",
      calLetterFilter: "all",
      clTrackSel: {
        kind: "class",
        value: CLASSES[0]
      },
      clSubSel: "none",
      clListView: "alphabet"
    },
    materials: {
      q: ""
    }
  });
  const saveTimer = useRef(null);
  const importRef = useRef(null);
  // При запуске «Уроки» и «Планирование» открываются на первом избранном классе (по порядку 5–11; нет — первый избранный курс).
  // Боковая панель «Открыть журнал» на странице «Уроки»: журнал рядом с уроком, обе части листаются независимо
  const [jPanel, setJPanel] = useState({
    open: false,
    trackKey: null,
    letter: "all",
    width: 620
  });
  const [winW, setWinW] = useState(typeof window !== "undefined" ? window.innerWidth : 1200);
  useEffect(() => {
    const onR = () => setWinW(window.innerWidth);
    window.addEventListener("resize", onR);
    return () => window.removeEventListener("resize", onR);
  }, []);
  const leftPaneRef = useRef(null);
  const toggleJournalPanel = () => {
    if (jPanel.open) {
      setJPanel(p => ({
        ...p,
        open: false
      }));
      return;
    }
    const lesson = nav.lessons.openId ? data.lessons.find(l => l.id === nav.lessons.openId) : null;
    let trackKey, letter;
    if (lesson) {
      trackKey = trackKeyOf(lesson);
      letter = lesson.classLetterId || (!lesson.courseId && lettersForGrade(data, lesson.class).length ? "none" : "all");
    } else {
      trackKey = trackKeyOf(nav.lessons.target.kind, nav.lessons.target.value);
      letter = nav.lessons.letterFilter;
    }
    setJPanel(p => ({
      ...p,
      open: true,
      trackKey,
      letter,
      width: Math.round(Math.max(420, Math.min(900, window.innerWidth * 0.5)))
    }));
  };
  const startJournalPanelResize = e => {
    e.preventDefault();
    const move = ev => setJPanel(p => ({
      ...p,
      width: Math.max(380, Math.min(window.innerWidth - 360, window.innerWidth - ev.clientX))
    }));
    const up = () => {
      window.removeEventListener("pointermove", move);
      window.removeEventListener("pointerup", up);
    };
    window.addEventListener("pointermove", move);
    window.addEventListener("pointerup", up);
  };
  const favInitDone = useRef(false);
  useEffect(() => {
    if (!data || favInitDone.current) return;
    favInitDone.current = true;
    const fav = data.favorites || {
      classes: [],
      courses: []
    };
    const c = CLASSES.find(x => fav.classes.includes(x));
    const course = c ? null : data.courses.find(x => fav.courses.includes(x.id));
    if (!c && !course) return;
    const kind = c ? "class" : "course",
      value = c || course.id;
    const firstLetter = c ? (lettersForGrade(data, c)[0] || {}).id : null;
    const key = trackKeyOf(kind, value);
    setNav(n => ({
      ...n,
      lessons: {
        ...n.lessons,
        target: {
          kind,
          value
        },
        letterFilter: firstLetter || "all"
      },
      planning: {
        ...n.planning,
        journalTrackKey: key,
        journalLetterFilter: firstLetter || "all",
        calFilterKey: key,
        calLetterFilter: firstLetter || "all",
        clTrackSel: {
          kind,
          value
        },
        clSubSel: firstLetter || "none"
      }
    }));
  }, [data]);

  // Настройки журнала, которые не теряются при переходах между разделами: «Скрыть пустые даты» (запоминается и между
  // запусками), режим правки отметок и введённые, но ещё не записанные значения.
  const [jui, setJuiState] = useState(() => {
    let hideEmpty = false;
    try {
      hideEmpty = localStorage.getItem("lc.hideEmptyDates") === "1";
    } catch (e) {/* без запоминания */}
    return {
      hideEmpty,
      editing: false,
      draft: {}
    };
  });
  const setJui = patch => setJuiState(p => {
    const next = {
      ...p,
      ...(typeof patch === "function" ? patch(p) : patch)
    };
    if (next.hideEmpty !== p.hideEmpty) {
      try {
        localStorage.setItem("lc.hideEmptyDates", next.hideEmpty ? "1" : "0");
      } catch (e) {/* без запоминания */}
    }
    return next;
  });
  // История изменений журнала (отметки и итоговые отметки): Ctrl+Z / Ctrl+Y и кнопки «Отменить / Вернуть»; до 40 шагов.
  // При замене всего проекта (открытие файла, восстановление) история сбрасывается.
  const undoStack = useRef([]);
  const redoStack = useRef([]);
  const lastSetData = useRef(null);
  const [histTick, setHistTick] = useState(0);
  const update = useCallback(fn => setData(prev => {
    const next = fn(structuredClone(prev));
    try {
      if (next && next.lessons) normalizeSystemArchive(next);
    } catch (e) {/* архив систем — необязательная часть */}
    try {
      if (prev && next && JSON.stringify([prev.calendar.marks, prev.calendar.periodMarks]) !== JSON.stringify([next.calendar.marks, next.calendar.periodMarks])) {
        undoStack.current.push({
          marks: prev.calendar.marks,
          periodMarks: prev.calendar.periodMarks || []
        });
        if (undoStack.current.length > 40) undoStack.current.shift();
        redoStack.current = [];
        queueMicrotask(() => setHistTick(t => t + 1));
      }
    } catch (e) {/* история — необязательная часть */}
    lastSetData.current = next;
    return next;
  }), []);
  useEffect(() => {
    if (data && data !== lastSetData.current) {
      undoStack.current = [];
      redoStack.current = [];
      setHistTick(t => t + 1);
    }
  }, [data]);
  const stepHistory = (from, to, text) => {
    const snap = from.current.pop();
    if (!snap) return;
    setData(prev => {
      to.current.push({
        marks: prev.calendar.marks,
        periodMarks: prev.calendar.periodMarks || []
      });
      const next = {
        ...prev,
        calendar: {
          ...prev.calendar,
          marks: snap.marks,
          periodMarks: snap.periodMarks
        }
      };
      lastSetData.current = next;
      return next;
    });
    setHistTick(t => t + 1);
    showToast(text);
  };
  const undoMarks = () => stepHistory(undoStack, redoStack, "↶ Отменено последнее изменение отметок");
  const redoMarks = () => stepHistory(redoStack, undoStack, "↷ Изменение возвращено");
  const journalVisibleRef = useRef(false);
  const histFnRef = useRef({});
  histFnRef.current = {
    undoMarks,
    redoMarks
  };
  useEffect(() => {
    const h = e => {
      if (!(e.ctrlKey || e.metaKey) || !journalVisibleRef.current) return;
      const t = e.target;
      if (t && (t.tagName === "INPUT" || t.tagName === "TEXTAREA" || t.tagName === "SELECT" || t.isContentEditable)) return;
      if (e.code === "KeyZ" && !e.shiftKey) {
        e.preventDefault();
        histFnRef.current.undoMarks();
      } else if (e.code === "KeyY" || e.code === "KeyZ" && e.shiftKey) {
        e.preventDefault();
        histFnRef.current.redoMarks();
      }
    };
    window.addEventListener("keydown", h);
    return () => window.removeEventListener("keydown", h);
  }, []);
  useEffect(() => {
    (async () => {
      if (!window.litAPI) {
        setData(defaultState());
        setLoaded(true);
        return;
      }
      let opened = false;
      const last = await window.litAPI.getLastProjectPath();
      if (last) {
        const res = await window.litAPI.readPath(last);
        if (res.ok) {
          try {
            setData(migrateState(JSON.parse(res.text)));
            setProjectPath(res.path);
            opened = true;
          } catch (e) {/* файл повреждён — попробуем восстановиться из теневой копии ниже */}
        }
      }
      // Основной файл не найден (удалён с диска) или ещё не был выбран, или повреждён —
      // восстанавливаемся из теневой копии, которая пишется на каждое изменение независимо
      // от того, сохранён ли проект в конкретный файл.
      if (!opened) {
        const mirror = await window.litAPI.readMirror();
        if (mirror && mirror.ok) {
          try {
            setData(migrateState(JSON.parse(mirror.text)));
            opened = true;
            showToast(last ? "Файл проекта не найден на диске — данные восстановлены из автосохранения. Сохраните файл заново («Сохранить как»)." : "Восстановлены несохранённые данные из автосохранения. Рекомендуем сохранить файл («Сохранить как»).");
          } catch (e) {/* теневая копия тоже повреждена — начинаем с пустого проекта */}
        }
      }
      setLoaded(true);
    })();
  }, []);

  // Теневая копия пишется всегда, независимо от того, выбран ли файл проекта, — это то,
  // что позволяет ничего не терять, даже если пользователь удалит .json с диска или ещё
  // ни разу не сохранялся. Настоящий файл проекта (если выбран) сохраняется тут же следом.
  useEffect(() => {
    if (!loaded || !data || !window.litAPI) return;
    clearTimeout(saveTimer.current);
    saveTimer.current = setTimeout(() => {
      window.litAPI.saveMirror(JSON.stringify(data));
      if (projectPath) window.litAPI.save(projectPath, JSON.stringify(data));
    }, 300);
    return () => clearTimeout(saveTimer.current);
  }, [data, loaded, projectPath]);

  // Проверка вложений: ссылки на файлы (litfile:///…) должны иметь файлы в папке «<проект>_files».
  useEffect(() => {
    window.__litProjectPath = projectPath;
  }, [projectPath]);
  useEffect(() => {
    const f = () => setAttTick(t => t + 1);
    window.addEventListener("lit-attachments-changed", f);
    return () => window.removeEventListener("lit-attachments-changed", f);
  }, []);
  useEffect(() => {
    if (!loaded || !projectPath || !window.litAPI || !window.litAPI.missingAttachments) {
      setMissingAtt([]);
      return;
    }
    let cancelled = false;
    const t = setTimeout(async () => {
      try {
        const json = JSON.stringify(data);
        const refs = new Map();
        const re = /litfile:\/\/\/([^"\\\s<>']+)(?:[^>]*>)?([^<]*)/g;
        let m;
        while (m = re.exec(json)) {
          const rel = decodeURIComponent(m[1]);
          if (!refs.has(rel) || !refs.get(rel) && m[2]) refs.set(rel, (m[2] || "").replace(/^⬇\s*/, "").trim());
        }
        if (!refs.size) {
          if (!cancelled) setMissingAtt([]);
          return;
        }
        const res = await window.litAPI.missingAttachments(projectPath, Array.from(refs.keys()));
        if (!cancelled && res && res.ok) setMissingAtt(res.missing.map(rel => ({
          rel,
          name: refs.get(rel) || rel
        })));
      } catch (e) {/* проверка — необязательная часть */}
    }, 1500);
    return () => {
      cancelled = true;
      clearTimeout(t);
    };
  }, [loaded, projectPath, attTick, data]);
  const restoreMissing = async a => {
    const res = await window.litAPI.restoreAttachment(projectPath, a.rel, a.name);
    if (res && res.ok) {
      showToast("Файл «" + a.name + "» возвращён");
      setAttTick(t => t + 1);
    } else if (res && res.error) showToast("Не удалось вернуть файл: " + res.error);
  };

  // На планшете проект всегда хранится в браузерном хранилище — отдельный «файл» выбирать не нужно.
  useEffect(() => {
    if (loaded && data && !projectPath && window.litDefaultProjectPath) setProjectPath(window.litDefaultProjectPath);
  }, [loaded, data, projectPath]);

  // ---- Синхронизация проекта через Яндекс.Диск (файл literature.js в папке приложения на Диске) ----
  // Версии разных устройств СЛИВАЮТСЯ (mergeProjects): правки, сделанные на ноутбуке и на планшете, не теряются.
  const [syncInfo, setSyncInfo] = useState({
    clientId: "",
    connected: false,
    lastSyncAt: null
  });
  const [syncStatus, setSyncStatus] = useState("off"); // off | idle | syncing | ok | error | paused
  const [syncError, setSyncError] = useState("");
  const [syncNote, setSyncNote] = useState(""); // что сделала последняя синхронизация
  const [syncModalOpen, setSyncModalOpen] = useState(false);
  const [backupsOpen, setBackupsOpen] = useState(false);
  const syncBusy = useRef(false);
  const syncSwitched = useRef(false); // вручную открыт другой файл / начат новый проект — перед синхронизацией спросим
  const dataRef = useRef(null);
  const lastSyncedDataRef = useRef(null); // объект data, который последним совпадал с Диском
  const syncInfoRef = useRef(syncInfo);
  const syncOnceRef = useRef(null);
  dataRef.current = data;
  syncInfoRef.current = syncInfo;
  const hasSyncApi = !!(window.litAPI && window.litAPI.syncInfo);
  const refreshSyncInfo = async () => {
    if (!hasSyncApi) return syncInfo;
    const info = await window.litAPI.syncInfo();
    setSyncInfo(info);
    setSyncStatus(st => !info.connected ? "off" : st === "off" ? "idle" : st);
    return info;
  };
  useEffect(() => {
    refreshSyncInfo().catch(() => {});
  }, []);
  const handleSyncFailure = res => {
    if (res.auth) {
      setSyncInfo(i => ({
        ...i,
        connected: false
      }));
      setSyncStatus("off");
    } else setSyncStatus("error");
    setSyncError(res.error || "Ошибка синхронизации");
    showToast("⚠ Синхронизация не удалась: " + (res.error || "ошибка"));
  };
  const syncOnce = async () => {
    if (syncBusy.current || !syncInfoRef.current.connected) return;
    syncBusy.current = true;
    setSyncStatus("syncing");
    setSyncError("");
    setSyncNote("");
    let retrySoon = false;
    let note = "Изменений нет — на Диске и здесь одно и то же.";
    try {
      const pull = await window.litAPI.syncPull();
      if (!pull.ok) {
        handleSyncFailure(pull);
        return;
      }
      if (pull.status === "remote-newer") {
        // На Диске есть версия, которой здесь нет: СЛИВАЕМ её с локальной, а не выбираем одну из двух.
        const baseRes = await window.litAPI.syncGetBase().catch(() => null);
        let remoteObj;
        try {
          remoteObj = JSON.parse(pull.text);
        } catch (e) {
          throw new Error("Файл на Диске повреждён — синхронизация остановлена");
        }
        // ↓ от чтения текущих данных до записи результата — без пауз, чтобы не потерять правки, сделанные во время синхронизации
        const cur = dataRef.current;
        let baseObj = null;
        try {
          baseObj = baseRes && baseRes.ok && baseRes.text ? JSON.parse(baseRes.text) : null;
        } catch (e) {/* без базы */}
        const curObj = cur ? JSON.parse(JSON.stringify(cur)) : null;
        // Новое устройство без общей версии и без своих данных (пустой проект «с нуля») — берём версию с Диска целиком,
        // иначе к ней добавились бы пустые стандартные теги и т.п.
        const takeRemote = !curObj || !baseObj && isEmptyProject(curObj);
        const merged = takeRemote ? remoteObj : mergeProjects(baseObj, curObj, remoteObj);
        const localDiffers = curObj && !sameJson(curObj, remoteObj);
        const mergedChangesLocal = !curObj || !sameJson(merged, curObj);
        let finalState = cur;
        if (mergedChangesLocal) {
          finalState = migrateState(merged);
          dataRef.current = finalState;
          setData(finalState);
        }
        const mergedText = JSON.stringify(finalState);
        if (localDiffers && !takeRemote && window.litAPI.snapshot) {
          // страховка: обе исходные версии сохраняются отдельными копиями
          window.litAPI.snapshot("до-слияния-это-устройство", JSON.stringify(curObj)).catch(() => {});
          window.litAPI.snapshot("до-слияния-с-диска", pull.text).catch(() => {});
        }
        if (sameJson(merged, remoteObj)) {
          await window.litAPI.syncMarkSynced(pull.md5, mergedText);
          lastSyncedDataRef.current = finalState;
          note = takeRemote ? "Проект получен с Яндекс.Диска." : "С Диска получены новые правки (у вас на Диске ничего нового не было, отправлять нечего).";
        } else {
          await window.litAPI.syncAcceptRemote(pull.md5, pull.text);
          const push = await window.litAPI.syncPush(mergedText, false);
          if (!push.ok) {
            handleSyncFailure(push);
            return;
          }
          if (push.status === "conflict") {
            retrySoon = true;
            note = "Пока шла синхронизация, на Диске появилась ещё одна версия — повторяю.";
          } else {
            lastSyncedDataRef.current = finalState;
            note = mergedChangesLocal ? "Версии объединены: правки с Диска добавлены сюда, ваши — отправлены на Диск. Копии «до слияния» сохранены." : "Ваши правки отправлены на Диск.";
          }
        }
      } else if (dataRef.current && dataRef.current !== lastSyncedDataRef.current) {
        const cur = dataRef.current;
        const push = await window.litAPI.syncPush(JSON.stringify(cur), false);
        if (!push.ok) {
          handleSyncFailure(push);
          return;
        }
        if (push.status === "conflict") {
          retrySoon = true;
          note = "Пока шла синхронизация, на Диске появилась новая версия — повторяю.";
        } // на следующем круге она сольётся
        else {
          lastSyncedDataRef.current = cur;
          note = push.status === "nochange" ? "Изменений нет — на Диске уже то же самое." : "Ваши правки отправлены на Яндекс.Диск.";
        }
      } else if (dataRef.current) {
        lastSyncedDataRef.current = dataRef.current;
        note = "Изменений нет — на Диске и здесь одно и то же.";
      }
      setSyncInfo(await window.litAPI.syncInfo());
      setSyncStatus("ok");
      setSyncNote(note);
      showToast("🔄 " + note);
    } catch (e) {
      setSyncStatus("error");
      setSyncError(String(e && e.message || e));
      showToast("⚠ Синхронизация не удалась: " + String(e && e.message || e).slice(0, 140));
    } finally {
      syncBusy.current = false;
      if (retrySoon) setTimeout(() => syncOnceRef.current(), 1500);
    }
  };
  syncOnceRef.current = syncOnce;
  // Синхронизация только РУЧНАЯ: запускается кнопкой «Синхронизировать сейчас» в окне «☁». Сама программа ничего не отправляет и не подтягивает.
  // Есть ли правки, ещё не отправленные на Диск: сравниваем отпечаток текущего проекта с отпечатком при последней синхронизации.
  const [syncDirty, setSyncDirty] = useState(false);
  useEffect(() => {
    if (!loaded || !data || !syncInfo.connected || !window.litAPI.syncIsLocalChanged) {
      setSyncDirty(false);
      return;
    }
    const t = setTimeout(async () => {
      try {
        setSyncDirty(!!(await window.litAPI.syncIsLocalChanged(JSON.stringify(data), null)));
      } catch (e) {/* подсказка — необязательная часть */}
    }, 800);
    return () => clearTimeout(t);
  }, [data, loaded, syncInfo.connected, syncInfo.lastSyncAt, syncStatus]);
  const syncLabel = syncStatus === "off" ? "Диск" : syncStatus === "syncing" ? "Синхр…" : syncStatus === "error" ? "Ошибка!" : syncInfo.lastSyncAt ? new Date(syncInfo.lastSyncAt).toLocaleTimeString("ru-RU", {
    hour: "2-digit",
    minute: "2-digit"
  }) : "Подключён";

  // Кнопка «Синхронизировать»: если перед этим вручную открыт другой файл или начат новый проект — сначала спрашиваем.
  const runManualSync = async () => {
    if (syncBusy.current) {
      showToast("Синхронизация уже идёт — подождите");
      return;
    }
    if (syncSwitched.current) {
      if (!(await Confirm("Вы открыли другой файл. При синхронизации его данные будут ОБЪЕДИНЕНЫ с данными на Диске, а записи, которых в этом файле нет, могут быть удалены и с Диска. Синхронизировать этот файл с Диском?", "Синхронизировать"))) return;
      syncSwitched.current = false;
    }
    syncOnceRef.current();
  };

  // Ежедневная автокопия проекта (ноутбук — папка «Резервные копии» рядом с файлом, планшет — хранилище программы).
  useEffect(() => {
    if (!loaded || !data || !projectPath || !window.litAPI || !window.litAPI.dailyBackup) return;
    const run = () => {
      try {
        window.litAPI.dailyBackup(projectPath).catch(() => {});
      } catch (e) {/* копия — необязательная часть */}
    };
    const t0 = setTimeout(run, 4000);
    const iv = setInterval(run, 3600000);
    return () => {
      clearTimeout(t0);
      clearInterval(iv);
    };
  }, [loaded, !!data, projectPath]);
  const openProjectDialog = async () => {
    const res = await window.litAPI.openDialog();
    if (res && res.ok) {
      try {
        setData(migrateState(JSON.parse(res.text)));
        setProjectPath(res.path);
        setOpenId(null);
        syncSwitched.current = true;
      } catch (e) {
        showToast("Не удалось прочитать файл");
      }
    }
  };
  const newProject = () => {
    setData(defaultState());
    setProjectPath(null);
    syncSwitched.current = true;
    showToast("Новый проект создан — сохраните его через «Сохранить как»");
  };
  const saveAs = async () => {
    const res = await window.litAPI.saveAsDialog(JSON.stringify(data), "Мои материалы по литературе.json");
    if (res && res.ok) {
      setProjectPath(res.path);
      showToast("Сохранено: " + res.path);
    }
  };
  const backup = async () => {
    if (!projectPath) {
      showToast("Сначала сохраните проект");
      return;
    }
    const res = await window.litAPI.backup(projectPath);
    if (res && res.ok) showToast("Резервная копия сохранена");
  };
  const exportJson = () => {
    const blob = new Blob([JSON.stringify(data)], {
      type: "application/json"
    });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `konstruktor-literatury-export-${new Date().toISOString().slice(0, 10)}.json`;
    a.click();
    URL.revokeObjectURL(url);
  };
  const handleImportFile = file => {
    const reader = new FileReader();
    reader.onload = async () => {
      try {
        const parsed = JSON.parse(reader.result);
        if (!(await Confirm("Импортировать файл? Текущие данные будут заменены — при необходимости сначала сделайте резервную копию."))) return;
        setData(migrateState(parsed));
        setOpenId(null);
      } catch (e) {
        showToast("Не удалось прочитать файл импорта");
      }
    };
    reader.readAsText(file);
  };
  if (!loaded) return /*#__PURE__*/React.createElement("div", {
    className: "gate"
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0417\u0430\u0433\u0440\u0443\u0437\u043A\u0430\u2026"));
  const syncModalEl = /*#__PURE__*/React.createElement(React.Fragment, null, syncModalOpen && /*#__PURE__*/React.createElement(SyncModal, {
    info: syncInfo,
    status: syncStatus,
    error: syncError,
    onClose: () => setSyncModalOpen(false),
    onChanged: refreshSyncInfo,
    onSyncNow: runManualSync,
    note: syncNote,
    dirty: syncDirty,
    onDisconnected: () => {
      lastSyncedDataRef.current = null;
      setSyncStatus("off");
    }
  }), backupsOpen && /*#__PURE__*/React.createElement(BackupsModal, {
    data: data,
    projectPath: projectPath,
    onClose: () => setBackupsOpen(false),
    onRestore: obj => {
      setData(migrateState(obj));
      setOpenId(null);
      setBackupsOpen(false);
      showToast("Копия восстановлена");
    }
  }));
  if (!data) {
    return /*#__PURE__*/React.createElement("div", {
      className: "gate"
    }, syncModalEl, /*#__PURE__*/React.createElement("img", {
      src: "vendor/logo.png",
      alt: ""
    }), /*#__PURE__*/React.createElement("h1", null, "\u041A\u043E\u043D\u0441\u0442\u0440\u0443\u043A\u0442\u043E\u0440 \u0443\u0440\u043E\u043A\u043E\u0432 \u043B\u0438\u0442\u0435\u0440\u0430\u0442\u0443\u0440\u044B"), /*#__PURE__*/React.createElement("p", {
      className: "muted"
    }, "\u041D\u0430\u0447\u043D\u0438\u0442\u0435 \u043D\u043E\u0432\u044B\u0439 \u043F\u0440\u043E\u0435\u043A\u0442 \u0438\u043B\u0438 \u043E\u0442\u043A\u0440\u043E\u0439\u0442\u0435 \u0443\u0436\u0435 \u0441\u0443\u0449\u0435\u0441\u0442\u0432\u0443\u044E\u0449\u0438\u0439 \u0444\u0430\u0439\u043B \u043C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u043E\u0432."), /*#__PURE__*/React.createElement("div", {
      className: "row"
    }, /*#__PURE__*/React.createElement(Btn, {
      kind: "primary",
      onClick: newProject
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "plus",
      size: 15
    }), " \u041D\u043E\u0432\u044B\u0439 \u043F\u0440\u043E\u0435\u043A\u0442"), /*#__PURE__*/React.createElement(Btn, {
      onClick: openProjectDialog
    }, /*#__PURE__*/React.createElement(Icon, {
      name: "folder",
      size: 15
    }), " \u041E\u0442\u043A\u0440\u044B\u0442\u044C \u0444\u0430\u0439\u043B\u2026"), hasSyncApi && /*#__PURE__*/React.createElement(Btn, {
      onClick: () => setSyncModalOpen(true)
    }, "\u2601 \u0412\u0437\u044F\u0442\u044C \u0441 \u042F\u043D\u0434\u0435\u043A\u0441.\u0414\u0438\u0441\u043A\u0430")), hasSyncApi && /*#__PURE__*/React.createElement("p", {
      className: "hint",
      style: {
        maxWidth: 420,
        textAlign: "center"
      }
    }, "\u0423\u0436\u0435 \u0440\u0430\u0431\u043E\u0442\u0430\u0435\u0442\u0435 \u043D\u0430 \u0434\u0440\u0443\u0433\u043E\u043C \u0443\u0441\u0442\u0440\u043E\u0439\u0441\u0442\u0432\u0435? \u041D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB\u0412\u0437\u044F\u0442\u044C \u0441 \u042F\u043D\u0434\u0435\u043A\u0441.\u0414\u0438\u0441\u043A\u0430\xBB, \u0432\u043E\u0439\u0434\u0438\u0442\u0435 \u0432 \u042F\u043D\u0434\u0435\u043A\u0441 \u2014 \u043F\u0440\u043E\u0435\u043A\u0442 \u0437\u0430\u0433\u0440\u0443\u0437\u0438\u0442\u0441\u044F \u0441\u044E\u0434\u0430."));
  }
  const cfg = tab === "hub" ? SECTION_CONFIGS[hubSub] : null;
  journalVisibleRef.current = tab === "today" || tab === "calendar" && nav.planning.subTab === "journal" || tab === "lessons" && jPanel.open;
  const split = tab === "lessons" && jPanel.open && winW >= 900;
  const juiFull = {
    ...jui,
    undo: undoMarks,
    redo: redoMarks,
    canUndo: undoStack.current.length > 0,
    canRedo: redoStack.current.length > 0
  };
  const headerEl = /*#__PURE__*/React.createElement("header", {
    className: "top"
  }, /*#__PURE__*/React.createElement("div", {
    className: "top-row"
  }, /*#__PURE__*/React.createElement("div", {
    className: "brand"
  }, /*#__PURE__*/React.createElement("img", {
    className: "mark",
    src: "vendor/logo.png",
    alt: ""
  }), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("h1", null, "\u041A\u043E\u043D\u0441\u0442\u0440\u0443\u043A\u0442\u043E\u0440 \u0443\u0440\u043E\u043A\u043E\u0432 \u043B\u0438\u0442\u0435\u0440\u0430\u0442\u0443\u0440\u044B"), /*#__PURE__*/React.createElement("p", null, projectPath ? projectPath.split(/[\\/]/).pop() : "Проект ещё не сохранён"))), /*#__PURE__*/React.createElement("div", {
    className: "topnav"
  }, /*#__PURE__*/React.createElement("button", {
    "aria-pressed": tab === "today" ? "true" : undefined,
    onClick: () => goTab("today")
  }, "\u0421\u0435\u0433\u043E\u0434\u043D\u044F"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": tab === "lessons" ? "true" : undefined,
    onClick: () => goTab("lessons")
  }, "\u0423\u0440\u043E\u043A\u0438"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": tab === "calendar" ? "true" : undefined,
    onClick: () => goTab("calendar")
  }, "\u041F\u043B\u0430\u043D\u0438\u0440\u043E\u0432\u0430\u043D\u0438\u0435"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": tab === "hub" ? "true" : undefined,
    onClick: () => goTab("hub")
  }, "\u041C\u0430\u0442\u0435\u0440\u0438\u0430\u043B\u044B"), /*#__PURE__*/React.createElement("button", {
    "aria-pressed": tab === "tags" ? "true" : undefined,
    onClick: () => goTab("tags"),
    title: "\u0422\u0435\u0433\u0438 \u0438 \u0432\u0438\u0434\u044B"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "tag",
    size: 13
  }))), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1
    }
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: openProjectDialog,
    title: "\u041E\u0442\u043A\u0440\u044B\u0442\u044C \u0444\u0430\u0439\u043B"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "folder",
    size: 13
  })), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: saveAs,
    title: "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C \u043A\u0430\u043A"
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "save",
    size: 13
  })), hasSyncApi && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setSyncModalOpen(true),
    title: syncInfo.connected ? "Синхронизация с Яндекс.Диском" + (syncStatus === "error" ? ": " + syncError : "") : "Подключить Яндекс.Диск — общий проект на ноутбуке и планшете"
  }, /*#__PURE__*/React.createElement("span", {
    style: syncStatus === "error" ? {
      color: "var(--neg)",
      fontWeight: 700
    } : undefined
  }, "\u2601 ", syncLabel, syncDirty ? /*#__PURE__*/React.createElement("span", {
    title: "\u0415\u0441\u0442\u044C \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u044F, \u043A\u043E\u0442\u043E\u0440\u044B\u0435 \u0435\u0449\u0451 \u043D\u0435 \u043E\u0442\u043F\u0440\u0430\u0432\u043B\u0435\u043D\u044B \u043D\u0430 \u0414\u0438\u0441\u043A",
    style: {
      color: "#c98a2b",
      fontWeight: 900
    }
  }, " \u25CF") : null)), hasSyncApi && syncInfo.connected && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    disabled: syncStatus === "syncing",
    onClick: runManualSync,
    title: "\u0421\u0438\u043D\u0445\u0440\u043E\u043D\u0438\u0437\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0441 \u042F\u043D\u0434\u0435\u043A\u0441.\u0414\u0438\u0441\u043A\u043E\u043C \u0441\u0435\u0439\u0447\u0430\u0441 (\u0432\u0440\u0443\u0447\u043D\u0443\u044E): \u043E\u0431\u044A\u0435\u0434\u0438\u043D\u0438\u0442\u044C \u043F\u0440\u0430\u0432\u043A\u0438 \u0441 \u0414\u0438\u0441\u043A\u0430 \u0438 \u043E\u0442\u043F\u0440\u0430\u0432\u0438\u0442\u044C \u0441\u0432\u043E\u0438"
  }, "\uD83D\uDD04 \u0421\u0438\u043D\u0445\u0440\u043E\u043D\u0438\u0437\u0438\u0440\u043E\u0432\u0430\u0442\u044C"), window.litAPI && window.litAPI.listBackups && /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setBackupsOpen(true),
    title: "\u0415\u0436\u0435\u0434\u043D\u0435\u0432\u043D\u044B\u0435 \u0430\u0432\u0442\u043E\u043A\u043E\u043F\u0438\u0438 \u043F\u0440\u043E\u0435\u043A\u0442\u0430 \u2014 \u043C\u043E\u0436\u043D\u043E \u0432\u0435\u0440\u043D\u0443\u0442\u044C\u0441\u044F \u043A \u043B\u044E\u0431\u043E\u0439 \u0438\u0437 \u043F\u043E\u0441\u043B\u0435\u0434\u043D\u0438\u0445 14"
  }, "\u041A\u043E\u043F\u0438\u0438"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: backup,
    title: "\u0420\u0435\u0437\u0435\u0440\u0432\u043D\u0430\u044F \u043A\u043E\u043F\u0438\u044F"
  }, "\u0411\u044D\u043A\u0430\u043F"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: exportJson,
    title: "\u042D\u043A\u0441\u043F\u043E\u0440\u0442 JSON"
  }, "\u042D\u043A\u0441\u043F\u043E\u0440\u0442"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => importRef.current?.click(),
    title: "\u0418\u043C\u043F\u043E\u0440\u0442 JSON"
  }, "\u0418\u043C\u043F\u043E\u0440\u0442"), /*#__PURE__*/React.createElement("input", {
    ref: importRef,
    type: "file",
    accept: "application/json",
    style: {
      display: "none"
    },
    onChange: e => {
      if (e.target.files[0]) handleImportFile(e.target.files[0]);
      e.target.value = "";
    }
  })));
  const contentEl = /*#__PURE__*/React.createElement("div", {
    className: "appwrap",
    style: split ? {
      padding: "0 18px 60px",
      maxWidth: "none"
    } : undefined
  }, missingAtt.length > 0 && !attHidden && /*#__PURE__*/React.createElement("div", {
    style: {
      background: "#fff4e0",
      border: "1.5px solid #e8b960",
      borderRadius: 10,
      padding: "10px 14px",
      margin: "10px 0",
      fontSize: 13
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      justifyContent: "space-between",
      alignItems: "flex-start",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      fontWeight: 700
    }
  }, "\u26A0 \u041D\u0435 \u043D\u0430\u0439\u0434\u0435\u043D\u044B \u0444\u0430\u0439\u043B\u044B \u0432\u043B\u043E\u0436\u0435\u043D\u0438\u0439: ", missingAtt.length, ". \u0421\u0441\u044B\u043B\u043A\u0438 \u0432 \u043F\u0440\u043E\u0435\u043A\u0442\u0435 \u0435\u0441\u0442\u044C, \u0430 \u0441\u0430\u043C\u0438\u0445 \u0444\u0430\u0439\u043B\u043E\u0432 \u0432 \u043F\u0430\u043F\u043A\u0435 \xAB", projectPath ? projectPath.split(/[\\/]/).pop().replace(/\.[^.]*$/, "") : "", "_files\xBB \u043D\u0435\u0442 \u2014 \u0441\u043A\u043E\u0440\u0435\u0435 \u0432\u0441\u0435\u0433\u043E, \u044D\u0442\u0430 \u043F\u0430\u043F\u043A\u0430 \u043D\u0435 \u043F\u0435\u0440\u0435\u043D\u0435\u0441\u0435\u043D\u0430 \u043D\u0430 \u044D\u0442\u043E\u0442 \u043A\u043E\u043C\u043F\u044C\u044E\u0442\u0435\u0440 \u0432\u043C\u0435\u0441\u0442\u0435 \u0441 \u0444\u0430\u0439\u043B\u043E\u043C \u043F\u0440\u043E\u0435\u043A\u0442\u0430."), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setAttHidden(true)
  }, "\u0421\u043A\u0440\u044B\u0442\u044C")), /*#__PURE__*/React.createElement("div", {
    style: {
      marginTop: 6,
      display: "grid",
      gap: 4
    }
  }, missingAtt.slice(0, 8).map(a => /*#__PURE__*/React.createElement("div", {
    key: a.rel,
    className: "row",
    style: {
      gap: 8
    }
  }, /*#__PURE__*/React.createElement("span", {
    style: {
      flex: 1,
      minWidth: 0,
      overflow: "hidden",
      textOverflow: "ellipsis",
      whiteSpace: "nowrap"
    }
  }, "\uD83D\uDCCE ", a.name), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => restoreMissing(a)
  }, "\u0423\u043A\u0430\u0437\u0430\u0442\u044C \u0444\u0430\u0439\u043B\u2026"))), missingAtt.length > 8 && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u2026\u0438 \u0435\u0449\u0451 ", missingAtt.length - 8))), tab === "tags" && /*#__PURE__*/React.createElement(TagsAdminTab, {
    data: data,
    update: update
  }), tab === "today" && /*#__PURE__*/React.createElement(TodayTab, {
    data: data,
    update: update,
    onOpenLesson: lesson => openLessonFrom("today", lesson),
    onOpenJournal: openJournalFor,
    hist: juiFull
  }), tab === "lessons" && nav.lessons.openId && /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      justifyContent: "flex-end",
      marginBottom: 6
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: toggleJournalPanel,
    title: "\u0416\u0443\u0440\u043D\u0430\u043B \u044D\u0442\u043E\u0433\u043E \u043A\u043B\u0430\u0441\u0441\u0430 \u0440\u044F\u0434\u043E\u043C \u0441\u043E \u0441\u0442\u0440\u0430\u043D\u0438\u0446\u0435\u0439 \u0443\u0440\u043E\u043A\u0430 \u2014 \u043C\u043E\u0436\u043D\u043E \u0441\u0442\u0430\u0432\u0438\u0442\u044C \u043E\u0442\u043C\u0435\u0442\u043A\u0438, \u043D\u0435 \u0443\u0445\u043E\u0434\u044F \u0441 \u0443\u0440\u043E\u043A\u0430"
  }, jPanel.open ? "✕ Закрыть журнал" : "📒 Открыть журнал")), tab === "lessons" && /*#__PURE__*/React.createElement(LessonsTab, {
    data: data,
    update: update,
    projectPath: projectPath,
    initialLessonId: pendingLessonId,
    onConsumedInitial: () => setPendingLessonId(null),
    nav: nav,
    setNav: setNav,
    originBack: lessonOriginBack,
    compact: split,
    journalOpen: jPanel.open,
    onToggleJournal: nav.lessons.openId ? null : toggleJournalPanel
  }), tab === "calendar" && /*#__PURE__*/React.createElement(PlanningTab, {
    jui: juiFull,
    setJui: setJui,
    data: data,
    update: update,
    onOpenLesson: lesson => openLessonFrom("calendar", lesson),
    nav: nav,
    setNav: setNav
  }), tab === "hub" && /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement("div", {
    className: "viewtoggle",
    style: {
      marginBottom: 14
    }
  }, [["personalities", "Персоналии"], ["theoryHistory", "История и теория"], ["works", "Произведения"], ["exercises", "Упражнения"], ["articles", "Материалы для статей"]].map(([k, t]) => /*#__PURE__*/React.createElement("button", {
    key: k,
    "aria-pressed": hubSub === k ? "true" : undefined,
    onClick: () => goHub(k)
  }, t))), hubSub === "exercises" && /*#__PURE__*/React.createElement(ExerciseBankTab, {
    data: data,
    update: update,
    projectPath: projectPath,
    nav: nav,
    setNav: setNav
  }), hubSub === "articles" && /*#__PURE__*/React.createElement(MaterialsForArticlesTab, {
    data: data,
    update: update,
    projectPath: projectPath,
    nav: nav,
    setNav: setNav
  }), cfg && (openId ? /*#__PURE__*/React.createElement(EntityPage, {
    data: data,
    update: update,
    cfg: cfg,
    id: openId,
    onBack: () => setOpenId(null),
    projectPath: projectPath
  }) : /*#__PURE__*/React.createElement(EntityList, {
    data: data,
    update: update,
    cfg: cfg,
    onOpen: setOpenId,
    nav: nav,
    setNav: setNav
  }))));
  const journalPanelEl = tab === "lessons" && jPanel.open ? /*#__PURE__*/React.createElement(React.Fragment, null, split && /*#__PURE__*/React.createElement("div", {
    onPointerDown: startJournalPanelResize,
    title: "\u041F\u043E\u0442\u044F\u043D\u0438\u0442\u0435, \u0447\u0442\u043E\u0431\u044B \u0438\u0437\u043C\u0435\u043D\u0438\u0442\u044C \u0448\u0438\u0440\u0438\u043D\u0443 \u043E\u043A\u043E\u043D",
    style: {
      width: 12,
      flexShrink: 0,
      cursor: "col-resize",
      background: "#dfe7f0",
      display: "flex",
      alignItems: "center",
      justifyContent: "center",
      touchAction: "none"
    }
  }, /*#__PURE__*/React.createElement("div", {
    style: {
      width: 4,
      height: 46,
      borderRadius: 3,
      background: "#9fb3c8"
    }
  })), /*#__PURE__*/React.createElement("div", {
    style: split ? {
      width: jPanel.width,
      flexShrink: 0,
      background: "var(--bg)",
      display: "flex",
      flexDirection: "column",
      minHeight: 0,
      borderLeft: "2px solid var(--accent)"
    } : {
      position: "fixed",
      top: 0,
      right: 0,
      bottom: 0,
      width: "100vw",
      zIndex: 800,
      background: "var(--bg)",
      display: "flex",
      flexDirection: "column"
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      justifyContent: "space-between",
      padding: "10px 16px",
      background: "#fff",
      borderBottom: "1px solid var(--line)"
    }
  }, /*#__PURE__*/React.createElement("h2", {
    style: {
      fontFamily: "var(--font-display)",
      fontSize: 17,
      margin: 0
    }
  }, "\u0416\u0443\u0440\u043D\u0430\u043B"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => setJPanel(p => ({
      ...p,
      open: false
    }))
  }, "\u2715 \u0417\u0430\u043A\u0440\u044B\u0442\u044C \u0436\u0443\u0440\u043D\u0430\u043B")), /*#__PURE__*/React.createElement("div", {
    style: {
      flex: 1,
      minHeight: 0,
      overflow: "auto",
      padding: "12px 16px 40px"
    }
  }, /*#__PURE__*/React.createElement(JournalSection, {
    data: data,
    update: update,
    trackKey: jPanel.trackKey,
    setTrackKey: k => setJPanel(p => ({
      ...p,
      trackKey: k
    })),
    tracks: allTracks(data),
    letterFilter: jPanel.letter,
    setLetterFilter: v => setJPanel(p => ({
      ...p,
      letter: v
    })),
    jui: juiFull,
    setJui: setJui
  })))) : null;
  if (!split) return /*#__PURE__*/React.createElement(React.Fragment, null, headerEl, contentEl, journalPanelEl, syncModalEl);
  return /*#__PURE__*/React.createElement("div", {
    style: {
      height: "100vh",
      display: "flex",
      flexDirection: "column",
      overflow: "hidden"
    }
  }, syncModalEl, headerEl, /*#__PURE__*/React.createElement("div", {
    style: {
      display: "flex",
      flex: 1,
      minHeight: 0
    }
  }, /*#__PURE__*/React.createElement("div", {
    ref: leftPaneRef,
    style: {
      flex: 1,
      minWidth: 0,
      overflowY: "auto",
      overflowX: "hidden"
    }
  }, contentEl), journalPanelEl));
}

// ---- Слияние версий проекта при синхронизации ----
// base — общая версия при прошлой синхронизации (если есть), local — это устройство, remote — Диск.
// Что добавлено или изменено на любом устройстве — сохраняется; записи с id (уроки, ученики, отметки, упражнения…)
// сливаются по id, поля — по отдельности. Удаление применяется, только если вторая сторона эту запись не меняла.
// При изменении одного и того же поля на обоих: с базой побеждает это устройство, без базы (первое подключение) — Диск.
function stableStringify(v) {
  if (v === null || typeof v !== "object") return JSON.stringify(v);
  if (Array.isArray(v)) return "[" + v.map(stableStringify).join(",") + "]";
  return "{" + Object.keys(v).filter(k => v[k] !== undefined).sort().map(k => JSON.stringify(k) + ":" + stableStringify(v[k])).join(",") + "}";
}
const sameJson = (a, b) => stableStringify(a) === stableStringify(b);
function isEmptyProject(o) {
  return ["personalities", "works", "theoryHistory", "examples", "lessons"].every(k => !Array.isArray(o && o[k]) || o[k].length === 0);
}
function mergeValue(b, l, r, hb) {
  // поле отсутствует с одной стороны (например, старая версия программы его не знает) — берём то, что есть
  if (l === undefined) return r;
  if (r === undefined) return l;
  if (sameJson(l, r)) return l;
  const isObj = x => x && typeof x === "object" && !Array.isArray(x);
  if (hb && b !== undefined && !(isObj(l) && isObj(r))) {
    if (sameJson(b, l)) return r;
    if (sameJson(b, r)) return l;
  }
  if (isObj(l) && isObj(r)) return mergeObjectsJson(isObj(b) ? b : undefined, l, r, hb);
  if (Array.isArray(l) && Array.isArray(r)) return mergeArraysJson(Array.isArray(b) ? b : undefined, l, r, hb);
  return pickConflict(b, l, r, hb);
}
// Спорное значение: пустое («», null, []) уступает непустому; иначе с базой побеждает это устройство, без базы — Диск.
function pickConflict(b, l, r, hb) {
  const empty = x => x === "" || x === null || Array.isArray(x) && x.length === 0;
  if (b === undefined || !hb) {
    if (empty(l) && !empty(r)) return r;
    if (empty(r) && !empty(l)) return l;
  }
  return hb ? l : r;
}
function mergeObjectsJson(b, l, r, hb) {
  const out = {};
  new Set([...Object.keys(l), ...Object.keys(r)]).forEach(k => {
    const v = mergeValue(b ? b[k] : undefined, l[k], r[k], hb && !!b);
    if (v !== undefined) out[k] = v;
  });
  return out;
}
function mergeArraysJson(b, l, r, hb) {
  const isIdItem = x => x && typeof x === "object" && !Array.isArray(x) && x.id !== undefined && x.id !== null;
  const idList = a => a.length > 0 && a.every(isIdItem);
  const prim = x => x === null || typeof x !== "object";
  const bUse = hb && Array.isArray(b);
  if ((idList(l) || !l.length) && (idList(r) || !r.length)) {
    const B = new Map((bUse ? b : []).filter(isIdItem).map(x => [x.id, x]));
    const Rm = new Map(r.map(x => [x.id, x]));
    const out = [];
    const seen = new Set();
    l.forEach(x => {
      seen.add(x.id);
      const bx = B.get(x.id);
      const rx = Rm.get(x.id);
      if (rx !== undefined) out.push(mergeValue(bx, x, rx, bx !== undefined));else if (bx !== undefined && sameJson(bx, x)) return; // на Диске удалено, здесь не менялось — удаляем
      else out.push(x);
    });
    r.forEach(y => {
      if (seen.has(y.id)) return;
      const bx = B.get(y.id);
      if (bx !== undefined && sameJson(bx, y)) return; // здесь удалено, на Диске не менялось — удаляем
      out.push(y);
    });
    return out;
  }
  if (l.every(prim) && r.every(prim)) {
    const key = stableStringify;
    const rs = new Set(r.map(key));
    const ls = new Set(l.map(key));
    if (bUse) {
      const bs = new Set(b.map(key));
      const out = l.filter(x => !(bs.has(key(x)) && !rs.has(key(x))));
      r.forEach(y => {
        if (!ls.has(key(y)) && !bs.has(key(y))) out.push(y);
      });
      return out;
    }
    const out = [...l];
    r.forEach(y => {
      if (!ls.has(key(y))) out.push(y);
    });
    return out;
  }
  return pickConflict(b, l, r, hb); // массивы без id (например, схема рассадки) — целиком
}
function mergeProjects(base, local, remote) {
  return mergeValue(base || undefined, local, remote, !!base);
}

// Окно настройки синхронизации с Яндекс.Диском: ID приложения, вход, статус, отключение.
function SyncModal({
  info,
  status,
  error,
  note,
  dirty,
  onClose,
  onChanged,
  onSyncNow,
  onDisconnected
}) {
  const [clientId, setClientId] = useState(info.clientId || "");
  const [manualToken, setManualToken] = useState("");
  const [proxy, setProxy] = useState(info.proxyUrl || "");
  const [busy, setBusy] = useState(false);
  const [msg, setMsg] = useState("");
  const inp = {
    width: "100%",
    border: "1px solid var(--line)",
    borderRadius: 8,
    padding: "8px 10px"
  };
  const connect = async () => {
    if (!clientId.trim()) {
      setMsg("Вставьте ID приложения Яндекса (тот же, что в «Конструкторе уроков» по русскому).");
      return;
    }
    setBusy(true);
    setMsg("");
    await window.litAPI.syncSetClientId(clientId);
    const r = await window.litAPI.syncConnect();
    setBusy(false);
    if (r.ok) await onChanged();else if (!r.cancelled) setMsg(r.error || "Не удалось войти");
  };
  const useManualToken = async () => {
    if (!manualToken.trim()) return;
    setBusy(true);
    setMsg("");
    await window.litAPI.syncSetClientId(clientId);
    const r = await window.litAPI.syncSetToken(manualToken);
    setBusy(false);
    if (r.ok) {
      setManualToken("");
      await onChanged();
    } else setMsg(r.error || "Токен не подошёл");
  };
  const disconnect = async () => {
    if (!(await Confirm("Отключить Яндекс.Диск? Файл на Диске останется, но программа перестанет с ним синхронизироваться.", "Отключить"))) return;
    await window.litAPI.syncDisconnect();
    onDisconnected();
    await onChanged();
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 620
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u2601 \u0421\u0438\u043D\u0445\u0440\u043E\u043D\u0438\u0437\u0430\u0446\u0438\u044F \u0447\u0435\u0440\u0435\u0437 \u042F\u043D\u0434\u0435\u043A\u0441.\u0414\u0438\u0441\u043A"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      display: "grid",
      gap: 12
    }
  }, info.connected ? /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("div", null, "\u042F\u043D\u0434\u0435\u043A\u0441.\u0414\u0438\u0441\u043A \u043F\u043E\u0434\u043A\u043B\u044E\u0447\u0451\u043D. ", info.lastSyncAt ? "Последняя синхронизация: " + new Date(info.lastSyncAt).toLocaleString("ru-RU") + "." : "Синхронизаций ещё не было."), /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041F\u0440\u043E\u0435\u043A\u0442 \u043B\u0435\u0436\u0438\u0442 \u043D\u0430 \u0414\u0438\u0441\u043A\u0435 \u0432 \u043F\u0430\u043F\u043A\u0435 \xAB\u041F\u0440\u0438\u043B\u043E\u0436\u0435\u043D\u0438\u044F\xBB \u2192 \u043F\u0430\u043F\u043A\u0430 \u0432\u0430\u0448\u0435\u0433\u043E \u043F\u0440\u0438\u043B\u043E\u0436\u0435\u043D\u0438\u044F, \u0444\u0430\u0439\u043B literature.js (\u0444\u0430\u0439\u043B \xAB\u041A\u043E\u043D\u0441\u0442\u0440\u0443\u043A\u0442\u043E\u0440\u0430 \u043F\u043E \u0440\u0443\u0441\u0441\u043A\u043E\u043C\u0443\xBB \u0440\u044F\u0434\u043E\u043C \u043D\u0435 \u0437\u0430\u0442\u0440\u0430\u0433\u0438\u0432\u0430\u0435\u0442\u0441\u044F)."), /*#__PURE__*/React.createElement("div", {
    style: {
      background: "var(--accent-soft)",
      borderRadius: 10,
      padding: "10px 12px",
      fontSize: 13,
      lineHeight: 1.55
    }
  }, /*#__PURE__*/React.createElement("b", null, "\u0421\u0438\u043D\u0445\u0440\u043E\u043D\u0438\u0437\u0430\u0446\u0438\u044F \u2014 \u0442\u043E\u043B\u044C\u043A\u043E \u0432\u0440\u0443\u0447\u043D\u0443\u044E."), " \u041F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u0430\u043C\u0430 \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u043E\u0442\u043F\u0440\u0430\u0432\u043B\u044F\u0435\u0442 \u0438 \u043D\u0435 \u0437\u0430\u0433\u0440\u0443\u0436\u0430\u0435\u0442. \u041D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB\u0421\u0438\u043D\u0445\u0440\u043E\u043D\u0438\u0437\u0438\u0440\u043E\u0432\u0430\u0442\u044C \u0441\u0435\u0439\u0447\u0430\u0441\xBB \u2014 \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u0440\u0430\u0432\u043D\u0438\u0442 \u0432\u0435\u0440\u0441\u0438\u044E \u043D\u0430 \u0414\u0438\u0441\u043A\u0435 \u0441 \u0432\u0430\u0448\u0435\u0439: ", /*#__PURE__*/React.createElement("b", null, "\u043E\u0431\u044A\u0435\u0434\u0438\u043D\u0438\u0442"), " \u043F\u0440\u0430\u0432\u043A\u0438 \u043E\u0431\u0435\u0438\u0445 (\u0434\u043E\u0431\u0430\u0432\u043B\u0435\u043D\u043D\u043E\u0435 \u0438 \u0438\u0437\u043C\u0435\u043D\u0451\u043D\u043D\u043E\u0435 \u043D\u0430 \u043D\u043E\u0443\u0442\u0431\u0443\u043A\u0435 \u0438 \u043D\u0430 \u043F\u043B\u0430\u043D\u0448\u0435\u0442\u0435 \u0441\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u0441\u044F, \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u043F\u043E\u0442\u0435\u0440\u044F\u0435\u0442\u0441\u044F) \u0438 \u043E\u0442\u043F\u0440\u0430\u0432\u0438\u0442 \u0440\u0435\u0437\u0443\u043B\u044C\u0442\u0430\u0442 \u043D\u0430 \u0414\u0438\u0441\u043A. \u0414\u0435\u043B\u0430\u0439\u0442\u0435 \u044D\u0442\u043E, \u043A\u043E\u0433\u0434\u0430 \u0437\u0430\u043A\u043E\u043D\u0447\u0438\u043B\u0438 \u0440\u0430\u0431\u043E\u0442\u0443 \u043D\u0430 \u043E\u0434\u043D\u043E\u043C \u0443\u0441\u0442\u0440\u043E\u0439\u0441\u0442\u0432\u0435 \u0438 \u043F\u0435\u0440\u0435\u0434 \u043D\u0430\u0447\u0430\u043B\u043E\u043C \u043D\u0430 \u0434\u0440\u0443\u0433\u043E\u043C."), /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13
    }
  }, dirty ? /*#__PURE__*/React.createElement("span", {
    style: {
      color: "#a8741a",
      fontWeight: 700
    }
  }, "\u25CF \u0421 \u043F\u043E\u0441\u043B\u0435\u0434\u043D\u0435\u0439 \u0441\u0438\u043D\u0445\u0440\u043E\u043D\u0438\u0437\u0430\u0446\u0438\u0438 \u0435\u0441\u0442\u044C \u0438\u0437\u043C\u0435\u043D\u0435\u043D\u0438\u044F, \u043A\u043E\u0442\u043E\u0440\u044B\u0435 \u0435\u0449\u0451 \u043D\u0435 \u043E\u0442\u043F\u0440\u0430\u0432\u043B\u0435\u043D\u044B \u043D\u0430 \u0414\u0438\u0441\u043A.") : /*#__PURE__*/React.createElement("span", {
    style: {
      color: "var(--pos, #2e8b57)"
    }
  }, "\u2713 \u0412\u0441\u0435 \u0432\u0430\u0448\u0438 \u043F\u0440\u0430\u0432\u043A\u0438 \u0443\u0436\u0435 \u043D\u0430 \u0414\u0438\u0441\u043A\u0435 (\u043F\u043E\u0441\u043B\u0435 \u043F\u043E\u0441\u043B\u0435\u0434\u043D\u0435\u0439 \u0441\u0438\u043D\u0445\u0440\u043E\u043D\u0438\u0437\u0430\u0446\u0438\u0438 \u043D\u0438\u0447\u0435\u0433\u043E \u043D\u0435 \u043C\u0435\u043D\u044F\u043B\u043E\u0441\u044C).")), status === "ok" && note && /*#__PURE__*/React.createElement("div", {
    style: {
      fontSize: 13
    }
  }, "\u0418\u0442\u043E\u0433: ", note), status === "error" && error && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--neg)"
    }
  }, /*#__PURE__*/React.createElement("b", null, "\u041D\u0435 \u043F\u043E\u043B\u0443\u0447\u0438\u043B\u043E\u0441\u044C:"), " ", error), /*#__PURE__*/React.createElement("div", {
    className: "row wrap",
    style: {
      gap: 8
    }
  }, /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: status === "syncing",
    onClick: onSyncNow
  }, status === "syncing" ? "Синхронизация…" : "🔄 Синхронизировать сейчас"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: disconnect
  }, "\u041E\u0442\u043A\u043B\u044E\u0447\u0438\u0442\u044C"))) : /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gap: 12
    }
  }, window.litRedirectUri && /*#__PURE__*/React.createElement("div", {
    style: {
      background: "#fff8e6",
      borderRadius: 10,
      padding: "12px 14px",
      fontSize: 12.5,
      lineHeight: 1.55
    }
  }, /*#__PURE__*/React.createElement("b", null, "\u0412\u0435\u0440\u0441\u0438\u044F \u0434\u043B\u044F \u043F\u043B\u0430\u043D\u0448\u0435\u0442\u0430."), " \u041F\u0440\u0438\u043B\u043E\u0436\u0435\u043D\u0438\u0435 \u042F\u043D\u0434\u0435\u043A\u0441\u0430 (oauth.yandex.ru) \u0442\u043E \u0436\u0435, \u0447\u0442\u043E \u0434\u043B\u044F \u0440\u0443\u0441\u0441\u043A\u043E\u0433\u043E. \u041E\u0434\u0438\u043D \u0440\u0430\u0437 \u043D\u0443\u0436\u043D\u043E \u0434\u043E\u0431\u0430\u0432\u0438\u0442\u044C \u0432 \u0435\u0433\u043E \u043D\u0430\u0441\u0442\u0440\u043E\u0439\u043A\u0430\u0445 \u043D\u0430 oauth.yandex.ru \u0435\u0449\u0451 \u043E\u0434\u0438\u043D Redirect URI (\u043A\u043D\u043E\u043F\u043A\u0430 \xAB+\xBB):", /*#__PURE__*/React.createElement("div", {
    style: {
      fontFamily: "monospace",
      wordBreak: "break-all",
      margin: "6px 0",
      userSelect: "all",
      background: "#fff",
      padding: "6px 8px",
      borderRadius: 6
    }
  }, window.litRedirectUri), "\u0412\u0445\u043E\u0434 \u043E\u0442\u043A\u0440\u043E\u0435\u0442\u0441\u044F \u0432 \u044D\u0442\u043E\u043C \u0436\u0435 \u043E\u043A\u043D\u0435 \u0431\u0440\u0430\u0443\u0437\u0435\u0440\u0430 \u0438 \u0432\u0435\u0440\u043D\u0451\u0442 \u0432\u0430\u0441 \u043E\u0431\u0440\u0430\u0442\u043D\u043E."), /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0418\u0441\u043F\u043E\u043B\u044C\u0437\u0443\u0435\u0442\u0441\u044F \u0442\u043E \u0436\u0435 \u043F\u0440\u0438\u043B\u043E\u0436\u0435\u043D\u0438\u0435 \u042F\u043D\u0434\u0435\u043A\u0441\u0430, \u0447\u0442\u043E \u0438 \u0432 \xAB\u041A\u043E\u043D\u0441\u0442\u0440\u0443\u043A\u0442\u043E\u0440\u0435 \u0443\u0440\u043E\u043A\u043E\u0432\xBB \u043F\u043E \u0440\u0443\u0441\u0441\u043A\u043E\u043C\u0443: \u0432\u0441\u0442\u0430\u0432\u044C\u0442\u0435 \u0442\u043E\u0442 \u0436\u0435 ", /*#__PURE__*/React.createElement("b", null, "ID \u043F\u0440\u0438\u043B\u043E\u0436\u0435\u043D\u0438\u044F (Client ID)"), " \u0438 \u043D\u0430\u0436\u043C\u0438\u0442\u0435 \xAB\u0412\u043E\u0439\u0442\u0438 \u0447\u0435\u0440\u0435\u0437 \u042F\u043D\u0434\u0435\u043A\u0441\xBB."), /*#__PURE__*/React.createElement(Field, {
    label: "ID \u043F\u0440\u0438\u043B\u043E\u0436\u0435\u043D\u0438\u044F (Client ID)"
  }, /*#__PURE__*/React.createElement("input", {
    style: inp,
    value: clientId,
    onChange: e => setClientId(e.target.value),
    placeholder: "\u043D\u0430\u043F\u0440\u0438\u043C\u0435\u0440, 3f2a\u2026"
  })), /*#__PURE__*/React.createElement("div", null, /*#__PURE__*/React.createElement(Btn, {
    kind: "primary",
    disabled: busy,
    onClick: connect
  }, busy ? "Подождите…" : "Войти через Яндекс")), /*#__PURE__*/React.createElement("details", null, /*#__PURE__*/React.createElement("summary", {
    className: "hint",
    style: {
      cursor: "pointer"
    }
  }, "\u0415\u0441\u043B\u0438 \u043E\u043A\u043D\u043E \u0432\u0445\u043E\u0434\u0430 \u043D\u0435 \u0437\u0430\u043A\u0440\u044B\u043B\u043E\u0441\u044C \u0441\u0430\u043C\u043E \u2014 \u0432\u0441\u0442\u0430\u0432\u0438\u0442\u044C \u0442\u043E\u043A\u0435\u043D \u0432\u0440\u0443\u0447\u043D\u0443\u044E"), /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      marginTop: 8,
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("input", {
    style: inp,
    value: manualToken,
    onChange: e => setManualToken(e.target.value),
    placeholder: "\u0422\u043E\u043A\u0435\u043D \u0441\u043E \u0441\u0442\u0440\u0430\u043D\u0438\u0446\u044B \u042F\u043D\u0434\u0435\u043A\u0441\u0430"
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    disabled: busy,
    onClick: useManualToken
  }, "\u041F\u0440\u0438\u043C\u0435\u043D\u0438\u0442\u044C")))), window.litTabletMode && window.litAPI.syncSetProxy && /*#__PURE__*/React.createElement("details", {
    open: !info.proxyUrl && status === "error"
  }, /*#__PURE__*/React.createElement("summary", {
    className: "hint",
    style: {
      cursor: "pointer"
    }
  }, "\u041F\u043E\u0441\u0440\u0435\u0434\u043D\u0438\u043A \u0434\u043B\u044F \u0441\u043A\u0430\u0447\u0438\u0432\u0430\u043D\u0438\u044F ", info.proxyUrl ? "(настроен)" : "(нужен, если Яндекс не отдаёт файл браузеру)"), /*#__PURE__*/React.createElement("div", {
    style: {
      display: "grid",
      gap: 6,
      marginTop: 8
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0410\u0434\u0440\u0435\u0441 \u0432\u0430\u0448\u0435\u0433\u043E \u043F\u043E\u0441\u0440\u0435\u0434\u043D\u0438\u043A\u0430 \u043D\u0430 Cloudflare Workers (\u0432\u0438\u0434\u0430 https://\u2026workers.dev) \u2014 \u0442\u043E\u0442 \u0436\u0435, \u0447\u0442\u043E \u0434\u043B\u044F \u0440\u0443\u0441\u0441\u043A\u043E\u0433\u043E \u043A\u043E\u043D\u0441\u0442\u0440\u0443\u043A\u0442\u043E\u0440\u0430. \u0427\u0435\u0440\u0435\u0437 \u043D\u0435\u0433\u043E \u043F\u0440\u043E\u0445\u043E\u0434\u0438\u0442 \u0442\u043E\u043B\u044C\u043A\u043E \u0444\u0430\u0439\u043B \u043F\u0440\u043E\u0435\u043A\u0442\u0430, \u0442\u043E\u043A\u0435\u043D \u042F\u043D\u0434\u0435\u043A\u0441\u0430 \u043E\u043D \u043D\u0435 \u043F\u043E\u043B\u0443\u0447\u0430\u0435\u0442."), /*#__PURE__*/React.createElement("div", {
    className: "row",
    style: {
      gap: 6
    }
  }, /*#__PURE__*/React.createElement("input", {
    style: inp,
    value: proxy,
    onChange: e => setProxy(e.target.value),
    placeholder: "https://lessons-proxy.\u0432\u0430\u0448\u0435-\u0438\u043C\u044F.workers.dev"
  }), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: async () => {
      await window.litAPI.syncSetProxy(proxy);
      await onChanged();
      setMsg("Адрес посредника сохранён. Нажмите «Синхронизировать сейчас».");
    }
  }, "\u0421\u043E\u0445\u0440\u0430\u043D\u0438\u0442\u044C")))), msg && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--neg)"
    }
  }, msg))));
}

// Ежедневные автокопии: список за последние 14 дней и возврат к выбранной.
function BackupsModal({
  data,
  projectPath,
  onClose,
  onRestore
}) {
  const [items, setItems] = useState(null);
  const [where, setWhere] = useState("");
  const [err, setErr] = useState("");
  useEffect(() => {
    (async () => {
      try {
        const r = await window.litAPI.listBackups(projectPath);
        if (r.ok) {
          setItems(r.items);
          setWhere(r.where || "");
        } else {
          setItems([]);
          setErr(r.error || "Не удалось получить список копий");
        }
      } catch (e) {
        setItems([]);
        setErr(String(e && e.message || e));
      }
    })();
  }, []);
  const restore = async it => {
    if (!(await Confirm("Вернуть проект к копии от " + it.date + "? Всё, что сделано позже, исчезнет из программы (текущее состояние перед этим сохранится отдельной копией).", "Вернуть"))) return;
    const r = await window.litAPI.readBackup(projectPath, it.id);
    if (!r.ok) {
      setErr(r.error || "Не удалось прочитать копию");
      return;
    }
    let obj;
    try {
      obj = JSON.parse(r.text);
    } catch (e) {
      setErr("Копия повреждена");
      return;
    }
    if (window.litAPI.snapshot) await window.litAPI.snapshot("до-восстановления", JSON.stringify(data)).catch(() => {});
    onRestore(obj);
  };
  return /*#__PURE__*/React.createElement("div", {
    className: "overlay",
    onClick: onClose
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal",
    style: {
      maxWidth: 560
    },
    onClick: e => e.stopPropagation()
  }, /*#__PURE__*/React.createElement("div", {
    className: "modal-head"
  }, /*#__PURE__*/React.createElement("h3", {
    style: {
      flex: 1
    }
  }, "\u0415\u0436\u0435\u0434\u043D\u0435\u0432\u043D\u044B\u0435 \u043A\u043E\u043F\u0438\u0438 \u043F\u0440\u043E\u0435\u043A\u0442\u0430"), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    onClick: onClose
  }, /*#__PURE__*/React.createElement(Icon, {
    name: "close",
    size: 16
  }))), /*#__PURE__*/React.createElement("div", {
    className: "modal-body",
    style: {
      display: "grid",
      gap: 10
    }
  }, /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0420\u0430\u0437 \u0432 \u0434\u0435\u043D\u044C \u043F\u0440\u043E\u0433\u0440\u0430\u043C\u043C\u0430 \u0441\u0430\u043C\u0430 \u0434\u0435\u043B\u0430\u0435\u0442 \u043A\u043E\u043F\u0438\u044E \u043F\u0440\u043E\u0435\u043A\u0442\u0430 (\u0445\u0440\u0430\u043D\u044F\u0442\u0441\u044F \u043F\u043E\u0441\u043B\u0435\u0434\u043D\u0438\u0435 14). ", where ? "Папка: " + where : ""), err && /*#__PURE__*/React.createElement("div", {
    style: {
      color: "var(--neg)"
    }
  }, err), items === null && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u0417\u0430\u0433\u0440\u0443\u0437\u043A\u0430\u2026"), items && items.length === 0 && !err && /*#__PURE__*/React.createElement("div", {
    className: "hint"
  }, "\u041A\u043E\u043F\u0438\u0439 \u043F\u043E\u043A\u0430 \u043D\u0435\u0442 \u2014 \u043F\u0435\u0440\u0432\u0430\u044F \u043F\u043E\u044F\u0432\u0438\u0442\u0441\u044F \u0432\u0441\u043A\u043E\u0440\u0435 \u043F\u043E\u0441\u043B\u0435 \u0437\u0430\u043F\u0443\u0441\u043A\u0430."), items && items.map(it => /*#__PURE__*/React.createElement("div", {
    key: it.id,
    className: "row",
    style: {
      justifyContent: "space-between",
      borderBottom: "1px solid var(--line)",
      paddingBottom: 6
    }
  }, /*#__PURE__*/React.createElement("span", null, /*#__PURE__*/React.createElement("b", null, it.date), " ", /*#__PURE__*/React.createElement("span", {
    className: "hint"
  }, "\xB7 ", Math.round(it.size / 1024), " \u041A\u0411")), /*#__PURE__*/React.createElement(Btn, {
    kind: "ghost",
    sm: true,
    onClick: () => restore(it)
  }, "\u0412\u0435\u0440\u043D\u0443\u0442\u044C \u043A \u044D\u0442\u043E\u0439 \u043A\u043E\u043F\u0438\u0438"))))));
}
ReactDOM.createRoot(document.getElementById("root")).render( /*#__PURE__*/React.createElement(ErrorBoundary, null, /*#__PURE__*/React.createElement(DialogHost, null), /*#__PURE__*/React.createElement(App, null)));