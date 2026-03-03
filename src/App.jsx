import { useState, useRef, useCallback, useEffect } from "react";
import Tesseract from "tesseract.js";
import * as XLSX from "xlsx-js-style";

/* ─── Image Preprocessing ─── */
function preprocessImage(imgUrl) {
  return new Promise((resolve) => {
    const img = new Image();
    img.crossOrigin = "anonymous";
    img.onload = () => {
      const MIN = 1200, MAX = 2400;
      let w = img.width, h = img.height;
      if (w < MIN && h < MIN) { const r = Math.max(MIN / w, MIN / h); w = Math.round(w * r); h = Math.round(h * r); }
      else if (w > MAX || h > MAX) { const r = Math.min(MAX / w, MAX / h); w = Math.round(w * r); h = Math.round(h * r); }
      const c = document.createElement("canvas");
      c.width = w; c.height = h;
      const ctx = c.getContext("2d");
      ctx.drawImage(img, 0, 0, w, h);
      const id = ctx.getImageData(0, 0, w, h), d = id.data;
      const gray = new Uint8Array(w * h);
      for (let i = 0; i < d.length; i += 4) gray[i / 4] = Math.round(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2]);
      const halfWin = 15, threshC = 12;
      const integral = new Float64Array((w + 1) * (h + 1));
      for (let y = 0; y < h; y++) { let rs = 0; for (let x = 0; x < w; x++) { rs += gray[y * w + x]; integral[(y + 1) * (w + 1) + (x + 1)] = integral[y * (w + 1) + (x + 1)] + rs; } }
      for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) {
        const x1 = Math.max(0, x - halfWin), y1 = Math.max(0, y - halfWin), x2 = Math.min(w - 1, x + halfWin), y2 = Math.min(h - 1, y + halfWin);
        const cnt = (x2 - x1 + 1) * (y2 - y1 + 1);
        const sum = integral[(y2 + 1) * (w + 1) + (x2 + 1)] - integral[y1 * (w + 1) + (x2 + 1)] - integral[(y2 + 1) * (w + 1) + x1] + integral[y1 * (w + 1) + x1];
        const val = gray[y * w + x] < (sum / cnt - threshC) ? 0 : 255;
        const idx = (y * w + x) * 4;
        d[idx] = d[idx + 1] = d[idx + 2] = val;
      }
      ctx.putImageData(id, 0, 0);
      resolve(c.toDataURL("image/png"));
    };
    img.onerror = () => resolve(imgUrl);
    img.src = imgUrl;
  });
}

/* ─── OCR Text Cleanup ─── */
function cleanOcrLine(line) {
  return line
    .replace(/[«»„""]/g, '"')
    .replace(/[|\\{}[\]]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function cleanName(name) {
  if (!name) return '';
  let c = name;
  // Latin lookalikes → Cyrillic (OCR often confuses these)
  c = c.replace(/A/g, 'А').replace(/B/g, 'В').replace(/C/g, 'С').replace(/E/g, 'Е')
    .replace(/H/g, 'Н').replace(/I/g, 'І').replace(/K/g, 'К').replace(/M/g, 'М')
    .replace(/O/g, 'О').replace(/P/g, 'Р').replace(/T/g, 'Т').replace(/X/g, 'Х')
    .replace(/a/g, 'а').replace(/c/g, 'с').replace(/e/g, 'е')
    .replace(/i/g, 'і').replace(/o/g, 'о').replace(/p/g, 'р').replace(/x/g, 'х');
  // Remove everything non-Cyrillic except dots, spaces, hyphens, apostrophes
  c = c.replace(/[^А-ЯІЇЄҐа-яіїєґ.\s''\u2019-]/g, '').replace(/\s+/g, ' ').trim();
  // After initials pattern (e.g. "Прізвище Ю.С."), truncate everything after
  const initMatch = c.match(/^(.+?[А-ЯІЇЄҐа-яіїєґ]\.[А-ЯІЇЄҐа-яіїєґ]\.?)/);
  if (initMatch) c = initMatch[1].trim();
  // Remove isolated single Cyrillic letters (OCR noise)
  c = c.replace(/(?:^|\s)[А-ЯІЇЄҐа-яіїєґ](?:\s|$)/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  return c;
}

/* ─── Receipt Parser ─── */
function parseReceipt(text) {
  const lines = text.split("\n").map(l => cleanOcrLine(l)).filter(Boolean);
  let fop = "", td = "", gotivka = 0, kartka = 0, razom = 0, monthIdx = -1;

  for (let i = 0; i < lines.length; i++) {
    let line = lines[i];
    // Normalize Latin lookalikes in the line for keyword matching
    const norm = line.replace(/A/g,'А').replace(/B/g,'В').replace(/C/g,'С').replace(/E/g,'Е')
      .replace(/H/g,'Н').replace(/I/g,'І').replace(/K/g,'К').replace(/M/g,'М')
      .replace(/O/g,'О').replace(/P/g,'Р').replace(/T/g,'Т').replace(/X/g,'Х');
    // Match ФОП with OCR variants
    if (/[ФфF«][ОоО0][ПпР]\s+/i.test(norm) && !fop) {
      fop = norm.replace(/^.*?[ФфF«][ОоО0][ПпР]\s*/i, "").trim();
    }
    if (/Від\s+\d{4}/i.test(line)) {
      const mt = line.match(/Від\s+(\d{4})-(\d{2})-\d{2}/i);
      if (mt) monthIdx = parseInt(mt[2]) - 1;
    }
    const nm = line.match(/([\d\s]+[.,]\d{2})\s*$/);
    const num = nm ? parseFloat(nm[1].replace(/\s/g, "").replace(",", ".")) : null;
    if (num !== null && num > 0) {
      // Require keywords at the START of the line
      // Don't match БЕЗГОТІВК — OCR often misreads ГОТІВКА as БЕЗГОТІВКА
      if (/^\s*ГОТІВКА/i.test(norm) && !gotivka) gotivka = num;
      if (/^\s*КАРТКА/i.test(norm) && !kartka) kartka = num;
      if (/^\s*РАЗОМ/i.test(norm) && !razom) razom = num;
    }
  }

  // Derive card amount mathematically: if РАЗОМ > ГОТІВКА, the difference is card
  if (razom > 0 && gotivka > 0 && razom > gotivka) {
    kartka = Math.round((razom - gotivka) * 100) / 100;
  } else if (razom > 0 && gotivka === 0 && kartka === 0) {
    // Only РАЗОМ found — treat as cash
    gotivka = razom;
  }

  if (fop) {
    for (let i = 0; i < lines.length; i++) {
      const ln = lines[i].replace(/A/g,'А').replace(/O/g,'О').replace(/P/g,'Р');
      if (/[ФфF«][ОоО0][ПпР]\s+/i.test(ln)) {
        for (let j = i + 1; j < Math.min(i + 3, lines.length); j++) {
          const candidate = lines[j].trim();
          if (candidate.length > 1 && candidate.length < 40 &&
            !/м\.\s|район|проспект|буд\.|ІД\s|ПЕРІОДИЧ|КОРОТКИЙ|ЗВІТ|ФІСКАЛЬН/i.test(candidate) &&
            !/^\d/.test(candidate)) {
            td = cleanName(candidate);
            break;
          }
        }
        break;
      }
    }
  }

  if (!fop) {
    for (const line of lines) {
      if (/[А-ЯІЇЄҐа-яіїєґ]{3,}/.test(line) && line.length > 5 && line.length < 60) {
        const c = line.replace(/^.*?[ФфF«][ОоО0][ПпР]\s*/i, "").trim();
        if (c.length > 3 && !/ПЕРІОДИЧ|КОРОТКИЙ|КАФЕ|ЗВІТ|ФІСКАЛЬН|Profit|ВСЬОГО|РАЗОМ|ГОТІВКА|КАРТКА|БЕЗГОТІВК|ПОВЕРНЕННЯ|СУМА|ОБІГ|ПОДАТОК/i.test(c)) { fop = c; break; }
      }
    }
  }

  fop = cleanName(fop);
  td = cleanName(td);

  return { fop, td, gotivka, kartka, razom: razom || (gotivka + kartka), monthIdx, rawText: text };
}

/* ─── Milana Messages ─── */
const MILANA = {
  ua: {
    scanning: [
      "Міланочко, зачекай трішки...",
      "Мілана п'є каву, а я працюю...",
      "Очі Мілани відпочивають...",
      "Читаю чеки замість Мілани...",
      "Мілана, ще секундочку...",
      "Мілана відпочиває, я працюю...",
    ],
    subtitle: "2026",
    ocrLoading: "Готуюсь для Мілани...",
    ocrReady: "Мілана, готова до роботи!",
    drop: "Міланочко, кидай чеки сюди!",
    dropSub: "або натисни тут",
    allDone: "Все готово, Міланочко!",
    stillScanning: "Ще скануються...",
    exported: "Excel для Мілани готовий!",
    noData: "Міланочко, спочатку завантаж чеки",
    autoSaved: "автоматично збережено",
    photosAdded: "фото додано",
    startBtn: "Міланочко, починаємо сканування!",
    pendingMsg: "Готові до сканування",
    deleted: "Мілана прибрала",
  },
  tr: {
    scanning: [
      "Milana biraz beklemeli...",
      "Milana kahve icsin, ben hallederim...",
      "Milana'nin gozleri dinleniyor...",
      "Fisleri Milana icin okuyorum...",
      "Milana, bir saniye daha...",
      "Milana dinlensin, ben calisiyorum...",
    ],
    subtitle: "2026",
    ocrLoading: "Milana icin hazirlaniyor...",
    ocrReady: "Milana, hazirim!",
    drop: "Milana, fisleri buraya birak!",
    dropSub: "veya buraya tikla",
    allDone: "Hepsi tamam, Milana!",
    stillScanning: "Taranmaya devam ediyor...",
    exported: "Milana'nin Excel'i hazir!",
    noData: "Milana, once fisleri yukle",
    autoSaved: "otomatik kaydedildi",
    photosAdded: "foto eklendi",
    startBtn: "Milana, taramayı başlat!",
    pendingMsg: "Taranmaya hazir",
    deleted: "Milana temizledi",
  },
  en: {
    scanning: [
      "Milana, just a moment...",
      "Milana's having coffee, I'm working...",
      "Milana's eyes are resting...",
      "Reading receipts for Milana...",
      "One more second, Milana...",
      "Milana's resting, I'm on it...",
    ],
    subtitle: "2026",
    ocrLoading: "Getting ready for Milana...",
    ocrReady: "Ready for you, Milana!",
    drop: "Milana, drop receipts here!",
    dropSub: "or click here",
    allDone: "All done, Milana!",
    stillScanning: "Still scanning...",
    exported: "Milana's Excel is ready!",
    noData: "Milana, upload receipts first",
    autoSaved: "auto-saved",
    photosAdded: "photos added",
    startBtn: "Milana, start scanning!",
    pendingMsg: "Ready to scan",
    deleted: "Milana cleaned up",
  }
};

/* ─── Translations ─── */
const LANGS = {
  ua: {
    months: ["січень", "лютий", "березень", "квітень", "травень", "червень", "липень", "серпень", "вересень", "жовтень", "листопад", "грудень"],
    monthShort: ["Січ", "Лют", "Бер", "Кві", "Тра", "Чер", "Лип", "Сер", "Вер", "Жов", "Лис", "Гру"],
    title: "Помічник Мілани",
    entry: "Чеки", table: "Таблиця", export: "Завантажити Excel",
    gotivka: "Готівка", bezgotivka: "Готівка", vsogo: "Всього",
    fopLabel: "ФОП", tdLabel: "ТД",
    gotShort: "Готівка", bezShort: "Готівка", vsoShort: "Всього",
    month: "Місяць", rawText: "Текст чеку",
    prev: "Попередні", year: "2026",
    tip: { prev: "Попередній чек", next: "Наступний чек", zoomIn: "Збільшити", zoomOut: "Зменшити", del: "Видалити чек", addPhoto: "Додати фото", entry: "Перегляд чеків", table: "Зведена таблиця", exportXls: "Завантажити Excel-файл", raw: "Показати текст чеку", start: "Запустити сканування всіх чеків", theme: "Змінити тему" },
  },
  tr: {
    months: ["Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran", "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik"],
    monthShort: ["Oca", "Sub", "Mar", "Nis", "May", "Haz", "Tem", "Agu", "Eyl", "Eki", "Kas", "Ara"],
    title: "Milana Asistani",
    entry: "Fisler", table: "Tablo", export: "Excel Indir",
    gotivka: "Nakit", bezgotivka: "Banka", vsogo: "Toplam",
    fopLabel: "FOP", tdLabel: "TD",
    gotShort: "Nakit", bezShort: "Banka", vsoShort: "Toplam",
    month: "Ay", rawText: "Fis metni",
    prev: "Oncekiler", year: "2026",
    tip: { prev: "Önceki fiş", next: "Sonraki fiş", zoomIn: "Yakınlaştır", zoomOut: "Uzaklaştır", del: "Fişi sil", addPhoto: "Fotoğraf ekle", entry: "Fiş görünümü", table: "Özet tablo", exportXls: "Excel dosyasını indir", raw: "Fiş metnini göster", start: "Tüm fişleri taramayı başlat", theme: "Temayı değiştir" },
  },
  en: {
    months: ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"],
    monthShort: ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"],
    title: "Milana's Helper",
    entry: "Receipts", table: "Table", export: "Download Excel",
    gotivka: "Cash", bezgotivka: "Bank", vsogo: "Total",
    fopLabel: "FOP", tdLabel: "TD",
    gotShort: "Cash", bezShort: "Bank", vsoShort: "Total",
    month: "Month", rawText: "Raw text",
    prev: "Previous", year: "2026",
    tip: { prev: "Previous receipt", next: "Next receipt", zoomIn: "Zoom in", zoomOut: "Zoom out", del: "Delete receipt", addPhoto: "Add photo", entry: "Receipt view", table: "Summary table", exportXls: "Download Excel file", raw: "Show receipt text", start: "Start scanning all receipts", theme: "Toggle theme" },
  }
};

const fmt = (n) => n ? Number(n).toLocaleString("uk-UA", { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : "";
const parseNum = (s) => s ? parseFloat(String(s).replace(/\s/g, "").replace(",", ".")) || 0 : 0;

export default function App() {
  /* ─── Theme & Language (persisted) ─── */
  const [theme, setTheme] = useState(() => localStorage.getItem("milana-theme") || "dark");
  const [lang, setLang] = useState(() => {
    const saved = localStorage.getItem("milana-lang");
    return saved && LANGS[saved] ? saved : "ua";
  });
  const t = LANGS[lang];
  const m = MILANA[lang];

  useEffect(() => {
    document.documentElement.setAttribute("data-theme", theme);
    localStorage.setItem("milana-theme", theme);
  }, [theme]);

  useEffect(() => {
    localStorage.setItem("milana-lang", lang);
  }, [lang]);

  const toggleTheme = () => setTheme(p => p === "dark" ? "light" : "dark");

  /* ─── App State ─── */
  const [view, setView] = useState("entry");
  const [notif, setNotif] = useState(null);
  const [dragOver, setDragOver] = useState(false);
  const fileRef = useRef(null);
  const [receipts, setReceipts] = useState([]);
  const [activeIdx, setActiveIdx] = useState(0);
  const [workers, setWorkers] = useState([]);
  const [ocrReady, setOcrReady] = useState(false);
  const [showRaw, setShowRaw] = useState(false);
  const [zoom, setZoom] = useState(1);
  const processingRef = useRef(false);
  const receiptsRef = useRef(receipts);
  receiptsRef.current = receipts;
  const workersRef = useRef(workers);
  workersRef.current = workers;
  const allDoneRef = useRef(false);

  const notify = (msg, type = "success") => { setNotif({ msg, type }); setTimeout(() => setNotif(null), 2500); };

  const prevPairs = [...new Map(
    receipts.filter(r => r.status === "done")
      .filter(r => r.editFop.trim())
      .map(r => [`${r.editFop.trim()}|||${r.editTd.trim()}`, { fop: r.editFop.trim(), td: r.editTd.trim() }])
  ).values()];

  const getLoadingMsg = (id) => m.scanning[id % m.scanning.length];

  const pendingCount = receipts.filter(r => r.status === "pending").length;
  const scanC = receipts.filter(r => ["scanning", "queue"].includes(r.status)).length;
  const doneC = receipts.filter(r => r.status === "done").length;

  /* ─── OCR Workers ─── */
  useEffect(() => {
    let mounted = true;
    (async () => {
      const ws = [];
      for (let i = 0; i < 3; i++) {
        try {
          const w = await Tesseract.createWorker("ukr+rus");
          await w.setParameters({
            tessedit_pageseg_mode: "6",       // Single uniform block of text
            preserve_interword_spaces: "1",   // Keep word spacing
            tessedit_char_blacklist: "~`@#$%^&*{}[]|\\<>",  // Block noise chars
          });
          ws.push(w);
        } catch (e) { console.error(e); }
      }
      if (mounted && ws.length) { setWorkers(ws); setOcrReady(true); }
    })();
    return () => { mounted = false; };
  }, []);

  const processQueue = useCallback(async () => {
    if (processingRef.current) return;
    processingRef.current = true;
    const claimed = new Set();
    const go = async (wi) => {
      while (true) {
        const recs = receiptsRef.current;
        const ni = recs.findIndex((r, i) => r.status === "queue" && !claimed.has(i));
        if (ni === -1 || !workersRef.current[wi]) return;
        claimed.add(ni);
        setReceipts(p => p.map((r, i) => i === ni ? { ...r, status: "scanning" } : r));
        try {
          const pp = await preprocessImage(recs[ni].imgUrl);
          const { data } = await workersRef.current[wi].recognize(pp);
          const parsed = parseReceipt(data.text);
          setReceipts(p => p.map((r, i) => i === ni ? {
            ...r, status: "done", parsed,
            editFop: parsed.fop,
            editTd: parsed.td,
            editGot: parsed.gotivka.toString(),
            editBez: parsed.kartka.toString(),
            editMonth: parsed.monthIdx >= 0 ? parsed.monthIdx : 0,
          } : r));
        } catch {
          setReceipts(p => p.map((r, i) => i === ni ? { ...r, status: "error" } : r));
        }
      }
    };
    await Promise.all(workersRef.current.map((_, i) => go(i)));
    processingRef.current = false;
  }, []);

  useEffect(() => {
    if (ocrReady && receipts.some(r => r.status === "queue")) processQueue();
  }, [receipts, ocrReady, processQueue]);

  useEffect(() => {
    if (receipts.length === 0) { allDoneRef.current = false; return; }
    const allProcessed = receipts.every(r => r.status === "done" || r.status === "error" || r.status === "pending");
    const hasActive = receipts.some(r => r.status === "queue" || r.status === "scanning");
    const hasDone = receipts.some(r => r.status === "done");
    if (allProcessed && hasDone && !allDoneRef.current) {
      allDoneRef.current = true;
      const dc = receipts.filter(r => r.status === "done").length;
      notify(`${m.allDone} (${dc} ✓)`);
    }
    if (hasActive) allDoneRef.current = false;
  }, [receipts, m.allDone]);

  const handleFiles = useCallback((files) => {
    const nr = [];
    Array.from(files).forEach((f, i) => {
      if (f.type.startsWith("image/"))
        nr.push({ id: Date.now() + i, imgUrl: URL.createObjectURL(f), name: f.name, status: "pending", parsed: null, editFop: "", editTd: "", editGot: "", editBez: "", editMonth: 0 });
    });
    if (nr.length) { setReceipts(p => [...p, ...nr]); notify(`${nr.length} ${m.photosAdded} +`); }
  }, [m]);

  const startOcr = () => {
    if (!ocrReady) return;
    allDoneRef.current = false;
    setReceipts(p => p.map(r => r.status === "pending" ? { ...r, status: "queue" } : r));
  };

  const deleteReceipt = (idx) => {
    const name = receipts[idx]?.name;
    setReceipts(p => p.filter((_, i) => i !== idx));
    setActiveIdx(prev => {
      const newLen = receipts.length - 1;
      if (newLen <= 0) return 0;
      if (idx < prev) return prev - 1;
      if (idx === prev) return Math.min(prev, newLen - 1);
      return prev;
    });
    if (name) notify(`${name} ${m.deleted}`);
  };

  const upd = (i, f, v) => setReceipts(p => p.map((r, j) => j === i ? { ...r, [f]: v } : r));
  const applyPrev = (pair) => setReceipts(p => p.map((r, j) => j === activeIdx ? { ...r, editFop: pair.fop, editTd: pair.td } : r));

  /* ─── Table Data & Export ─── */
  const buildTableData = () => {
    const map = new Map();
    receipts.filter(r => r.status === "done").forEach(r => {
      const fop = r.editFop.trim();
      const td = r.editTd.trim();
      if (!fop) return;
      const key = `${fop}|||${td}`;
      if (!map.has(key)) map.set(key, { fop, td, months: {} });
      const entry = map.get(key);
      const mo = r.editMonth;
      if (!entry.months[mo]) entry.months[mo] = { g: 0, b: 0 };
      entry.months[mo].g += parseNum(r.editGot);
      entry.months[mo].b += parseNum(r.editBez);
    });
    return [...map.values()];
  };

  const exportXlsx = () => {
    const UA = LANGS.ua;
    const tableData = buildTableData();
    if (!tableData.length) { notify(m.noData, "error"); return; }

    const h1 = ["", ""];
    const h2 = ["\u0424\u041E\u041F", "\u0422\u0414"];
    UA.months.forEach(mo => { h1.push(mo, "", ""); h2.push("\u0413\u043E\u0442\u0456\u0432\u043A\u0430", "\u0411\u0435\u0437\u0433\u043E\u0442\u0456\u0432\u043A\u0430", "\u0412\u0441\u044C\u043E\u0433\u043E"); });
    h1.push(""); h2.push("");
    h1.push("2026 \u0440", "", ""); h2.push("\u0413\u043E\u0442\u0456\u0432\u043A\u0430", "\u0411\u0435\u0437\u0433\u043E\u0442\u0456\u0432\u043A\u0430", "\u0412\u0441\u044C\u043E\u0433\u043E");
    h1.push("\u0437\u0430\u0433\u0430\u043B\u044C\u043D\u0430 \u0441\u0443\u043C\u0430 2026"); h2.push("");

    const rows = [h1, h2];
    tableData.forEach(f => {
      const row = [f.fop, f.td];
      let yG = 0, yB = 0;
      for (let mo = 0; mo < 12; mo++) {
        const e = f.months[mo];
        const g = e?.g || 0, b = e?.b || 0;
        row.push(g || undefined, b || undefined, (g + b) || 0);
        yG += g; yB += b;
      }
      row.push(undefined);
      row.push(yG, yB, yG + yB);
      row.push(yG + yB);
      rows.push(row);
    });

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [{ wch: 22 }, { wch: 18 }];
    for (let i = 0; i < 12; i++) ws["!cols"].push({ wch: 14 }, { wch: 14 }, { wch: 14 });
    ws["!cols"].push({ wch: 2 }, { wch: 14 }, { wch: 14 }, { wch: 14 }, { wch: 18 });

    const mg = [];
    for (let i = 0; i < 12; i++) { const c = 2 + i * 3; mg.push({ s: { r: 0, c }, e: { r: 0, c: c + 2 } }); }
    mg.push({ s: { r: 0, c: 39 }, e: { r: 0, c: 41 } });
    mg.push({ s: { r: 0, c: 42 }, e: { r: 1, c: 42 } });
    ws["!merges"] = mg;

    /* ─── Cell Styling ─── */
    const thin = { style: "thin", color: { rgb: "D0D0D0" } };
    const border = { top: thin, bottom: thin, left: thin, right: thin };
    const headerFill = { fgColor: { rgb: "4472C4" } };
    const subHeaderFill = { fgColor: { rgb: "D9E2F3" } };
    const stripeFill = { fgColor: { rgb: "F2F2F2" } };
    const yearHeaderFill = { fgColor: { rgb: "E2EFDA" } };
    const grandFill = { fgColor: { rgb: "E2EFDA" } };
    const headerFont = { bold: true, color: { rgb: "FFFFFF" }, sz: 10, name: "Calibri" };
    const subHeaderFont = { bold: true, color: { rgb: "333333" }, sz: 9, name: "Calibri" };
    const bodyFont = { sz: 9, name: "Calibri", color: { rgb: "333333" } };
    const numFont = { sz: 9, name: "Calibri", color: { rgb: "333333" } };
    const yearFont = { bold: true, sz: 9, name: "Calibri", color: { rgb: "375623" } };
    const totalCols = rows[0].length;

    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < totalCols; c++) {
        const a = XLSX.utils.encode_cell({ r, c });
        if (!ws[a]) ws[a] = { v: "", t: "s" };
        const cell = ws[a];

        if (r === 0) {
          // Top header row
          const isYear = c >= 39 && c <= 41;
          cell.s = {
            fill: isYear ? yearHeaderFill : headerFill,
            font: isYear ? { bold: true, color: { rgb: "375623" }, sz: 10, name: "Calibri" } : headerFont,
            border,
            alignment: { horizontal: c < 2 ? "left" : "center", vertical: "center", wrapText: true },
          };
          if (c === 42) cell.s.fill = yearHeaderFill;
        } else if (r === 1) {
          // Sub header row
          cell.s = {
            fill: subHeaderFill,
            font: subHeaderFont,
            border,
            alignment: { horizontal: c < 2 ? "left" : "center", vertical: "center" },
          };
        } else {
          // Data rows
          const isStripe = (r - 2) % 2 === 1;
          const isNum = c >= 2 && typeof cell.v === "number";
          const isYearCol = c >= 39 && c <= 41;
          const isGrand = c === 42;

          if (isNum && cell.v !== 0) { cell.t = "n"; cell.z = "#,##0.00"; }

          cell.s = {
            fill: isGrand ? grandFill : isStripe ? stripeFill : { fgColor: { rgb: "FFFFFF" } },
            font: isGrand ? yearFont : isYearCol ? yearFont : isNum ? numFont : bodyFont,
            border,
            alignment: { horizontal: isNum || isYearCol || isGrand ? "right" : "left", vertical: "center" },
          };
        }
      }
    }

    ws["!rows"] = [{ hpt: 22 }, { hpt: 18 }];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "\u0432 \u0440\u043E\u0437\u0440\u0456\u0437\u0456 \u0424\u041E\u041F");
    XLSX.writeFile(wb, "\u041F\u043E\u0434\u0430\u0442\u043A\u043E\u0432\u0435_\u043D\u0430\u0432\u0430\u043D\u0442\u0430\u0436\u0435\u043D\u043D\u044F_2026.xlsx");
    notify(m.exported);
  };

  /* ─── Derived ─── */
  const ar = receipts[activeIdx];
  const sc = (s) => s === "done" ? "var(--success)" : s === "pending" ? "var(--accent)" : s === "scanning" ? "var(--info)" : s === "error" ? "var(--error)" : "var(--text-dim)";
  const tableData = buildTableData();

  return (
    <div className="app-root">
      {/* ─── Toast ─── */}
      {notif && <div className={`toast toast--${notif.type}`}>{notif.msg}</div>}

      {/* ─── Header ─── */}
      <header className="app-header">
        <div className="header-brand">
          <div className="header-logo">{"🧾"}</div>
          <div>
            <h1 className="header-title">{t.title}</h1>
            <p className="header-subtitle">
              {m.subtitle}
              {receipts.length > 0 && <span> {"·"} {doneC}/{receipts.length} {"✓"}</span>}
              {pendingCount > 0 && <span className="stat-pending"> {"·"} {pendingCount} {"⏳"}</span>}
              {scanC > 0 && <span className="stat-scanning"> {"·"} {scanC} {"..."}</span>}
              {ocrReady && receipts.length === 0 && <span className="stat-ready"> {"·"} OCR {"✓"}</span>}
            </p>
          </div>
        </div>

        <div className="header-actions">
          {/* Language */}
          <div className="btn-group">
            {Object.entries(LANGS).map(([k]) => (
              <button key={k} onClick={() => setLang(k)} className={`lang-btn${lang === k ? " lang-btn--active" : ""}`} data-tooltip={k === "ua" ? "Українська" : k === "tr" ? "Türkçe" : "English"} data-tooltip-pos="bottom">
                {k === "ua" ? "🇺🇦" : k === "tr" ? "🇹🇷" : "🇬🇧"}
              </button>
            ))}
          </div>

          {/* Theme Toggle */}
          <button onClick={toggleTheme} className="theme-toggle" data-tooltip={t.tip.theme} data-tooltip-pos="bottom">
            {theme === "dark" ? "☀️" : "🌙"}
          </button>

          {/* View Tabs */}
          <div className="btn-group">
            <button onClick={() => setView("entry")} className={`tab${view === "entry" ? " tab--active" : ""}`} data-tooltip={t.tip.entry} data-tooltip-pos="bottom">
              <span>{"📝"}</span> {t.entry}
            </button>
            <button onClick={() => setView("table")} className={`tab${view === "table" ? " tab--active" : ""}`} data-tooltip={t.tip.table} data-tooltip-pos="bottom">
              <span>{"📊"}</span> {t.table}
              {doneC > 0 && <span className="badge">{tableData.length}</span>}
            </button>
          </div>

          {/* Export */}
          <button onClick={exportXlsx} className="btn btn--success" disabled={scanC > 0 || pendingCount > 0} data-tooltip={t.tip.exportXls} data-tooltip-pos="bottom-end">
            {"⬇"} {t.export}
          </button>
        </div>
      </header>

      {view === "entry" ? (
        <div className="split-view">
          {/* ─── Left: Image ─── */}
          <div className="split-left">
            {receipts.length === 0 ? (
              <div
                className={`dropzone${dragOver ? " dropzone--active" : ""}`}
                onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
                onDrop={e => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
                onClick={() => fileRef.current?.click()}
              >
                <div className="dropzone__icon">{"📸"}</div>
                <p className="dropzone__title">{m.drop}</p>
                <p className="dropzone__sub">{m.dropSub}</p>
                <div className={`dropzone__status dropzone__status--${ocrReady ? "ready" : "loading"}`}>
                  {ocrReady ? m.ocrReady : m.ocrLoading}
                  {!ocrReady && <span className="scan-spinner scan-spinner--sm scan-spinner--inline" />}
                </div>
                <input ref={fileRef} type="file" multiple accept="image/*" style={{ display: "none" }} onChange={e => handleFiles(e.target.files)} />
              </div>
            ) : (
              <div className="viewer">
                {/* Toolbar */}
                <div className="viewer__toolbar">
                  <div className="viewer__toolbar-left">
                    <button onClick={() => setActiveIdx(Math.max(0, activeIdx - 1))} disabled={activeIdx === 0} className="btn btn--icon btn--sm" data-tooltip={t.tip.prev} data-tooltip-pos="bottom">{"◀"}</button>
                    <span className="viewer__counter">{activeIdx + 1}/{receipts.length}</span>
                    <button onClick={() => setActiveIdx(Math.min(receipts.length - 1, activeIdx + 1))} disabled={activeIdx === receipts.length - 1} className="btn btn--icon btn--sm" data-tooltip={t.tip.next} data-tooltip-pos="bottom">{"▶"}</button>
                    {ar && <span className="viewer__status" style={{ background: `color-mix(in srgb, ${sc(ar.status)} 15%, transparent)`, color: sc(ar.status) }}>{ar.status === "done" ? "✓" : ar.status === "pending" ? "●" : ar.status === "scanning" ? "..." : ar.status === "error" ? "!" : "..."}</span>}
                  </div>
                  <div className="viewer__toolbar-right">
                    <button onClick={() => setZoom(z => Math.max(.3, z - .25))} className="btn btn--icon btn--sm" data-tooltip={t.tip.zoomOut} data-tooltip-pos="bottom">{"−"}</button>
                    <button onClick={() => setZoom(z => Math.min(3, z + .25))} className="btn btn--icon btn--sm" data-tooltip={t.tip.zoomIn} data-tooltip-pos="bottom">{"+"}</button>
                    <button onClick={() => deleteReceipt(activeIdx)} className="btn btn--sm" data-tooltip={t.tip.del} data-tooltip-pos="bottom" style={{ background: "var(--error-bg)", color: "var(--error)", fontWeight: 700 }}>{"✕"}</button>
                    <button onClick={() => fileRef.current?.click()} className="btn btn--sm" data-tooltip={t.tip.addPhoto} data-tooltip-pos="bottom" style={{ background: "var(--accent-bg)", color: "var(--accent)", fontWeight: 700 }}>{"+ foto"}</button>
                    <input ref={fileRef} type="file" multiple accept="image/*" style={{ display: "none" }} onChange={e => handleFiles(e.target.files)} />
                  </div>
                </div>

                {/* Canvas */}
                <div className="viewer__canvas"
                  onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                  onDragLeave={() => setDragOver(false)}
                  onDrop={e => { e.preventDefault(); setDragOver(false); handleFiles(e.dataTransfer.files); }}
                >
                  {ar && <img src={ar.imgUrl} alt="" style={{ maxWidth: `${zoom * 100}%`, maxHeight: zoom > 1 ? "none" : "100%" }} />}
                  {ar?.status === "scanning" && (
                    <div className="scan-overlay">
                      <div className="scan-line" />
                      <div className="scan-overlay__card">
                        <div className="scan-spinner" />
                        <p className="scan-msg">{getLoadingMsg(ar.id)}</p>
                      </div>
                    </div>
                  )}
                  {ar?.status === "queue" && (
                    <div className="queue-overlay">
                      <div className="queue-overlay__inner">
                        <div className="queue-overlay__icon">{"⏳"}</div>
                        <p className="queue-overlay__msg">{m.scanning[0]}</p>
                      </div>
                    </div>
                  )}
                </div>

                {/* Scanning progress bar */}
                {scanC > 0 && (
                  <div className="viewer__progress">
                    <div className="viewer__progress-bar" style={{ width: `${(doneC / receipts.length) * 100}%` }} />
                  </div>
                )}

                {/* Thumbnails */}
                {receipts.length > 1 && (
                  <div className="thumb-strip">
                    {receipts.map((r, i) => (
                      <div key={r.id} className="thumb">
                        <img src={r.imgUrl} alt="" onClick={() => setActiveIdx(i)} className={`thumb__img${i === activeIdx ? " thumb__img--active" : ""}`} />
                        <span className="thumb__status" style={{ background: sc(r.status) }} />
                        <button onClick={(e) => { e.stopPropagation(); deleteReceipt(i); }} className="thumb__delete">{"✕"}</button>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            )}
          </div>

          {/* ─── Right: Edit Panel ─── */}
          <div className="split-right">
            {!ar ? (
              <div className="panel-center">
                <span className="panel-center__icon">{"←"}</span>
                <p className="panel-center__sub">{m.drop}</p>
              </div>
            ) : ar.status === "pending" ? (
              <div className="panel-center">
                <div className="panel-center__icon">{"📷"}</div>
                <p className="panel-center__title">{ar.name}</p>
                <p className="panel-center__sub">{m.pendingMsg}</p>

                {pendingCount > 0 && ocrReady && (
                  <button onClick={startOcr} className="btn btn--start" data-tooltip={t.tip.start}>
                    {"▶"} {m.startBtn}
                    <span className="badge">{pendingCount}</span>
                  </button>
                )}
                {!ocrReady && (
                  <div style={{ display: "flex", alignItems: "center", gap: 6, color: "var(--info)", fontSize: 11, fontWeight: 600 }}>
                    <span className="scan-spinner scan-spinner--sm" />
                    {m.ocrLoading}
                  </div>
                )}

                <button onClick={() => deleteReceipt(activeIdx)} className="btn btn--ghost btn--sm" data-tooltip={t.tip.del}>{"✕"}</button>
              </div>
            ) : ar.status === "queue" || ar.status === "scanning" ? (
              <div className="panel-center">
                <div className="scan-spinner" />
                <p style={{ color: "var(--info-light)", fontWeight: 700, fontSize: 14 }}>{getLoadingMsg(ar.id)}</p>
                <p style={{ color: "var(--text-dim)", fontSize: 10 }}>{activeIdx + 1}/{receipts.length}</p>
              </div>
            ) : ar.status === "error" ? (
              <div className="panel-center">
                <span style={{ fontSize: 32 }}>{"😕"}</span>
                <p className="panel-center__error">OCR error</p>
                <button onClick={() => deleteReceipt(activeIdx)} className="btn btn--error btn--sm" data-tooltip={t.tip.del}>{"✕"}</button>
              </div>
            ) : (
              <>
                {/* Scanning progress banner */}
                {scanC > 0 && (
                  <div className="scanning-banner">
                    <span className="scan-spinner scan-spinner--sm" />
                    <span>{m.stillScanning} {doneC}/{receipts.length}</span>
                  </div>
                )}
                {/* Auto-saved */}
                <div className="card--auto-saved">
                  <span className="check">{"✓"}</span>
                  <span className="label-text">{m.autoSaved}</span>
                  <span className="file-name">{ar.name}</span>
                </div>

                {/* FOP */}
                <div className="field">
                  <label className="label">{t.fopLabel}</label>
                  <input value={ar.editFop} onChange={e => upd(activeIdx, "editFop", e.target.value)} placeholder="..." className="input" />
                </div>

                {/* TD */}
                <div className="field">
                  <label className="label">{t.tdLabel}</label>
                  <input value={ar.editTd} onChange={e => upd(activeIdx, "editTd", e.target.value)} placeholder="..." className="input" />
                </div>

                {/* Previous suggestions */}
                {prevPairs.length > 0 && (
                  <div>
                    <span className="label">{t.prev}:</span>
                    <div className="prev-pairs">
                      {prevPairs.map((p, i) => (
                        <button key={i} onClick={() => applyPrev(p)} className="prev-pair">
                          {p.fop}{p.td ? ` — ${p.td}` : ""}
                        </button>
                      ))}
                    </div>
                  </div>
                )}

                {/* Month */}
                <div>
                  <label className="label">{t.month}</label>
                  <div className="month-grid">
                    {t.monthShort.map((mo, i) => (
                      <button key={i} onClick={() => upd(activeIdx, "editMonth", i)} className={`month-btn${ar.editMonth === i ? " month-btn--active" : ""}`}>{mo}</button>
                    ))}
                  </div>
                </div>

                {/* Amounts */}
                <div className="field-row">
                  <div>
                    <label className="label">{t.gotivka}</label>
                    <input value={ar.editGot} onChange={e => upd(activeIdx, "editGot", e.target.value)} className="input input--mono" />
                  </div>
                  <div>
                    <label className="label">{t.bezgotivka}</label>
                    <input value={ar.editBez} onChange={e => upd(activeIdx, "editBez", e.target.value)} className="input input--mono" />
                  </div>
                </div>

                {/* Total */}
                <div className="card--total">
                  <span className="total-label">{t.vsogo}</span>
                  <span className="total-value">{fmt(parseNum(ar.editGot) + parseNum(ar.editBez))} {"₴"}</span>
                </div>

                {/* Raw text */}
                <button onClick={() => setShowRaw(!showRaw)} className="raw-toggle" data-tooltip={t.tip.raw}>
                  {showRaw ? "▼" : "▶"} {t.rawText}
                </button>
                {showRaw && <pre className="raw-pre">{ar.parsed?.rawText || ""}</pre>}
              </>
            )}
          </div>
        </div>
      ) : (
        /* ─── Table View ─── */
        <div className="table-view">
          {tableData.length === 0 ? (
            <div className="empty-state">
              <span className="empty-state__icon">{"📊"}</span>
              <p>{m.noData}</p>
            </div>
          ) : (
            <div className="data-table-wrap">
              <table className="data-table">
                <thead>
                  <tr>
                    <th className="th-top" style={{ textAlign: "left" }} rowSpan={2}>{t.fopLabel}</th>
                    <th className="th-top" style={{ textAlign: "left" }} rowSpan={2}>{t.tdLabel}</th>
                    {t.monthShort.map((mo, i) => <th key={i} className="th-top" style={{ textAlign: "center" }} colSpan={3}>{mo}</th>)}
                    <th className="th-top th-year" style={{ textAlign: "center" }} colSpan={3}>{t.year}</th>
                  </tr>
                  <tr>
                    {[...Array(13)].map((_, i) => [t.gotShort, t.bezShort, t.vsoShort].map((h, j) => (
                      <th key={`${i}${j}`} className={`th-sub${j === 2 ? " th-divider" : ""}`} style={{ textAlign: "center" }}>{h}</th>
                    )))}
                  </tr>
                </thead>
                <tbody>
                  {tableData.map((f, fi) => {
                    let yG = 0, yB = 0;
                    return (
                      <tr key={fi}>
                        <td className="td-fop">{f.fop}</td>
                        <td className="td-td">{f.td}</td>
                        {[...Array(12)].map((_, mo) => {
                          const e = f.months[mo];
                          const g = e?.g || 0, b = e?.b || 0;
                          yG += g; yB += b;
                          return [
                            <td key={`g${fi}${mo}`} className="td-num">{g ? fmt(g) : ""}</td>,
                            <td key={`b${fi}${mo}`} className="td-num">{b ? fmt(b) : ""}</td>,
                            <td key={`v${fi}${mo}`} className="td-total td-total--divider">{(g + b) ? fmt(g + b) : ""}</td>
                          ];
                        })}
                        <td className="td-year">{yG ? fmt(yG) : ""}</td>
                        <td className="td-year">{yB ? fmt(yB) : ""}</td>
                        <td className="td-grand">{(yG + yB) ? fmt(yG + yB) : ""}</td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}
        </div>
      )}
    </div>
  );
}
