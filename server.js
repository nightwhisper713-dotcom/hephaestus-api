// ═══════════════════════════════════════════════════════════
// HEPHAESTUS API — 文件鍛造服務
// 支援：PPT / Excel（多檔合併）/ Word
// 部署：Render 免費方案
// ═══════════════════════════════════════════════════════════

const express  = require('express');
const cors     = require('cors');
const multer   = require('multer');
const XLSX     = require('xlsx');
const PptxGenJS = require('pptxgenjs');
const { Document, Packer, Paragraph, TextRun, HeadingLevel,
        Table, TableRow, TableCell, AlignmentType, WidthType,
        BorderStyle, ShadingType, Header, Footer, PageNumber,
        TableOfContents } = require('docx');
const archiver = require('archiver');
const path     = require('path');
const fs       = require('fs');
const os       = require('os');
let sharp;
try { sharp = require('sharp'); } catch(e) { sharp = null; console.log('sharp not available, skipping image conversion'); }

const app = express();
app.use(cors());
app.use(express.json({ limit: '50mb' }));

// Multer — 記憶體儲存，支援多檔
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 20 * 1024 * 1024 } // 20MB per file
});

// ── 健康檢查 ───────────────────────────────────────────────
app.get('/', (req, res) => res.json({
  service: 'HEPHAESTUS API',
  version: '1.0.0',
  endpoints: ['/generate/ppt', '/generate/excel', '/generate/word', '/generate/merge-excel', '/generate/bundle']
}));

// ═══════════════════════════════════════════════════════════
// PPT 生成
// POST /generate/ppt
// Body: { filename, theme?, slides: [{type, title, bullets?, ...}] }
// ═══════════════════════════════════════════════════════════
app.post('/generate/ppt', async (req, res) => {
  try {
    const { filename = '簡報', theme = {}, slides = [] } = req.body;
    const buf = await buildPptx(filename, theme, slides);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}.pptx`,
      'Content-Length': buf.length
    });
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════
// Excel 生成
// POST /generate/excel
// Body: { filename, sheets: [{name, headers, rows, widths?, formulas?}] }
// ═══════════════════════════════════════════════════════════
app.post('/generate/excel', async (req, res) => {
  try {
    const { filename = '試算表', sheets = [] } = req.body;
    const buf = buildExcel(filename, sheets);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}.xlsx`,
      'Content-Length': buf.length
    });
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════
// Excel 多檔合併
// POST /generate/merge-excel   (multipart/form-data)
// Fields: files[] (xlsx), mode (stack|sheet), outputName
// ═══════════════════════════════════════════════════════════
app.post('/generate/merge-excel', upload.array('files', 20), async (req, res) => {
  try {
    const { mode = 'sheet', outputName = '合併試算表' } = req.body;
    if (!req.files || req.files.length === 0) throw new Error('未收到任何檔案');

    const wb = XLSX.utils.book_new();
    const allData = []; // for stack mode

    req.files.forEach((file, idx) => {
      const srcWb = XLSX.read(file.buffer, { type: 'buffer', cellStyles: true, cellDates: true });

      if (mode === 'sheet') {
        // 每個檔案的每個工作表獨立放進新工作簿
        srcWb.SheetNames.forEach(name => {
          const ws = srcWb.Sheets[name];
          const sheetName = req.files.length > 1
            ? `${path.basename(file.originalname, path.extname(file.originalname))}_${name}`.slice(0, 31)
            : name.slice(0, 31);
          XLSX.utils.book_append_sheet(wb, ws, sheetName);
        });
      } else {
        // stack mode：把所有第一個工作表的資料垂直疊加
        const ws = srcWb.Sheets[srcWb.SheetNames[0]];
        const data = XLSX.utils.sheet_to_json(ws, { header: 1 });
        if (idx === 0) {
          allData.push(...data); // 保留標題列
        } else {
          allData.push(...data.slice(1)); // 跳過標題列
        }
      }
    });

    if (mode === 'stack') {
      const ws = XLSX.utils.aoa_to_sheet(allData);
      // 自動欄寬
      const cols = allData[0] ? allData[0].map((_, i) => ({
        wch: Math.max(...allData.map(row => String(row[i] || '').length), 8)
      })) : [];
      ws['!cols'] = cols;
      // 凍結首列
      ws['!freeze'] = { xSplit: 0, ySplit: 1 };
      XLSX.utils.book_append_sheet(wb, ws, '合併資料');

      // 加總覽工作表
      const summaryWs = XLSX.utils.aoa_to_sheet([
        ['統計項目', '數值'],
        ['來源檔案數', req.files.length],
        ['合計資料列數', allData.length - 1],
        ['欄位數', allData[0] ? allData[0].length : 0],
        ['產生時間', new Date().toLocaleString('zh-TW')]
      ]);
      summaryWs['!cols'] = [{ wch: 16 }, { wch: 24 }];
      XLSX.utils.book_append_sheet(wb, summaryWs, '合併總覽');
    }

    const buf = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}.xlsx`
    });
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════
// Word 生成
// POST /generate/word
// Body: { filename, sections: [{type, text, items?, headers?, rows?}] }
// ═══════════════════════════════════════════════════════════
app.post('/generate/word', async (req, res) => {
  try {
    const { filename = '文件', sections = [], brand = '幻翔商用設計' } = req.body;
    const buf = await buildWord(filename, sections, brand);
    res.set({
      'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(filename)}.docx`
    });
    res.send(buf);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════
// 打包下載（同時生成多種格式，回傳 ZIP）
// POST /generate/bundle
// Body: { ppt?, excel?, word?, zipName? }
// ═══════════════════════════════════════════════════════════
app.post('/generate/bundle', async (req, res) => {
  try {
    const { ppt, excel, word, zipName = '幻翔文件包' } = req.body;
    res.set({
      'Content-Type': 'application/zip',
      'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(zipName)}.zip`
    });
    const archive = archiver('zip', { zlib: { level: 6 } });
    archive.pipe(res);

    if (ppt) {
      const buf = await buildPptx(ppt.filename || '簡報', ppt.theme || {}, ppt.slides || []);
      archive.append(buf, { name: `${ppt.filename || '簡報'}.pptx` });
    }
    if (excel) {
      const buf = buildExcel(excel.filename || '試算表', excel.sheets || []);
      archive.append(buf, { name: `${excel.filename || '試算表'}.xlsx` });
    }
    if (word) {
      const buf = await buildWord(word.filename || '文件', word.sections || []);
      archive.append(buf, { name: `${word.filename || '文件'}.docx` });
    }
    await archive.finalize();
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ═══════════════════════════════════════════════════════════
// PPT 建構函數
// ═══════════════════════════════════════════════════════════
async function buildPptx(filename, theme, slides) {
  const pres = new PptxGenJS();
  pres.layout = 'LAYOUT_WIDE';
  pres.author = '幻翔商用設計 · HEPHAESTUS API';

  const T = {
    bg:      theme.bg      || '1A1614',
    gold:    theme.gold    || 'B8956A',
    text:    theme.text    || 'F0EBE2',
    accent:  theme.accent  || '5FA878',
    light:   theme.light   || 'F7F4F0',
    dark:    theme.dark    || '1A1614',
    white:   'FFFFFF',
    gray:    '8A8078',
    grayL:   'C8C0B8',
  };
  const W = 13.3, H = 7.5;
  const mk = () => ({ type: 'outer', color: '000000', blur: 8, offset: 2, angle: 45, opacity: 0.10 });

  slides.forEach(slide => {
    const sl = pres.addSlide();

    if (slide.type === 'title') {
      sl.background = { color: T.bg };
      sl.addShape(pres.shapes.OVAL, { x: 9, y: -1, w: 6, h: 6, fill: { color: T.gold, transparency: 88 }, line: { color: T.gold, transparency: 88 } });
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.06, fill: { color: T.gold }, line: { color: T.gold } });
      sl.addText(slide.title || '', { x: 1, y: 1.8, w: 11.3, h: 1.2, fontSize: 44, bold: true, color: T.gold, fontFace: 'Cambria', align: 'center', margin: 0 });
      if (slide.subtitle) sl.addText(slide.subtitle, { x: 1, y: 3.15, w: 11.3, h: 0.7, fontSize: 22, color: T.text, fontFace: 'Arial', align: 'center', italic: true, margin: 0 });
      if (slide.date)  sl.addText(slide.date,  { x: 1, y: 4.0,  w: 11.3, h: 0.4, fontSize: 13, color: T.gray, fontFace: 'Arial', align: 'center', margin: 0 });
      if (slide.brand) sl.addText(slide.brand, { x: 0, y: 7.1,  w: W,    h: 0.3, fontSize: 10, color: T.gray, fontFace: 'Arial', align: 'center', margin: 0 });

    } else if (slide.type === 'content') {
      sl.background = { color: T.light };
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.1, fill: { color: T.dark }, line: { color: T.dark } });
      sl.addText(slide.title || '', { x: 0.6, y: 0.22, w: 12, h: 0.68, fontSize: 26, bold: true, color: T.gold, fontFace: 'Cambria', margin: 0 });
      if (slide.subtitle) sl.addText(slide.subtitle, { x: 0.6, y: 0.82, w: 12, h: 0.25, fontSize: 10, color: T.grayL, fontFace: 'Arial', italic: true, margin: 0 });
      const bullets = slide.bullets || [];
      bullets.forEach((b, i) => {
        sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.45, y: 1.25 + i * 0.92, w: 12.4, h: 0.78, fill: { color: T.white }, line: { color: 'E8E0D8', width: 0.5 }, rectRadius: 0.06, shadow: mk() });
        sl.addText(b, { x: 0.65, y: 1.25 + i * 0.92, w: 12.1, h: 0.78, fontSize: 15, color: T.dark, fontFace: 'Arial', valign: 'middle', margin: 0 });
      });

    } else if (slide.type === 'stats') {
      sl.background = { color: T.bg };
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.1, fill: { color: '120F0D' }, line: { color: '120F0D' } });
      sl.addText(slide.title || '', { x: 0.6, y: 0.22, w: 12, h: 0.68, fontSize: 26, bold: true, color: T.gold, fontFace: 'Cambria', margin: 0 });
      const stats = slide.stats || [];
      const cols = Math.min(stats.length, 4);
      const cardW = (W - 0.4 * (cols + 1)) / cols;
      const statColors = [T.accent, T.gold, '6A8FD4', 'D46A6A', 'C9A84C'];
      stats.forEach((s, i) => {
        const x = 0.4 + i * (cardW + 0.4);
        sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 1.28, w: cardW, h: 2.2, fill: { color: '1E1A18' }, line: { color: statColors[i % 5], width: 1.5 }, rectRadius: 0.14, shadow: mk() });
        sl.addText(s.value || '', { x, y: 1.4, w: cardW, h: 0.9, fontSize: 46, bold: true, color: statColors[i % 5], fontFace: 'Cambria', align: 'center', margin: 0 });
        sl.addText(s.label || '', { x: x + 0.1, y: 2.38, w: cardW - 0.2, h: 0.65, fontSize: 14, color: T.text, fontFace: 'Arial', align: 'center', margin: 0 });
        if (s.sub) sl.addText(s.sub, { x: x + 0.1, y: 3.08, w: cardW - 0.2, h: 0.32, fontSize: 11, color: T.gray, fontFace: 'Arial', align: 'center', italic: true, margin: 0 });
      });
      if (slide.note) sl.addText(slide.note, { x: 0.4, y: 3.6, w: 12.5, h: 0.4, fontSize: 12, color: T.gold, fontFace: 'Arial', align: 'center', italic: true, margin: 0 });

    } else if (slide.type === 'two_col') {
      sl.background = { color: T.light };
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.1, fill: { color: T.dark }, line: { color: T.dark } });
      sl.addText(slide.title || '', { x: 0.6, y: 0.22, w: 12, h: 0.68, fontSize: 26, bold: true, color: T.gold, fontFace: 'Cambria', margin: 0 });
      const leftItems = slide.left || [];
      const rightItems = slide.right || [];
      const leftLabel = slide.leftLabel || '左欄';
      const rightLabel = slide.rightLabel || '右欄';
      // Left header
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 1.22, w: 5.9, h: 0.5, fill: { color: T.gold }, line: { color: T.gold }, rectRadius: 0.06 });
      sl.addText(leftLabel, { x: 0.4, y: 1.22, w: 5.9, h: 0.5, fontSize: 14, bold: true, color: T.dark, fontFace: 'Arial', align: 'center', margin: 0 });
      leftItems.forEach((item, i) => {
        sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.4, y: 1.85 + i * 0.82, w: 5.9, h: 0.7, fill: { color: T.white }, line: { color: 'E8E0D8', width: 0.5 }, rectRadius: 0.06, shadow: mk() });
        sl.addText(item, { x: 0.55, y: 1.85 + i * 0.82, w: 5.6, h: 0.7, fontSize: 13, color: T.dark, fontFace: 'Arial', valign: 'middle', margin: 0 });
      });
      // Right header
      sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.0, y: 1.22, w: 5.9, h: 0.5, fill: { color: '6A8FD4' }, line: { color: '6A8FD4' }, rectRadius: 0.06 });
      sl.addText(rightLabel, { x: 7.0, y: 1.22, w: 5.9, h: 0.5, fontSize: 14, bold: true, color: T.white, fontFace: 'Arial', align: 'center', margin: 0 });
      rightItems.forEach((item, i) => {
        sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.0, y: 1.85 + i * 0.82, w: 5.9, h: 0.7, fill: { color: T.white }, line: { color: 'E8E0D8', width: 0.5 }, rectRadius: 0.06, shadow: mk() });
        sl.addText(item, { x: 7.15, y: 1.85 + i * 0.82, w: 5.6, h: 0.7, fontSize: 13, color: T.dark, fontFace: 'Arial', valign: 'middle', margin: 0 });
      });

    } else if (slide.type === 'table') {
      sl.background = { color: T.light };
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.1, fill: { color: T.dark }, line: { color: T.dark } });
      sl.addText(slide.title || '', { x: 0.6, y: 0.22, w: 12, h: 0.68, fontSize: 26, bold: true, color: T.gold, fontFace: 'Cambria', margin: 0 });
      const headers = slide.headers || [];
      const rows    = slide.rows    || [];
      const nCols   = headers.length || (rows[0] ? rows[0].length : 1);
      const colW    = slide.colWidths || headers.map(() => 12.5 / nCols);
      const tData   = [
        headers.map(h => ({ text: h, options: { bold: true, fontSize: 11, fontFace: 'Arial', color: T.white, fill: { color: T.dark }, align: 'center', valign: 'middle' } })),
        ...rows.map((row, ri) => row.map((cell, ci) => ({
          text: String(cell ?? ''),
          options: { fontSize: 12, fontFace: 'Arial', color: ci === 0 ? T.dark : '3A3028', fill: { color: ri % 2 === 0 ? 'FAFAF8' : T.white }, align: ci === 0 ? 'left' : 'center', valign: 'middle' }
        })))
      ];
      const tH = Math.min(1.1 + (rows.length + 1) * 0.65, 6.1);
      sl.addTable(tData, { x: 0.4, y: 1.22, w: 12.5, h: tH, colW, border: { pt: 0.5, color: 'E0D8D0' }, rowH: 0.58 });

    } else if (slide.type === 'chart') {
      // 圖表版型：左側圖表 + 右側重點卡片
      sl.background = { color: T.bg };
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.1, fill: { color: '120F0D' }, line: { color: '120F0D' } });
      sl.addText(slide.title || '', { x: 0.6, y: 0.22, w: 12, h: 0.68, fontSize: 26, bold: true, color: T.gold, fontFace: 'Cambria', margin: 0 });

      const chartType = slide.chartType || 'bar'; // bar, line, pie, doughnut
      const chartData = slide.chartData || { labels: [], values: [] };
      const chartColors = ['B8956A', '5FA878', '6A8FD4', 'D46A6A', 'C9A84C', 'A0B4C8'];

      const pptChartType = {
        bar: pres.charts.BAR,
        line: pres.charts.LINE,
        pie: pres.charts.PIE,
        doughnut: pres.charts.DOUGHNUT,
        column: pres.charts.BAR,
      }[chartType] || pres.charts.BAR;

      const chartW = slide.notes ? 7.8 : 12.5;
      sl.addChart(pptChartType, [{
        name: slide.chartLabel || '數據',
        labels: chartData.labels || [],
        values: chartData.values || [],
      }], {
        x: 0.4, y: 1.25, w: chartW, h: 5.95,
        barDir: chartType === 'column' ? 'col' : 'bar',
        chartColors,
        chartArea: { fill: { color: '1E1A18' }, roundedCorners: true },
        catAxisLabelColor: T.grayL,
        valAxisLabelColor: T.grayL,
        valGridLine: { color: '2E2A28', size: 0.5 },
        catGridLine: { style: 'none' },
        showValue: true,
        dataLabelColor: T.white,
        showLegend: false,
        lineSize: chartType === 'line' ? 3 : undefined,
        lineSmooth: chartType === 'line' ? true : undefined,
        showPercent: (chartType === 'pie' || chartType === 'doughnut'),
        legendColor: T.grayL,
        legendPos: (chartType === 'pie' || chartType === 'doughnut') ? 'r' : undefined,
      });

      // 右側重點（選用）
      if (slide.notes && Array.isArray(slide.notes)) {
        const noteColors = ['B8956A', '5FA878', '6A8FD4'];
        slide.notes.forEach((n, i) => {
          const y = 1.25 + i * 2.0;
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 8.5, y, w: 4.4, h: 1.75, fill: { color: '1E1A18' }, line: { color: noteColors[i % 3], width: 1.5 }, rectRadius: 0.14, shadow: mk() });
          if (n.value) sl.addText(n.value, { x: 8.5, y: y + 0.1, w: 4.4, h: 0.7, fontSize: 36, bold: true, color: noteColors[i % 3], fontFace: 'Cambria', align: 'center', margin: 0 });
          if (n.label) sl.addText(n.label, { x: 8.6, y: y + 0.85, w: 4.2, h: 0.65, fontSize: 13, color: T.text, fontFace: 'Arial', align: 'center', margin: 0 });
        });
      }

    } else if (slide.type === 'cards') {
      // 卡片矩陣版型：標題 + 2×2 或 1×3 卡片格
      sl.background = { color: T.light };
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 1.1, fill: { color: T.dark }, line: { color: T.dark } });
      sl.addText(slide.title || '', { x: 0.6, y: 0.22, w: 12, h: 0.68, fontSize: 26, bold: true, color: T.gold, fontFace: 'Cambria', margin: 0 });

      const cards = slide.cards || [];
      const cardColors = ['B8956A', '5FA878', '6A8FD4', 'D46A6A', 'C9A84C', 'A0B4C8'];
      const cols = cards.length <= 3 ? cards.length : 2;
      const rows = Math.ceil(cards.length / cols);
      const cW = (W - 0.4 * (cols + 1)) / cols;
      const cH = (5.85 - 0.3 * (rows + 1)) / rows;

      cards.forEach((card, i) => {
        const col = i % cols;
        const row = Math.floor(i / cols);
        const x = 0.4 + col * (cW + 0.4);
        const y = 1.25 + row * (cH + 0.3);
        const color = cardColors[i % cardColors.length];

        sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cW, h: cH, fill: { color: T.white }, line: { color: 'E8E0D8', width: 0.5 }, rectRadius: 0.12, shadow: mk() });
        // 頂部色條
        sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: cW, h: 0.28, fill: { color: color }, line: { color: color }, rectRadius: 0.06 });
        sl.addShape(pres.shapes.RECTANGLE, { x, y: y + 0.14, w: cW, h: 0.14, fill: { color: color }, line: { color: color } });

        if (card.icon) sl.addText(card.icon, { x, y: y + 0.38, w: cW, h: 0.55, fontSize: 24, align: 'center', margin: 0 });
        if (card.title) sl.addText(card.title, { x: x + 0.1, y: y + (card.icon ? 0.95 : 0.4), w: cW - 0.2, h: 0.5, fontSize: 14, bold: true, color: color, fontFace: 'Arial', align: 'center', margin: 0 });
        if (card.desc) sl.addText(card.desc, { x: x + 0.12, y: y + (card.icon ? 1.5 : 0.98), w: cW - 0.24, h: cH - (card.icon ? 1.65 : 1.1), fontSize: 12, color: T.dark, fontFace: 'Arial', align: 'center', margin: 0 });
      });

    } else if (slide.type === 'closing') {
      sl.background = { color: T.bg };
      sl.addShape(pres.shapes.OVAL, { x: 8, y: -0.8, w: 6.5, h: 6.5, fill: { color: T.gold, transparency: 90 }, line: { color: T.gold, transparency: 90 } });
      sl.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: W, h: 0.06, fill: { color: T.gold }, line: { color: T.gold } });
      sl.addText(slide.title || '', { x: 0.5, y: 1.8, w: 12.3, h: 1.1, fontSize: 48, bold: true, color: T.gold, fontFace: 'Cambria', align: 'center', margin: 0 });
      if (slide.subtitle) sl.addText(slide.subtitle, { x: 0.5, y: 3.1, w: 12.3, h: 0.7, fontSize: 22, color: T.text, fontFace: 'Arial', align: 'center', margin: 0 });
      if (slide.contact) {
        const items = Array.isArray(slide.contact) ? slide.contact : [slide.contact];
        items.forEach((c, i) => {
          const x = 0.5 + i * (12.3 / items.length);
          const w = 12.3 / items.length - 0.2;
          sl.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y: 4.1, w, h: 1.1, fill: { color: '221E1B' }, line: { color: '3A3430', width: 1 }, rectRadius: 0.12, shadow: mk() });
          sl.addText(c, { x, y: 4.12, w, h: 1.1, fontSize: 12, color: T.text, fontFace: 'Arial', align: 'center', valign: 'middle', margin: 0 });
        });
      }
      if (slide.brand) sl.addText(slide.brand, { x: 0, y: 7.1, w: W, h: 0.3, fontSize: 10, color: T.gray, fontFace: 'Arial', align: 'center', margin: 0 });
    }
  });

  return pres.write({ outputType: 'nodebuffer' });
}

// ═══════════════════════════════════════════════════════════
// Excel 建構函數（智慧試算表）
// ═══════════════════════════════════════════════════════════
function buildExcel(filename, sheets) {
  const wb = XLSX.utils.book_new();

  sheets.forEach(sheet => {
    const headers = sheet.headers || [];
    const rows    = sheet.rows    || [];
    const aoa     = [headers, ...rows];
    const ws      = XLSX.utils.aoa_to_sheet(aoa);

    // 自動欄寬
    const maxW = headers.map((h, ci) =>
      Math.max(
        String(h).length + 2,
        ...rows.map(r => String(r[ci] ?? '').length),
        6
      )
    );
    ws['!cols'] = maxW.map(w => ({ wch: Math.min(w, 40) }));

    // 凍結首列
    ws['!freeze'] = { xSplit: 0, ySplit: 1 };

    // 自動篩選
    if (headers.length > 0 && rows.length > 0) {
      ws['!autofilter'] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: rows.length, c: headers.length - 1 } }) };
    }

    // 加入公式列（如果 sheet.addSummary = true）
    if (sheet.addSummary && rows.length > 0) {
      const summaryRow = headers.map((h, ci) => {
        const colLetter = XLSX.utils.encode_col(ci);
        const isNumeric = rows.every(r => r[ci] !== null && r[ci] !== undefined && !isNaN(Number(r[ci])));
        if (isNumeric && ci > 0) {
          return { f: `SUM(${colLetter}2:${colLetter}${rows.length + 1})` };
        }
        return ci === 0 ? '合計' : '';
      });
      XLSX.utils.sheet_add_aoa(ws, [summaryRow], { origin: -1 });
    }

    // 自訂公式（sheet.formulas = [{cell, formula}]）
    if (sheet.formulas) {
      sheet.formulas.forEach(f => {
        ws[f.cell] = { t: 'n', f: f.formula };
      });
    }

    XLSX.utils.book_append_sheet(wb, ws, (sheet.name || `Sheet${sheets.indexOf(sheet)+1}`).slice(0, 31));
  });

  return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
}

// ═══════════════════════════════════════════════════════════
// Word 建構函數（品牌樣式）
// ═══════════════════════════════════════════════════════════
async function buildWord(filename, sections, brand = '幻翔商用設計') {
  const GOLD = 'B8956A';
  const DARK = '1A1614';

  const children = [];

  sections.forEach(sec => {
    if (sec.type === 'h1') {
      children.push(new Paragraph({
        children: [new TextRun({ text: sec.text, bold: true, size: 36, color: GOLD, font: 'Cambria' })],
        heading: HeadingLevel.HEADING_1,
        spacing: { before: 360, after: 180 },
        border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: GOLD, space: 4 } }
      }));
    } else if (sec.type === 'h2') {
      children.push(new Paragraph({
        children: [new TextRun({ text: sec.text, bold: true, size: 28, color: DARK, font: 'Cambria' })],
        heading: HeadingLevel.HEADING_2,
        spacing: { before: 280, after: 140 }
      }));
    } else if (sec.type === 'h3') {
      children.push(new Paragraph({
        children: [new TextRun({ text: sec.text, bold: true, size: 24, color: '3A3028', font: 'Cambria' })],
        heading: HeadingLevel.HEADING_3,
        spacing: { before: 200, after: 100 }
      }));
    } else if (sec.type === 'p') {
      children.push(new Paragraph({
        children: [new TextRun({ text: sec.text, size: 24, font: 'Arial' })],
        spacing: { before: 80, after: 80, line: 360 },
        alignment: AlignmentType.JUSTIFIED
      }));
    } else if (sec.type === 'bullet') {
      (sec.items || []).forEach(item => {
        children.push(new Paragraph({
          children: [new TextRun({ text: item, size: 24, font: 'Arial' })],
          bullet: { level: 0 },
          spacing: { before: 60, after: 60 }
        }));
      });
    } else if (sec.type === 'numbered') {
      (sec.items || []).forEach((item, i) => {
        children.push(new Paragraph({
          children: [new TextRun({ text: `${i + 1}.  ${item}`, size: 24, font: 'Arial' })],
          spacing: { before: 60, after: 60 },
          indent: { left: 400 }
        }));
      });
    } else if (sec.type === 'callout') {
      // 強調框
      children.push(new Paragraph({
        children: [new TextRun({ text: sec.text, size: 24, bold: true, color: DARK, font: 'Arial' })],
        spacing: { before: 160, after: 160 },
        indent: { left: 480, right: 480 },
        shading: { type: ShadingType.SOLID, fill: 'FFF8EE', color: 'FFF8EE' },
        border: {
          left:   { style: BorderStyle.SINGLE, size: 18, color: GOLD, space: 8 },
          top:    { style: BorderStyle.NONE },
          right:  { style: BorderStyle.NONE },
          bottom: { style: BorderStyle.NONE }
        }
      }));
    } else if (sec.type === 'table') {
      const hdrs = sec.headers || [];
      const rows = sec.rows    || [];
      const nCols = hdrs.length || (rows[0] ? rows[0].length : 1);
      const colW  = Math.floor(9360 / nCols); // twips

      const tableRows = [
        // Header row
        new TableRow({
          children: hdrs.map(h => new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 22, color: 'FFFFFF', font: 'Arial' })], alignment: AlignmentType.CENTER })],
            shading: { type: ShadingType.SOLID, fill: DARK, color: DARK },
            margins: { top: 80, bottom: 80, left: 120, right: 120 },
            width: { size: colW, type: WidthType.DXA }
          }))
        }),
        // Data rows
        ...rows.map((row, ri) => new TableRow({
          children: row.map((cell, ci) => new TableCell({
            children: [new Paragraph({ children: [new TextRun({ text: String(cell ?? ''), size: 22, font: 'Arial' })], alignment: ci === 0 ? AlignmentType.LEFT : AlignmentType.CENTER })],
            shading: { type: ShadingType.SOLID, fill: ri % 2 === 0 ? 'FAFAF8' : 'FFFFFF', color: 'auto' },
            margins: { top: 60, bottom: 60, left: 120, right: 120 },
            width: { size: colW, type: WidthType.DXA }
          }))
        }))
      ];
      children.push(new Table({
        rows: tableRows,
        width: { size: 9360, type: WidthType.DXA },
        margins: { top: 160, bottom: 160 }
      }));
      children.push(new Paragraph({ text: '', spacing: { after: 160 } }));
    } else if (sec.type === 'pagebreak') {
      children.push(new Paragraph({ pageBreakBefore: true }));
    }
  });

  const doc = new Document({
    sections: [{
      properties: {
        page: {
          margin: { top: 1440, bottom: 1440, left: 1440, right: 1440 }
        }
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            children: [
              new TextRun({ text: brand + '  ·  ', size: 18, color: GOLD, font: 'Arial' }),
              new TextRun({ text: filename, size: 18, color: '8A8078', font: 'Arial' })
            ],
            alignment: AlignmentType.RIGHT,
            border: { bottom: { style: BorderStyle.SINGLE, size: 3, color: GOLD, space: 4 } }
          })]
        })
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            children: [
              new TextRun({ text: `© ${new Date().getFullYear()} ${brand}    `, size: 16, color: '8A8078', font: 'Arial' }),
              new TextRun({ children: [PageNumber.CURRENT], size: 16, color: '8A8078', font: 'Arial' }),
              new TextRun({ text: ' / ', size: 16, color: '8A8078', font: 'Arial' }),
              new TextRun({ children: [PageNumber.TOTAL_PAGES], size: 16, color: '8A8078', font: 'Arial' }),
            ],
            alignment: AlignmentType.CENTER
          })]
        })
      },
      children
    }]
  });

  return Packer.toBuffer(doc);
}

// ═══════════════════════════════════════════════════════════
// /generate/from-file — 核心新端點
// 前端只要丟檔案 + 指令，後端包辦：
//   1. 解析各種格式（PDF/XLSX/PPTX/DOCX/圖片/文字）
//   2. 呼叫 Claude API 分析並生成結構化 JSON
//   3. 呼叫對應 builder 產出精美文件
//   4. 直接回傳二進位檔案
//
// multipart/form-data:
//   files[]   — 任意數量與格式的檔案
//   format    — 'ppt' | 'excel' | 'word'
//   instruction — 用戶指令（可選）
//   apiKey    — Claude API Key（由前端傳入，後端不儲存）
// ═══════════════════════════════════════════════════════════
app.post('/generate/from-file',
  upload.array('files', 20),
  async (req, res) => {
    const { format = 'ppt', instruction = '', apiKey } = req.body;

    if (!apiKey) {
      return res.status(400).json({ error: '需要提供 Claude API Key' });
    }
    if (!req.files || req.files.length === 0) {
      return res.status(400).json({ error: '請至少上傳一個檔案' });
    }

    try {
      // ── Step 1: 解析所有上傳的檔案，提取文字內容 ──────────
      const extractedParts = [];

      for (const file of req.files) {
        const ext = path.extname(file.originalname).toLowerCase();
        const name = file.originalname;

        if (['.xlsx', '.xls', '.csv'].includes(ext)) {
          // Excel / CSV 解析
          try {
            const wb = XLSX.read(file.buffer, { type: 'buffer', cellDates: true });
            for (const sheetName of wb.SheetNames) {
              const ws = wb.Sheets[sheetName];
              const csv = XLSX.utils.sheet_to_csv(ws);
              extractedParts.push(`【檔案：${name} / 工作表：${sheetName}】\n${csv.slice(0, 6000)}`);
            }
          } catch (e) {
            extractedParts.push(`【檔案：${name}】（Excel 解析失敗，略過）`);
          }

        } else if (['.txt', '.md', '.csv'].includes(ext)) {
          // 純文字
          const text = file.buffer.toString('utf-8').slice(0, 8000);
          extractedParts.push(`【檔案：${name}】\n${text}`);

        } else if (['.jpg', '.jpeg', '.png', '.gif', '.webp'].includes(ext)) {
          // 圖片：強制轉成標準 PNG（最相容），送給 Claude Vision
          try {
            let imgBuf = file.buffer;
            if (sharp) {
              // 用 sharp 重新編碼成乾淨的 PNG
              imgBuf = await sharp(file.buffer).png().toBuffer();
            }
            const b64 = imgBuf.toString('base64');
            extractedParts.push({ type: 'image', name, b64, mime: 'image/png' });
          } catch (imgErr) {
            // sharp 轉換失敗，直接用原始 base64
            const b64 = file.buffer.toString('base64');
            const mime = (() => {
              const raw = file.mimetype || file.originalname;
              if (raw.includes('png')) return 'image/png';
              if (raw.includes('gif')) return 'image/gif';
              if (raw.includes('webp')) return 'image/webp';
              return 'image/jpeg';
            })();
            extractedParts.push({ type: 'image', name, b64, mime });
          }

        } else if (ext === '.pdf') {
          // PDF：嘗試用 XLSX 讀取（有時PDF有嵌入表格），否則標記為需Vision
          extractedParts.push(`【檔案：${name}】（PDF 格式，已轉為 Vision 分析）`);
          const b64 = file.buffer.toString('base64');
          extractedParts.push({ type: 'pdf', name, b64 });

        } else {
          // 其他格式：嘗試當文字讀
          try {
            const text = file.buffer.toString('utf-8').slice(0, 5000);
            if (text && !text.includes('\x00')) {
              extractedParts.push(`【檔案：${name}】\n${text}`);
            } else {
              extractedParts.push(`【檔案：${name}】（二進位格式，略過文字提取）`);
            }
          } catch (e) {
            extractedParts.push(`【檔案：${name}】（無法解析）`);
          }
        }
      }

      // ── Step 2: 組合 Claude API 的 messages ──────────────
      const claudeMessages = [];
      const contentBlocks = [];
      let textContent = '';

      for (const part of extractedParts) {
        if (typeof part === 'string') {
          textContent += part + '\n\n';
        } else if (part.type === 'image') {
          if (textContent) {
            contentBlocks.push({ type: 'text', text: textContent });
            textContent = '';
          }
          contentBlocks.push({
            type: 'image',
            source: { type: 'base64', media_type: part.mime, data: part.b64 }
          });
          contentBlocks.push({ type: 'text', text: `（以上是圖片：${part.name}）` });
        } else if (part.type === 'pdf') {
          if (textContent) {
            contentBlocks.push({ type: 'text', text: textContent });
            textContent = '';
          }
          contentBlocks.push({
            type: 'document',
            source: { type: 'base64', media_type: 'application/pdf', data: part.b64 }
          });
        }
      }
      if (textContent) contentBlocks.push({ type: 'text', text: textContent });

      // 組合最終 prompt
      const fmtName = { ppt: 'PPT簡報', excel: 'Excel試算表', word: 'Word文件' };
      const schemas = {
        ppt: `{"filename":"文件名","slides":[
{"type":"title","title":"主標題","subtitle":"副標題","brand":"幻翔商用設計"},
{"type":"stats","title":"數字標題","stats":[{"value":"XX%","label":"說明","sub":"補充"},{"value":"XX","label":"說明","sub":""},{"value":"XX","label":"說明","sub":""}]},
{"type":"chart","title":"圖表標題","chartType":"bar","chartLabel":"數值","chartData":{"labels":["A","B","C","D"],"values":[80,65,50,40]},"notes":[{"value":"XX%","label":"重點1"},{"value":"XX","label":"重點2"},{"value":"XX","label":"重點3"}]},
{"type":"two_col","title":"比較標題","leftLabel":"問題/現況","left":["項目1","項目2","項目3","項目4"],"rightLabel":"解決方案/優勢","right":["方案1","方案2","方案3","方案4"]},
{"type":"cards","title":"特色標題","cards":[{"icon":"🎯","title":"特色名稱","desc":"2-3行說明文字"},{"icon":"📊","title":"特色名稱","desc":"說明"},{"icon":"⚡","title":"特色名稱","desc":"說明"},{"icon":"🔍","title":"特色名稱","desc":"說明"}]},
{"type":"closing","title":"結語","subtitle":"行動號召文字","brand":"幻翔商用設計"}]}

【強制規定】：
- 禁止使用 content 版型（純文字清單）
- 必須按順序使用 title → stats → chart → two_col → cards → closing
- stats 的 value 必須是從素材中找到的真實數字、百分比或關鍵指標
- chart 的數據必須來自素材中的比較數據或排名`,
        excel: `{"filename":"文件名","sheets":[{"name":"工作表名","headers":["欄1","欄2","欄3","欄4"],"rows":[["a","b","c","d"],["e","f","g","h"]],"addSummary":true}]}

【規定】：把素材所有數據完整整理，多個分類可建立多個 sheet`,
        word: `{"filename":"文件名","brand":"幻翔商用設計","sections":[
{"type":"h1","text":"主標題"},
{"type":"h2","text":"章節標題"},
{"type":"p","text":"段落內容"},
{"type":"bullet","items":["項目1","項目2","項目3"]},
{"type":"callout","text":"重要提示或結論"},
{"type":"table","headers":["欄1","欄2","欄3"],"rows":[["a","b","c"]]}]}

【規定】：完整保留素材所有內容，清楚分章節，重要數字用 callout 標示`
      };

      const promptText = `請分析以上所有素材，整理成精美的${fmtName[format]}。

${instruction ? `【用戶指令】${instruction}\n` : ''}
【輸出格式】只輸出純JSON，第一個字元是 { ，最後一個字元是 } ，不加任何說明或 markdown。

${schemas[format]}`;

      contentBlocks.push({ type: 'text', text: promptText });
      claudeMessages.push({ role: 'user', content: contentBlocks });

      // ── Step 3: 呼叫 Claude API ────────────────────────────
      const claudeResp = await fetch('https://api.anthropic.com/v1/messages', {
        method: 'POST',
        headers: {
          'content-type': 'application/json',
          'x-api-key': apiKey,
          'anthropic-version': '2023-06-01'
        },
        body: JSON.stringify({
          model: 'claude-sonnet-4-6',
          max_tokens: 4096,
          system: '你是 HEPHAESTUS，幻翔設計的文件鍛造師。你只輸出純JSON，第一個字元必須是 {，最後一個字元必須是 }。絕對不加任何說明文字、markdown 符號或程式碼區塊標記。',
          messages: claudeMessages
        })
      });

      if (!claudeResp.ok) {
        const errData = await claudeResp.json().catch(() => ({}));
        throw new Error(`Claude API 錯誤 ${claudeResp.status}: ${errData?.error?.message || ''}`);
      }

      const claudeData = await claudeResp.json();
      const rawJson = claudeData.content?.map(c => c.text || '').join('') || '';

      // ── Step 4: 解析 JSON ──────────────────────────────────
      let jsonData;
      try {
        let s = rawJson.replace(/^```json\s*/i, '').replace(/^```\s*/, '').replace(/```\s*$/, '').trim();
        const a = s.indexOf('{'), b = s.lastIndexOf('}');
        if (a !== -1 && b > a) s = s.slice(a, b + 1);
        jsonData = JSON.parse(s);
      } catch (e) {
        // 二次嘗試：移除換行
        try {
          let s2 = rawJson.replace(/[\r\n]/g, ' ');
          const a = s2.indexOf('{'), b = s2.lastIndexOf('}');
          if (a !== -1 && b > a) s2 = s2.slice(a, b + 1);
          jsonData = JSON.parse(s2);
        } catch (e2) {
          console.error('JSON parse failed:', rawJson.slice(0, 300));
          throw new Error('Claude 回傳格式解析失敗，請重試');
        }
      }

      // PPT 強制視覺版型：把 content 轉成 cards
      if (format === 'ppt' && jsonData.slides) {
        const icons = ['🎯','📊','⚡','🔍','💡','🚀','✅','🌟'];
        jsonData.slides = jsonData.slides.map(slide => {
          if (slide.type !== 'content') return slide;
          const bullets = slide.bullets || [];
          if (bullets.length >= 3) {
            return {
              type: 'cards', title: slide.title || '',
              cards: bullets.slice(0, 6).map((b, i) => {
                const parts = String(b).split(/[—\-：:｜|]/);
                return {
                  icon: icons[i % icons.length],
                  title: (parts[0] || '').trim().slice(0, 20) || `特色${i+1}`,
                  desc: parts.slice(1).join('').trim().slice(0, 60) || String(b).slice(0, 60)
                };
              })
            };
          }
          const half = Math.ceil(bullets.length / 2);
          return {
            type: 'two_col', title: slide.title || '',
            leftLabel: '說明', left: bullets.slice(0, half).map(b => String(b).slice(0, 50)),
            rightLabel: '內容', right: bullets.slice(half).map(b => String(b).slice(0, 50))
          };
        });
      }

      // ── Step 5: 生成精美文件 ───────────────────────────────
      let fileBuffer, contentType, fileName, fileExt;

      if (format === 'ppt') {
        fileBuffer = await buildPptx(jsonData.filename || '簡報', jsonData.theme || {}, jsonData.slides || []);
        contentType = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
        fileExt = 'pptx';
      } else if (format === 'excel') {
        fileBuffer = buildExcel(jsonData.filename || '試算表', jsonData.sheets || []);
        contentType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        fileExt = 'xlsx';
      } else if (format === 'word') {
        fileBuffer = await buildWord(jsonData.filename || '文件', jsonData.sections || [], jsonData.brand || '幻翔商用設計');
        contentType = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
        fileExt = 'docx';
      }

      fileName = `${jsonData.filename || '文件'}.${fileExt}`;

      res.set({
        'Content-Type': contentType,
        'Content-Disposition': `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`,
        'Content-Length': fileBuffer.length
      });
      res.send(fileBuffer);

    } catch (e) {
      console.error('/generate/from-file error:', e.message);
      res.status(500).json({ error: e.message || '伺服器錯誤，請重試' });
    }
  }
);

// ── 啟動 ──────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`⚙️  HEPHAESTUS API 啟動於 port ${PORT}`));
