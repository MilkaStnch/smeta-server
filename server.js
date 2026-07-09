const express = require('express');
const cors = require('cors');
const TelegramBot = require('node-telegram-bot-api');
const ExcelJS = require('exceljs');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const BOT_TOKEN = process.env.BOT_TOKEN;
const bot = new TelegramBot(BOT_TOKEN);

app.get('/', (req, res) => res.send('Smeta server works!'));

// ── ОТПРАВИТЬ HTML ──
app.post('/send-pdf', async (req, res) => {
  const { chatId, smetaData } = req.body;
  if (!chatId || !smetaData) return res.status(400).json({ error: 'Missing data' });
  try {
    const html = buildHTML(smetaData);
    const buffer = Buffer.from(html, 'utf-8');
    const fname = (smetaData.objName || 'smeta').replace(/\s+/g, '_') + '.html';
    await bot.sendDocument(chatId, buffer, {
      caption: '📋 Смета: ' + (smetaData.objName || '') + '\n💰 ' + getTotals(smetaData) + '\n\n📌 Откройте файл в браузере → нажмите "Сохранить PDF"'
    }, { filename: fname, contentType: 'text/html' });
    res.json({ ok: true });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ── ОТПРАВИТЬ .json ──
app.post('/send-smeta', async (req, res) => {
  const { chatId, smetaData } = req.body;
  if (!chatId || !smetaData) return res.status(400).json({ error: 'Missing data' });
  try {
    const content = JSON.stringify(smetaData, null, 2);
    const buffer = Buffer.from(content, 'utf-8');
    const fname = (smetaData.name || smetaData.objName || 'smeta').replace(/\s+/g, '_') + '.json';
    await bot.sendDocument(chatId, buffer, {
      caption: '📋 ' + (smetaData.name || smetaData.objName || 'Смета') + '\n💰 ' + getTotals(smetaData) + '\n\n📌 Загрузите файл в приложение чтобы открыть и редактировать'
    }, { filename: fname, contentType: 'application/json' });
    res.json({ ok: true });
  } catch(e) {
    res.status(500).json({ error: e.message });
  }
});

// ── ОТПРАВИТЬ EXCEL ──
app.post('/send-excel', async (req, res) => {
  const { chatId, smetaData } = req.body;
  if (!chatId || !smetaData) return res.status(400).json({ error: 'Missing data' });
  try {
    const buffer = await buildExcel(smetaData);
    const fname = (smetaData.objName || 'smeta').replace(/\s+/g, '_') + '.xlsx';
    await bot.sendDocument(chatId, buffer, {
      caption: '📊 Смета (Excel): ' + (smetaData.objName || '') + '\n💰 ' + getTotals(smetaData)
    }, { filename: fname, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    res.json({ ok: true });
  } catch(e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

// ── HELPERS ──
function getTotals(d) {
  let mT = 0, wT = 0;
  (d.matData || []).forEach(r => mT += r.sum);
  (d.workData || []).forEach(r => wT += r.sum);
  return 'Итого: ' + fmt(mT + wT) + ' руб.';
}

function fmt(n) { return Number(n || 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

// ── КРАСИВЫЙ EXCEL ──
async function buildExcel(d) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Смета');

  const GREEN  = '1D9E75';
  const DKGREEN = '0F6E56';
  const LTGREEN = 'E1F5EE';
  const STRIPE  = 'F2FAF7';
  const WHITE   = 'FFFFFF';

  ws.columns = [
    { width: 5 },
    { width: 38 },
    { width: 12 },
    { width: 13 },
    { width: 16 },
    { width: 18 },
  ];

  function addRow(vals, styles) {
    const row = ws.addRow(vals);
    styles && styles.forEach((s, i) => {
      if (!s) return;
      const cell = row.getCell(i + 1);
      if (s.bold || s.color) cell.font = { name: 'Arial', bold: !!s.bold, color: s.color ? { argb: 'FF' + s.color } : undefined, size: s.size || 10 };
      if (s.bg) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + s.bg } };
      if (s.align) cell.alignment = { horizontal: s.align, vertical: 'middle', wrapText: true };
      if (s.border) cell.border = { bottom: { style: 'thin', color: { argb: 'FFD0EDE3' } } };
    });
    row.height = styles && styles[0] && styles[0].height ? styles[0].height : 18;
    return row;
  }

  function mergeRow(rowNum, from, to, val, font, fill, align) {
    ws.mergeCells(rowNum, from, rowNum, to);
    const cell = ws.getCell(rowNum, from);
    cell.value = val;
    if (font) cell.font = { name: 'Arial', ...font };
    if (fill) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + fill } };
    if (align) cell.alignment = { horizontal: align, vertical: 'middle', wrapText: true };
  }

  let R = 1;

  // Заголовок
  ws.addRow(['СМЕТА НА СТРОИТЕЛЬНО-МОНТАЖНЫЕ РАБОТЫ', '', '', '', '', '']);
  ws.mergeCells(R, 1, R, 6);
  const titleCell = ws.getCell(R, 1);
  titleCell.font = { name: 'Arial', bold: true, size: 14, color: { argb: 'FF' + WHITE } };
  titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GREEN } };
  titleCell.alignment = { horizontal: 'center', vertical: 'middle' };
  ws.getRow(R).height = 30;
  R++;

  // Пустая строка
  ws.addRow([]); R++;

  // Инфо блок
  const infoStyle = [{ bold: true, color: DKGREEN, size: 10 }, { size: 10 }, null, null, null, null];
  const infos = [
    ['Объект:', d.objName || ''],
    ['Адрес:', d.address || ''],
    ['Заказчик:', d.client || ''],
    ['Составитель:', d.author || ''],
    ['Дата:', d.date || ''],
  ];
  if (d.note) infos.push(['Примечание:', d.note]);

  infos.forEach(([lbl, val]) => {
    const row = ws.addRow([lbl, val, '', '', '', '']);
    ws.mergeCells(R, 2, R, 6);
    row.getCell(1).font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FF' + DKGREEN } };
    row.getCell(2).font = { name: 'Arial', size: 10 };
    row.getCell(2).alignment = { wrapText: true };
    row.height = 18;
    R++;
  });

  ws.addRow([]); R++;

  // Секция helper
  function addSection(title, rows, type) {
    // Заголовок раздела
    ws.addRow([title, '', '', '', '', '']);
    ws.mergeCells(R, 1, R, 6);
    const secCell = ws.getCell(R, 1);
    secCell.value = title;
    secCell.font = { name: 'Arial', bold: true, size: 11, color: { argb: 'FF' + WHITE } };
    secCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + DKGREEN } };
    secCell.alignment = { horizontal: 'left', vertical: 'middle' };
    ws.getRow(R).height = 22;
    R++;

    // Заголовки колонок
    const hrow = ws.addRow(['№', 'Наименование', 'Ед. изм.', 'Количество', 'Цена (руб.)', 'Сумма (руб.)']);
    hrow.eachCell(cell => {
      cell.font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FF' + WHITE } };
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + GREEN } };
      cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    ws.getRow(R).height = 20;
    R++;

    // Данные
    let total = 0, num = 1, shade = false;
    rows.forEach(item => {
      const drow = ws.addRow([num++, item.name, item.unit, item.qty, item.price, item.sum]);
      const bg = shade ? STRIPE : WHITE;
      drow.eachCell((cell, ci) => {
        cell.font = { name: 'Arial', size: 10 };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + bg } };
        cell.border = { bottom: { style: 'hair', color: { argb: 'FFD0EDE3' } } };
        if (ci === 1) cell.alignment = { horizontal: 'center' };
        if (ci >= 4) { cell.alignment = { horizontal: 'right' }; cell.numFmt = '#,##0.00'; }
      });
      ws.getRow(R).height = 18;
      total += item.sum;
      shade = !shade;
      R++;
    });

    if (!rows.length) {
      ws.addRow(['—', '(нет позиций)', '', '', '', '']);
      R++;
    }

    // Итого раздела
    const srow = ws.addRow(['', '', '', '', 'Итого:', total]);
    ws.mergeCells(R, 1, R, 4);
    srow.getCell(5).font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FF' + DKGREEN } };
    srow.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + LTGREEN } };
    srow.getCell(5).alignment = { horizontal: 'right' };
    srow.getCell(6).font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FF' + DKGREEN } };
    srow.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + LTGREEN } };
    srow.getCell(6).alignment = { horizontal: 'right' };
    srow.getCell(6).numFmt = '#,##0.00';
    ws.getRow(R).height = 20;
    R++;

    ws.addRow([]); R++;
    return total;
  }

  const mTotal = addSection('РАЗДЕЛ 1. МАТЕРИАЛЫ', d.matData || [], 'mat');
  const wTotal = addSection('РАЗДЕЛ 2. РАБОТЫ', d.workData || [], 'work');

  // ИТОГО ВСЕГО
  const grand = ws.addRow(['', '', '', '', 'ИТОГО ВСЕГО:', mTotal + wTotal]);
  ws.mergeCells(R, 1, R, 4);
  grand.getCell(5).font = { name: 'Arial', bold: true, size: 12, color: { argb: 'FF' + WHITE } };
  grand.getCell(5).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + DKGREEN } };
  grand.getCell(5).alignment = { horizontal: 'right', vertical: 'middle' };
  grand.getCell(6).font = { name: 'Arial', bold: true, size: 12, color: { argb: 'FF' + WHITE } };
  grand.getCell(6).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF' + DKGREEN } };
  grand.getCell(6).alignment = { horizontal: 'right', vertical: 'middle' };
  grand.getCell(6).numFmt = '#,##0.00';
  ws.getRow(R).height = 26;
  R++;

  ws.addRow([]); R++;
  ws.addRow([]); R++;

  // Подписи
  const s1 = ws.addRow(['Заказчик:', '', '_______________________', '', d.client || '', '']);
  s1.getCell(1).font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FF' + DKGREEN } };
  ws.getRow(R).height = 20; R++;

  const s2 = ws.addRow(['Исполнитель:', '', '_______________________', '', d.author || '', '']);
  s2.getCell(1).font = { name: 'Arial', bold: true, size: 10, color: { argb: 'FF' + DKGREEN } };
  ws.getRow(R).height = 20;

  return await wb.xlsx.writeBuffer();
}

function buildHTML(d) {
  let mT = 0, wT = 0, matHTML = '', workHTML = '', num = 1;
  (d.matData || []).forEach(r => {
    const shade = num % 2 === 0 ? 'background:#F2FAF7;' : '';
    matHTML += `<tr style="${shade}"><td style="text-align:center">${num++}</td><td>${r.name}</td><td style="text-align:center">${r.unit}</td><td style="text-align:right">${fmt(r.qty)}</td><td style="text-align:right">${fmt(r.price)}</td><td style="text-align:right;font-weight:600">${fmt(r.sum)} руб.</td></tr>`;
    mT += r.sum;
  });
  if (!matHTML) matHTML = `<tr><td colspan="6" style="text-align:center;color:#aaa">— нет позиций —</td></tr>`;
  num = 1;
  (d.workData || []).forEach(r => {
    const shade = num % 2 === 0 ? 'background:#F2FAF7;' : '';
    workHTML += `<tr style="${shade}"><td style="text-align:center">${num++}</td><td>${r.name}</td><td style="text-align:center">${r.unit}</td><td style="text-align:right">${fmt(r.qty)}</td><td style="text-align:right">${fmt(r.price)}</td><td style="text-align:right;font-weight:600">${fmt(r.sum)} руб.</td></tr>`;
    wT += r.sum;
  });
  if (!workHTML) workHTML = `<tr><td colspan="6" style="text-align:center;color:#aaa">— нет позиций —</td></tr>`;

  return `<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8"><title>Смета</title>
<style>@page{size:A4;margin:15mm 14mm}*{box-sizing:border-box;margin:0;padding:0}
body{font-family:Arial,sans-serif;font-size:10pt;color:#1A1A18}
.print-btn{display:block;margin:0 auto 14px;padding:10px 28px;background:#1D9E75;color:#fff;border:none;border-radius:8px;font-size:13pt;font-weight:bold;cursor:pointer}
@media print{.print-btn{display:none}}
.hdr{background:#1D9E75;color:#fff;padding:10px 16px;border-radius:6px;margin-bottom:10px}
.hdr h1{font-size:13pt;font-weight:bold;margin-bottom:2px}.hdr p{font-size:9pt;opacity:.85}
.info{background:#F7F6F3;border:1px solid #E2E0D8;border-radius:6px;padding:10px 14px;margin-bottom:10px;display:grid;grid-template-columns:1fr 1fr;gap:4px 20px}
.ir{display:flex;gap:6px;font-size:9.5pt}.il{color:#0F6E56;font-weight:bold;white-space:nowrap;min-width:80px}
.sh{background:#0F6E56;color:#fff;font-weight:bold;font-size:10pt;padding:6px 10px;border-radius:4px 4px 0 0;margin-top:10px}
table{width:100%;border-collapse:collapse;font-size:9pt}
thead th{background:#1D9E75;color:#fff;padding:6px 8px;text-align:left}
thead th:nth-child(1){width:5%;text-align:center}thead th:nth-child(2){width:35%}
thead th:nth-child(3){width:12%;text-align:center}thead th:nth-child(4){width:13%;text-align:right}
thead th:nth-child(5){width:17%;text-align:right}thead th:nth-child(6){width:18%;text-align:right}
tbody td{padding:5px 8px;border-bottom:1px solid #E8F5EE}
tbody tr:nth-child(even) td{background:#F2FAF7}
.sub td{background:#E1F5EE!important;padding:6px 8px;font-weight:bold;color:#0F6E56}
.grand td{background:#0F6E56!important;color:#fff;font-weight:bold;font-size:11pt;padding:8px 10px}
.signs{margin-top:16px;display:grid;grid-template-columns:1fr 1fr;gap:10px;font-size:9.5pt}
.sline{border-bottom:1px solid #999;margin:14px 0 4px}.slbl{color:#6B6A65;font-size:9pt}
.ft{margin-top:14px;border-top:1px solid #E2E0D8;padding-top:6px;font-size:8pt;color:#aaa;display:flex;justify-content:space-between}
</style></head><body>
<button class="print-btn" onclick="window.print()">🖨️ Сохранить как PDF</button>
<div class="hdr"><h1>СМЕТА НА СТРОИТЕЛЬНО-МОНТАЖНЫЕ РАБОТЫ</h1><p>Строительство и ремонт</p></div>
<div class="info">
<div class="ir"><span class="il">Объект:</span><span>${d.objName}</span></div>
<div class="ir"><span class="il">Составитель:</span><span>${d.author}</span></div>
<div class="ir"><span class="il">Адрес:</span><span>${d.address}</span></div>
<div class="ir"><span class="il">Дата:</span><span>${d.date}</span></div>
<div class="ir"><span class="il">Заказчик:</span><span>${d.client}</span></div>
${d.note ? `<div class="ir"><span class="il">Примечание:</span><span>${d.note}</span></div>` : ''}
</div>
<div class="sh">РАЗДЕЛ 1. МАТЕРИАЛЫ</div>
<table><thead><tr><th>№</th><th>Наименование</th><th>Ед. изм.</th><th>Кол-во</th><th>Цена (руб.)</th><th>Сумма (руб.)</th></tr></thead>
<tbody>${matHTML}<tr class="sub"><td colspan="4"></td><td>Итого:</td><td style="text-align:right">${fmt(mT)} руб.</td></tr></tbody></table>
<div class="sh" style="margin-top:12px">РАЗДЕЛ 2. РАБОТЫ</div>
<table><thead><tr><th>№</th><th>Наименование</th><th>Ед. изм.</th><th>Кол-во</th><th>Цена (руб.)</th><th>Сумма (руб.)</th></tr></thead>
<tbody>${workHTML}<tr class="sub"><td colspan="4"></td><td>Итого:</td><td style="text-align:right">${fmt(wT)} руб.</td></tr></tbody></table>
<table style="margin-top:8px"><tbody><tr class="grand"><td colspan="4">ИТОГО ВСЕГО:</td><td colspan="2" style="text-align:right">${fmt(mT + wT)} руб.</td></tr></tbody></table>
<div class="signs">
<div><div class="slbl">Заказчик:</div><div class="sline"></div><div style="font-weight:bold">${d.client}</div></div>
<div><div class="slbl">Исполнитель:</div><div class="sline"></div><div style="font-weight:bold">${d.author}</div></div>
</div>
<div class="ft"><span>Сформировано автоматически</span><span>${d.date}</span></div>
</body></html>`;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log('Server running on port', PORT));
