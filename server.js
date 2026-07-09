const express = require('express');
const cors = require('cors');
const TelegramBot = require('node-telegram-bot-api');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const BOT_TOKEN = process.env.BOT_TOKEN;
const bot = new TelegramBot(BOT_TOKEN);

app.get('/', (req, res) => res.send('Smeta server works!'));

// ── ОТПРАВИТЬ HTML (смета для PDF) ──
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

// ── ОТПРАВИТЬ .json (смета для импорта) ──
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
    const xlsx = buildXLSX(smetaData);
    const fname = (smetaData.objName || 'smeta').replace(/\s+/g, '_') + '.xlsx';
    await bot.sendDocument(chatId, xlsx, {
      caption: '📊 Смета (Excel): ' + (smetaData.objName || '') + '\n💰 ' + getTotals(smetaData)
    }, { filename: fname, contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    res.json({ ok: true });
  } catch(e) {
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

// ── EXCEL builder (без внешних зависимостей, чистый XML) ──
function buildXLSX(d) {
  let mT = 0, wT = 0;
  const matRows = (d.matData || []).map((r, i) => { mT += r.sum; return [i+1, r.name, r.unit, r.qty, r.price, r.sum]; });
  const workRows = (d.workData || []).map((r, i) => { wT += r.sum; return [i+1, r.name, r.unit, r.qty, r.price, r.sum]; });

  // Строим данные для таблицы
  const rows = [];
  rows.push(['СМЕТА НА СТРОИТЕЛЬНО-МОНТАЖНЫЕ РАБОТЫ']);
  rows.push([]);
  rows.push(['Объект:', d.objName || '']);
  rows.push(['Адрес:', d.address || '']);
  rows.push(['Заказчик:', d.client || '']);
  rows.push(['Составитель:', d.author || '']);
  rows.push(['Дата:', d.date || '']);
  if(d.note) rows.push(['Примечание:', d.note]);
  rows.push([]);
  rows.push(['РАЗДЕЛ 1. МАТЕРИАЛЫ']);
  rows.push(['№', 'Наименование', 'Ед. изм.', 'Количество', 'Цена (руб.)', 'Сумма (руб.)']);
  matRows.forEach(r => rows.push(r));
  if(!matRows.length) rows.push(['—', '(нет позиций)', '', '', '', '']);
  rows.push(['', '', '', '', 'Итого материалы:', mT]);
  rows.push([]);
  rows.push(['РАЗДЕЛ 2. РАБОТЫ']);
  rows.push(['№', 'Наименование', 'Ед. изм.', 'Количество', 'Цена (руб.)', 'Сумма (руб.)']);
  workRows.forEach(r => rows.push(r));
  if(!workRows.length) rows.push(['—', '(нет позиций)', '', '', '', '']);
  rows.push(['', '', '', '', 'Итого работы:', wT]);
  rows.push([]);
  rows.push(['', '', '', '', 'ИТОГО ВСЕГО:', mT + wT]);
  rows.push([]);
  rows.push(['Заказчик:', '', '', '_______________________', d.client || '']);
  rows.push(['Исполнитель:', '', '', '_______________________', d.author || '']);

  // Генерируем XML для xlsx
  const xmlRows = rows.map(row => {
    const cells = row.map((val, ci) => {
      const col = String.fromCharCode(65 + ci);
      if(typeof val === 'number') {
        return `<c r="${col}${rows.indexOf(row)+1}" t="n"><v>${val}</v></c>`;
      } else {
        const escaped = String(val||'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');
        return `<c r="${col}${rows.indexOf(row)+1}" t="inlineStr"><is><t>${escaped}</t></is></c>`;
      }
    });
    return `<row>${cells.join('')}</row>`;
  }).join('');

  const sheet = `<?xml version="1.0" encoding="UTF-8"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
<sheetData>${xmlRows}</sheetData>
</worksheet>`;

  const workbook = `<?xml version="1.0" encoding="UTF-8"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
<sheets><sheet name="Смета" sheetId="1" r:id="rId1"/></sheets>
</workbook>`;

  const rels = `<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
</Relationships>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
</Types>`;

  // Собираем ZIP вручную (простой метод)
  return createZip([
    { name: '[Content_Types].xml', data: contentTypes },
    { name: '_rels/.rels', data: `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: 'xl/workbook.xml', data: workbook },
    { name: 'xl/_rels/workbook.xml.rels', data: rels },
    { name: 'xl/worksheets/sheet1.xml', data: sheet },
  ]);
}

function createZip(files) {
  // Простой ZIP без компрессии
  const buffers = [];
  const centralDir = [];
  let offset = 0;

  files.forEach(file => {
    const nameBytes = Buffer.from(file.name, 'utf-8');
    const dataBytes = Buffer.from(file.data, 'utf-8');
    const crc = crc32(dataBytes);

    const local = Buffer.alloc(30 + nameBytes.length);
    local.writeUInt32LE(0x04034b50, 0);
    local.writeUInt16LE(20, 4);
    local.writeUInt16LE(0, 6);
    local.writeUInt16LE(0, 8);
    local.writeUInt16LE(0, 10);
    local.writeUInt16LE(0, 12);
    local.writeUInt32LE(crc, 14);
    local.writeUInt32LE(dataBytes.length, 18);
    local.writeUInt32LE(dataBytes.length, 22);
    local.writeUInt16LE(nameBytes.length, 26);
    local.writeUInt16LE(0, 28);
    nameBytes.copy(local, 30);

    const central = Buffer.alloc(46 + nameBytes.length);
    central.writeUInt32LE(0x02014b50, 0);
    central.writeUInt16LE(20, 4);
    central.writeUInt16LE(20, 6);
    central.writeUInt16LE(0, 8);
    central.writeUInt16LE(0, 10);
    central.writeUInt16LE(0, 12);
    central.writeUInt16LE(0, 14);
    central.writeUInt32LE(crc, 16);
    central.writeUInt32LE(dataBytes.length, 20);
    central.writeUInt32LE(dataBytes.length, 24);
    central.writeUInt16LE(nameBytes.length, 28);
    central.writeUInt16LE(0, 30);
    central.writeUInt16LE(0, 32);
    central.writeUInt16LE(0, 34);
    central.writeUInt16LE(0, 36);
    central.writeUInt32LE(0, 38);
    central.writeUInt32LE(offset, 42);
    nameBytes.copy(central, 46);

    buffers.push(local, dataBytes);
    centralDir.push(central);
    offset += local.length + dataBytes.length;
  });

  const centralBuf = Buffer.concat(centralDir);
  const end = Buffer.alloc(22);
  end.writeUInt32LE(0x06054b50, 0);
  end.writeUInt16LE(0, 4);
  end.writeUInt16LE(0, 6);
  end.writeUInt16LE(files.length, 8);
  end.writeUInt16LE(files.length, 10);
  end.writeUInt32LE(centralBuf.length, 12);
  end.writeUInt32LE(offset, 16);
  end.writeUInt16LE(0, 20);

  return Buffer.concat([...buffers, centralBuf, end]);
}

function crc32(buf) {
  let crc = 0xFFFFFFFF;
  for(let i = 0; i < buf.length; i++) {
    crc ^= buf[i];
    for(let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
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
