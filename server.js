const express = require('express');
const cors = require('cors');
const puppeteer = require('puppeteer');
const TelegramBot = require('node-telegram-bot-api');

const app = express();
app.use(cors());
app.use(express.json({ limit: '10mb' }));

const BOT_TOKEN = process.env.BOT_TOKEN;
const bot = new TelegramBot(BOT_TOKEN);

app.get('/', (req, res) => res.send('Smeta server works!'));

app.post('/send-pdf', async (req, res) => {
  const { chatId, smetaData } = req.body;
  if (!chatId || !smetaData) return res.status(400).json({ error: 'Missing data' });

  try {
    const html = buildHTML(smetaData);
    const browser = await puppeteer.launch({ args: ['--no-sandbox', '--disable-setuid-sandbox'] });
    const page = await browser.newPage();
    await page.setContent(html, { waitUntil: 'networkidle0' });
    const pdf = await page.pdf({ format: 'A4', margin: { top: '15mm', bottom: '15mm', left: '14mm', right: '14mm' } });
    await browser.close();

    await bot.sendDocument(chatId, Buffer.from(pdf), {}, {
      filename: 'Смета_' + (smetaData.objName || 'документ').replace(/\s+/g, '_') + '.pdf',
      contentType: 'application/pdf'
    });

    res.json({ ok: true });
  } catch (e) {
    console.error(e);
    res.status(500).json({ error: e.message });
  }
});

function fmt(n) { return Number(n || 0).toLocaleString('ru-RU', { minimumFractionDigits: 2, maximumFractionDigits: 2 }); }

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

  return `<!DOCTYPE html><html lang="ru"><head><meta charset="UTF-8">
<style>*{box-sizing:border-box;margin:0;padding:0}body{font-family:Arial,sans-serif;font-size:10pt;color:#1A1A18}
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
.sub td{background:#E1F5EE!important;padding:6px 8px;font-weight:bold;color:#0F6E56}
.grand td{background:#0F6E56!important;color:#fff;font-weight:bold;font-size:11pt;padding:8px 10px}
.signs{margin-top:16px;display:grid;grid-template-columns:1fr 1fr;gap:10px;font-size:9.5pt}
.sline{border-bottom:1px solid #999;margin:14px 0 4px}.slbl{color:#6B6A65;font-size:9pt}
.ft{margin-top:14px;border-top:1px solid #E2E0D8;padding-top:6px;font-size:8pt;color:#aaa;display:flex;justify-content:space-between}
</style></head><body>
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
