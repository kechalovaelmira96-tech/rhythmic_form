// Full-stack single-file app: Node.js + Express serving a pixel-perfect web form
// Updates per user request:
// - Buttons moved to the bottom; upper buttons removed; "Отправить" isolated on the right
// - Removed link to download the Excel of all submissions from UI (endpoint kept optional)
// - On submit: append to Excel, generate .docx, and email it as attachment to WORK_EMAIL
// - .docx: no "ПРИЛОЖЕНИЕ №1", no date line, no caption "Индивидуальные упражнения"
// - Mobile-friendly styles
//
// How to run:
// 1) npm init -y
// 2) npm i express body-parser exceljs docx dayjs nodemailer dotenv
// 3) Create a .env file next to this file with SMTP settings (see below)
// 4) node app.js
// 5) Open http://localhost:3000
//
// .env example:
// SMTP_HOST=smtp.yourmail.com
// SMTP_PORT=587
// SMTP_USER=your_login
// SMTP_PASS=your_app_password
// FROM_EMAIL="Заявки турнир <noreply@yourmail.com>"
// WORK_EMAIL=work@yourcompany.com

const express = require('express');
const path = require('path');
const fs = require('fs');
const bodyParser = require('body-parser');
const ExcelJS = require('exceljs');
const {
  Document, Packer, Paragraph, TextRun, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  TableLayoutType, PageOrientation, ShadingType
} = require('docx');
// A4 размеры в твипах
const A4_WIDTH  = 11906;
const A4_HEIGHT = 16838;
// Безопасный фикс для layout
const FIXED_LAYOUT = (TableLayoutType && TableLayoutType.FIXED) || undefined;
const dayjs = require('dayjs');
const nodemailer = require('nodemailer');
require('dotenv').config();

const app = express();
app.use(bodyParser.json({ limit: '2mb' }));

// Ensure data directory exists
const DATA_DIR = path.join(__dirname, 'data');
if (!fs.existsSync(DATA_DIR)) fs.mkdirSync(DATA_DIR);
const EXCEL_PATH = path.join(DATA_DIR, 'submissions.xlsx');

// Mail transport
const transporter = nodemailer.createTransport({
  host: process.env.SMTP_HOST,
  port: Number(process.env.SMTP_PORT || 587),
  secure: true,
  auth: { user: process.env.SMTP_USER, pass: process.env.SMTP_PASS },
});
const FROM_EMAIL = process.env.FROM_EMAIL || 'noreply@example.com';
const WORK_EMAIL = process.env.WORK_EMAIL || 'you@example.com';

// Serve page
app.get('/', (_req, res) => {
  res.setHeader('Content-Type', 'text/html; charset=utf-8');
  res.end(getHtml());
});

// Submit: write to Excel, build docx, email it
app.post('/submit', async (req, res) => {
  try {
    const payload = sanitizeSubmission(req.body);
    await appendToExcel(payload);
    const buffer = await buildDocx(payload);
    await emailDocx(payload, buffer);
    res.json({ ok: true });
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: 'Не удалось сохранить и/или отправить письмо' });
  }
});

// Download .docx
app.post('/download-docx', async (req, res) => {
  try {
    const payload = sanitizeSubmission(req.body);
    const buffer = await buildDocx(payload);
    const safeName = fileSafe(payload.club || 'Заявка');
    const fileName = `Заявка_${safeName}.docx`;
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(fileName)}`);
    res.end(Buffer.from(buffer));
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: 'Failed to generate DOCX' });
  }
});

// (Optional) Download consolidated Excel
app.get('/download-excel', async (_req, res) => {
  try {
    if (!fs.existsSync(EXCEL_PATH)) await ensureWorkbook();
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent('submissions.xlsx')}`);
    fs.createReadStream(EXCEL_PATH).pipe(res);
  } catch (err) {
    console.error(err);
    res.status(500).json({ ok: false, error: 'Failed to download Excel' });
  }
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Running on http://localhost:${PORT}`));

// ------------------------ Helpers ------------------------
function sanitizeSubmission(b) {
  const pick = (v) => (v == null ? '' : String(v).trim());
  return {
    date: pick(b.date) || dayjs().format('DD.MM.YYYY'),
    city: pick(b.city),
    club: pick(b.club),
    contacts: pick(b.contacts),
    coach: pick(b.coach),
    judge: pick(b.judge),
    judgeCategory: pick(b.judgeCategory),
    participants: Array.isArray(b.participants)
      ? b.participants.map((p, i) => ({
          idx: i + 1,
          name: pick(p.name),
          birthYear: pick(p.birthYear),
          hasRank: pick(p.hasRank),
          performingRank: pick(p.performingRank),
          medicalVisa: pick(p.medicalVisa),
        }))
      : [],
  };
}

function fileSafe(name) {
  return name.replace(/[\\/:*?"<>|\n\r]+/g, '_').replace(/\s+/g, ' ').trim();
}

async function ensureWorkbook() {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('Submissions');
  ws.columns = [
    { header: 'Timestamp', key: 'ts', width: 20 },
    { header: 'Date (on form)', key: 'date', width: 16 },
    { header: 'City', key: 'city', width: 16 },
    { header: 'Club/School', key: 'club', width: 32 },
    { header: 'Contacts', key: 'contacts', width: 32 },
    { header: 'Coach (FIO)', key: 'coach', width: 28 },
    { header: 'Judge (FIO)', key: 'judge', width: 28 },
    { header: 'Judge Category', key: 'judgeCategory', width: 18 },
    { header: 'Participant #', key: 'p_idx', width: 12 },
    { header: 'Participant Name', key: 'p_name', width: 28 },
    { header: 'Birth Year', key: 'p_birth', width: 12 },
    { header: 'Has Rank', key: 'p_has', width: 14 },
    { header: 'Performing Rank', key: 'p_perf', width: 16 },
    { header: 'Medical Visa', key: 'p_med', width: 16 },
  ];
  await wb.xlsx.writeFile(EXCEL_PATH);
}

async function appendToExcel(payload) {
  const wb = new ExcelJS.Workbook();
  if (fs.existsSync(EXCEL_PATH)) {
    await wb.xlsx.readFile(EXCEL_PATH);
  } else {
    await ensureWorkbook();
    await wb.xlsx.readFile(EXCEL_PATH);
  }
  const ws = wb.getWorksheet('Submissions') || wb.addWorksheet('Submissions');
  const ts = dayjs().format('YYYY-MM-DD HH:mm:ss');
  if (!payload.participants.length) {
    ws.addRow({ ts, date: payload.date, city: payload.city, club: payload.club, contacts: payload.contacts, coach: payload.coach, judge: payload.judge, judgeCategory: payload.judgeCategory });
  } else {
    payload.participants.forEach((p) => {
      ws.addRow({ ts, date: payload.date, city: payload.city, club: payload.club, contacts: payload.contacts, coach: payload.coach, judge: payload.judge, judgeCategory: payload.judgeCategory, p_idx: p.idx, p_name: p.name, p_birth: p.birthYear, p_has: p.hasRank, p_perf: p.performingRank, p_med: p.medicalVisa });
    });
  }
  await wb.xlsx.writeFile(EXCEL_PATH);
}

async function emailDocx(payload, buffer) {
  const safeName = fileSafe(payload.club || 'Заявка');
  const fileName = `Заявка_${safeName}.docx`;
  await transporter.sendMail({
    from: FROM_EMAIL,
    to: WORK_EMAIL,
    subject: `Заявка: ${payload.club || 'без названия'}`,
    text: [
      `Клуб/школа: ${payload.club || '-'}`,
      `Город: ${payload.city || '-'}`,
      `Тренер: ${payload.coach || '-'}`,
      `Контакты: ${payload.contacts || '-'}`,
      `Судья: ${[payload.judge, payload.judgeCategory].filter(Boolean).join(', ') || '-'}`,
      `Участниц: ${payload.participants?.length || 0}`,
    ].join('\n'),
    attachments: [
      { filename: fileName, content: buffer, contentType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' },
    ],
  });
}

function cell(text, opts = {}) {
  return new TableCell({
    width: { size: opts.width || 1000, type: WidthType.DXA },
    margins: { top: 120, bottom: 120, left: 120, right: 120 },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 2 },
      bottom: { style: BorderStyle.SINGLE, size: 2 },
      left: { style: BorderStyle.SINGLE, size: 2 },
      right: { style: BorderStyle.SINGLE, size: 2 },
    },
    // <<< это и рисует серый фон в итоговом .docx
    shading: opts.shading
      ? { type: ShadingType.CLEAR, color: 'auto', fill: opts.shading }
      : undefined,
    children: [
      new Paragraph({
        alignment: opts.align || AlignmentType.LEFT,
        children: [ new TextRun({ text, font: 'Times New Roman', bold: !!opts.bold }) ],
      }),
    ],
  });
}

async function buildDocx(data) {
  // базовый шрифт/кегль
  const baseRun = (text) => new TextRun({ text, font: 'Times New Roman', size: 24 }); // 12pt

  const doc = new Document({
    sections: [
      {
        properties: {
          // Поля страницы — чтобы колонкам хватало места
          page: {
            ssize: { width: A4_WIDTH, height: A4_HEIGHT, orientation: PageOrientation.PORTRAIT },
            margin: { top: 1134, right: 1134, bottom: 1134, left: 1134 }, // ~2 см
          },
        },
        children: [
          // Заголовки
          new Paragraph({ alignment: AlignmentType.CENTER, children: [ new TextRun({ text: 'ЗАЯВКА', bold: true, font: 'Times New Roman', size: 28 }) ] }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [ baseRun('на участие в открытом турнире по художественной гимнастике') ] }),
          new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [ new TextRun({ text: '«Акварель', italics: true, bold: true, font: 'Times New Roman', size: 24 }),
                        new TextRun({ text: 'Dance»', italics: true, bold: true, font: 'Times New Roman', size: 24 }) ],
          }),
          new Paragraph({ alignment: AlignmentType.CENTER, children: [ baseRun(`г. Мытищи, ${data.date}`) ] }),
          new Paragraph({ text: ' ', spacing: { after: 200 } }),

          // ===== Таблица сведений (2 колонки) =====
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            layout: TableLayoutType.FIXED,           // ключ к корректному рендеру на мобилке
            columnWidths: [5500, 5500],               // фиксируем ширины столбцов (две колонки ~50/50)
            rows: [
              new TableRow({
                children: [
                  cell('Название клуба/спортивной школы', { width: 5500 }),
                  cell(data.club || '', { width: 5500 }),
                ],
              }),
              new TableRow({ children: [ cell('Город', { width: 5500 }), cell(data.city || '', { width: 5500 }) ] }),
              new TableRow({ children: [ cell('Контакты (телефон, электронная почта)', { width: 5500 }), cell(data.contacts || '', { width: 5500 }) ] }),
              new TableRow({ children: [ cell('Тренер (Ф.И.О)', { width: 5500 }), cell(data.coach || '', { width: 5500 }) ] }),
              new TableRow({ children: [ cell('Судья (Ф.И.О), судейская категория', { width: 5500 }), cell([data.judge, data.judgeCategory].filter(Boolean).join(', '), { width: 5500 }) ] }),
            ],
          }),

          new Paragraph({ text: ' ', spacing: { after: 160 } }),
          new Paragraph({
            spacing: { before: 200, after: 120 },       // небольшой отступ
            children: [
              new TextRun({
                text: 'Индивидуальные упражнения',
                font: 'Times New Roman',
                italics: true,                            // как на скрине
                bold: false,
                size: 28                                  // ~14pt
              }),
            ],
          }),

          // ===== Таблица участниц (6 колонок) =====
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            layout: FIXED_LAYOUT,           // фиксируем раскладку
            // ширины в twips; сумма примерно соответствует ширине контента
            columnWidths: [900, 3500, 1400, 1700, 1900, 1500],
            rows: [
              new TableRow({
                children: [
                  cell('№\nп/п',            { shading: 'D9D9D9', width: 900,  bold: true, align: AlignmentType.CENTER }),
                  cell('ФИО гимнастки',    { shading: 'D9D9D9', width: 3500, bold: true, align: AlignmentType.CENTER }),
                  cell('Год рождения',     { shading: 'D9D9D9', width: 1400, bold: true, align: AlignmentType.CENTER }),
                  cell('Имеет разряд',     { shading: 'D9D9D9', width: 1700, bold: true, align: AlignmentType.CENTER }),
                  cell('Выступает разряд', { shading: 'D9D9D9', width: 1900, bold: true, align: AlignmentType.CENTER }),
                  cell('Виза врача',       { shading: 'D9D9D9', width: 1500, bold: true, align: AlignmentType.CENTER }),
                ],
              }),              
              ...((data.participants && data.participants.length ? data.participants : new Array(8).fill(null)).map((p, i) =>
                new TableRow({
                  children: [
                    cell(String(p ? p.idx : i + 1), { width: 900 }),
                    cell(p ? (p.name || '') : '', { width: 3500 }),
                    cell(p ? (p.birthYear || '') : '', { width: 1400 }),
                    cell(p ? (p.hasRank || '') : '', { width: 1700 }),
                    cell(p ? (p.performingRank || '') : '', { width: 1900 }),
                    cell(p ? (p.medicalVisa || '') : '', { width: 1500 }),
                  ],
                })
              )),
            ],
          }),
        ],
      },
    ],
  });

  return await Packer.toBuffer(doc);
}

function getHtml() {
  return `<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>Заявка – Акварель Dance</title>
  <style>
    :root { --border: #000; }
    * { box-sizing: border-box; }
    body { font-family: 'Times New Roman', Georgia, serif; color:#000; background:#f7f7f7; margin:0; }
    .page { max-width: 860px; margin: 24px auto; background:#fff; padding: 28px 32px; box-shadow: 0 1px 6px rgba(0,0,0,.08); }
    .right { text-align: right; }
    .center { text-align: center; }
    h1 { font-size: 20px; margin: 6px 0; }
    .italic { font-style: italic; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; border: 2px solid var(--border); border-bottom: none; }
    .row { display: contents; }
    .cell { border-bottom:2px solid var(--border); border-right:2px solid var(--border); padding:10px; min-height:46px; }
    .cell:nth-child(2n) { border-right: none; }
    table { width: 100%; border-collapse: collapse; }
    th, td { border: 2px solid #000; padding: 8px; vertical-align: top; }
    th { font-weight: bold; }
    .controls { display:flex; gap: 8px; margin: 16px 0; }
    .btn { border:1px solid #111; background:#fff; padding:8px 12px; cursor:pointer; border-radius: 10px; }
    .btn:active{ transform: translateY(1px); }
    input, textarea, select { width:100%; border:1px solid #999; padding:6px 8px; font-family: inherit; font-size: 16px; }
    .muted { color:#444; }
    .table-wrap { overflow-x: auto; -webkit-overflow-scrolling: touch; }
    .table-wrap table { min-width: 720px; }

    .controls-bottom { display: flex; justify-content: space-between; align-items: center; margin-top: 14px; }
    .left-group { display: flex; gap: 8px; flex-wrap: wrap; }
    .right-group .btn.primary { font-weight: 600; border-width: 2px; }

    @media (max-width: 640px) {
      body { background: #fff; }
      .page { padding: 16px; max-width: 100%; box-shadow: none; }
      h1 { font-size: 18px; margin: 8px 0 6px; }
      .right { text-align: left; }
      .grid { display: block; border: none; }
      .grid .row { display: block; }
      .grid .cell { border: none; padding: 8px 0; }
      .grid .row:not(:first-child) { border-top: 1px solid #000; padding-top: 8px; }
      .grid .row .cell:first-child { font-weight: bold; margin-bottom: 6px; }
      .grid .row .cell:last-child { border: 1px solid #000; padding: 8px; }
      .controls-bottom { flex-direction: column; gap: 10px; align-items: stretch; }
      .left-group { order: 2; }
      .right-group { order: 1; }
      .right-group .btn { width: 100%; }
    }
  </style>
</head>
<body>
  <div class="page">
    <h1 class="center">ЗАЯВКА</h1>
    <div class="center italic">на участие в открытом турнире по художественной гимнастике</div>
    <div class="center italic"><strong>«Акварель Dance»</strong></div>
    <div class="center" style="margin-bottom:10px;">г. Мытищи, <span id="dateDisplay">12 октября 2025 г.</span></div>

    <div class="grid" id="infoTable">
      <div class="row">
        <div class="cell">Название клуба/спортивной школы</div>
        <div class="cell"><input id="club" placeholder="Введите название" /></div>
      </div>
      <div class="row">
        <div class="cell">Город</div>
        <div class="cell"><input id="city" placeholder="Например: Мытищи"/></div>
      </div>
      <div class="row">
        <div class="cell">Контакты (телефон, электронная почта)</div>
        <div class="cell"><textarea id="contacts" rows="2" placeholder="+7..., email@..."></textarea></div>
      </div>
      <div class="row">
        <div class="cell">Тренер (Ф.И.О)</div>
        <div class="cell"><input id="coach"/></div>
      </div>
      <div class="row">
        <div class="cell">Судья (Ф.И.О), судейская категория</div>
        <div class="cell"><input id="judge" placeholder="ФИО"/><br/><input id="judgeCategory" placeholder="Категория"/></div>
      </div>
    </div>

    <div class="table-wrap">
      <table id="participants">
        <thead>
          <tr>
            <th style="width:60px">№\nп/п</th>
            <th>ФИО гимнастки</th>
            <th style="width:120px">Год рождения</th>
            <th style="width:140px">Имеет разряд</th>
            <th style="width:160px">Выступает разряд</th>
            <th style="width:120px">Виза врача</th>
          </tr>
        </thead>
        <tbody></tbody>
      </table>
    </div>

    <div class="controls controls-bottom">
      <div class="left-group">
        <button class="btn" id="addRow" type="button">+ Добавить участницу</button>
        <button class="btn" id="removeRow" type="button">− Удалить последнюю</button>
        <button class="btn" id="docxBtn" type="button">Скачать .docx</button>
      </div>
      <div class="right-group">
        <button class="btn primary" id="submitBtn" type="button">Отправить</button>
      </div>
    </div>

    <div id="status" class="muted"></div>
  </div>

  <script>
    window.addEventListener('DOMContentLoaded', () => {
      const tbody = document.querySelector('#participants tbody');
      const dateDisplay = document.getElementById('dateDisplay');

      function renumber(){
        [...tbody.querySelectorAll('tr')].forEach((tr, i)=>{
          const idx = tr.querySelector('.idx');
          if (idx) idx.textContent = i + 1;
        });
      }

      function addRow(){
        const tr = document.createElement('tr');
        tr.innerHTML = '<td class="idx"></td>' +
          '<td><input placeholder="ФИО"/></td>' +
          '<td><input placeholder="дд.мм.гггг"/></td>' +
          '<td><input placeholder="Да/Нет, разряд"/></td>' +
          '<td><input placeholder="Разряд"/></td>' +
          '<td><input placeholder="Есть/Нет"/></td>';
        tbody.appendChild(tr); renumber();
      }

      function removeRow(){ if (tbody.children.length>1) { tbody.removeChild(tbody.lastElementChild); renumber(); } }

      function collect(){
        const get = id => (document.getElementById(id)?.value || '').trim();
        const participants = [...tbody.querySelectorAll('tr')].map(tr => {
          const [name, birthYear, hasRank, performingRank, medicalVisa] = [...tr.querySelectorAll('input')].map(i=>i.value.trim());
          return { name, birthYear, hasRank, performingRank, medicalVisa };
        });
        return {
          date: dateDisplay.textContent,
          city: get('city'), club: get('club'), contacts: get('contacts'), coach: get('coach'),
          judge: get('judge'), judgeCategory: get('judgeCategory'), participants
        };
      }

      function fileNameFromClub(){
        const club = (document.getElementById('club')?.value || 'Заявка').trim();
        return 'Заявка_' + club.replace(/[^\w\u0400-\u04FF\s.-]/g,'_').replace(/\s+/g,' ').trim() + '.docx';
      }

      async function submitForm(){
        const payload = collect(); const s = document.getElementById('status');
        try {
          const res = await fetch('/submit', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)});
          const data = await res.json(); s.textContent = data.ok ? 'Заявка отправлена на почту и сохранена.' : 'Ошибка: ' + (data.error||'');
        } catch(e){ s.textContent = 'Сеть/сервер недоступны'; console.error(e); }
      }

      async function downloadDocx(){
        const payload = collect();
        try { const res = await fetch('/download-docx', { method:'POST', headers:{'Content-Type':'application/json'}, body: JSON.stringify(payload)});
          const blob = await res.blob(); const url = URL.createObjectURL(blob); const a = document.createElement('a');
          a.href = url; a.download = fileNameFromClub(); a.click(); URL.revokeObjectURL(url);
        } catch(e){ console.error(e); alert('Не удалось сформировать .docx'); }
      }

      document.getElementById('addRow')?.addEventListener('click', addRow);
      document.getElementById('removeRow')?.addEventListener('click', removeRow);
      document.getElementById('submitBtn')?.addEventListener('click', submitForm);
      document.getElementById('docxBtn')?.addEventListener('click', downloadDocx);

      for (let i=0;i<1;i++) addRow();
    });
  </script>
</body>
</html>`;
}