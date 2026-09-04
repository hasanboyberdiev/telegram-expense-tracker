const { Telegraf, Markup, session } = require('telegraf');
const sqlite3 = require('sqlite3').verbose();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const cron = require('node-cron');
const { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, AlignmentType, WidthType, HeadingLevel } = require('docx');
require('dotenv').config();

const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) {
  console.error("❌ Хатогӣ: BOT_TOKEN ёфт нашуд!");
  process.exit(1);
}

const bot = new Telegraf(TOKEN);

// 1. ХОТИРА ВА СЕРВЕР
bot.use(session());
bot.use((ctx, next) => {
    ctx.session = ctx.session || { step: 'none' };
    return next();
});

const app = express();
app.use(cors());
app.use(express.json());

// 2. БАЗАИ МАЪЛУМОТ
const db = new sqlite3.Database('./database.db');
db.run(`CREATE TABLE IF NOT EXISTS expenses (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER,
  title TEXT,
  amount REAL,
  category TEXT DEFAULT 'Дигар',
  date TEXT
)`);

// 3. МЕНЮИ АСОСӢ
const mainKeyboard = Markup.keyboard([
  ['📝 Илова кардан', '📋 Рӯйхат'],
  ['📄 Содирот ба Word'], 
  [Markup.button.webApp('🌐 Web App -ро кушодан', 'https://malumot.gt.tc/?v=10')]
]).resize();

bot.start((ctx) => {
  ctx.session.step = 'none';
  ctx.reply("✨ Хуш омадед ба Назоратчии Буҷет! Амалро интихоб кунед:", mainKeyboard);
});

// 4. РӮЙХАТ ВА ТУГМАҲО
bot.hears('📋 Рӯйхат', (ctx) => {
  ctx.session.step = 'none';
  db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY id DESC LIMIT 15", [ctx.from.id], (err, rows) => {
    if (err || rows.length === 0) return ctx.reply("📭 Рӯйхати шумо холӣ аст.");
    let txt = "📋 <b>Рӯйхати 15 хароҷоти охирин:</b>\n\n";
    rows.forEach(r => txt += `🆔 <b>ID: ${r.id}</b> | ${r.title} - ${r.amount} смн.\n`);
    ctx.reply(txt, { 
      parse_mode: 'HTML',
      ...Markup.inlineKeyboard([
        [Markup.button.callback('✏️ Таҳрир', 'action_edit'), Markup.button.callback('❌ Нест кардан', 'action_delete')]
      ])
    });
  });
});

// 5. СОДИРОТ БА WORD
bot.hears('📄 Содирот ба Word', (ctx) => {
  ctx.session.step = 'none';
  db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY date DESC", [ctx.from.id], async (err, rows) => {
    if (err || rows.length === 0) return ctx.reply("📭 Шумо ҳоло ягон хароҷот надоред.");

    try {
      let totalSum = 0;
      const tableRows = [];

      // Сарлавҳаи ҷадвал
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "ID", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
            new TableCell({ children: [new Paragraph({ text: "Номи хароҷот", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
            new TableCell({ children: [new Paragraph({ text: "Маблағ (смн)", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
            new TableCell({ children: [new Paragraph({ text: "Сана", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
          ],
        })
      );

      // Пур кардани маълумот
      rows.forEach((r) => {
        totalSum += r.amount;
        tableRows.push(
          new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ text: r.id.toString(), alignment: AlignmentType.CENTER })] }),
              new TableCell({ children: [new Paragraph(r.title)] }),
              new TableCell({ children: [new Paragraph({ text: r.amount.toString(), alignment: AlignmentType.CENTER })] }),
              new TableCell({ children: [new Paragraph({ text: r.date, alignment: AlignmentType.CENTER })] }),
            ],
          })
        );
      });

      // Сатри ҷамъбаст
      tableRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "ҶАМЪИ УМУМӢ:", alignment: AlignmentType.RIGHT })], columnSpan: 2, shading: { fill: "D9D9D9" } }),
            new TableCell({ children: [new Paragraph({ text: totalSum.toString(), alignment: AlignmentType.CENTER })], shading: { fill: "A9D08E" } }),
            new TableCell({ children: [new Paragraph("")] }),
          ],
        })
      );

      // Сохтани ҳуҷҷат
      const doc = new Document({
        sections: [{
          properties: {},
          children: [
            new Paragraph({
              text: "ҲИСОБОТИ МУФАССАЛИ ХАРОҶОТ",
              heading: HeadingLevel.HEADING_2,
              alignment: AlignmentType.CENTER,
            }),
            new Paragraph({ text: "" }), // Сатри холӣ
            new Table({
              rows: tableRows,
              width: { size: 100, type: WidthType.PERCENTAGE },
            }),
          ],
        }],
      });

      ctx.reply("⏳ Ҳуҷҷати Word омода шуда истодааст...");
      
      const buffer = await Packer.toBuffer(doc);
      const fileName = `Hisobot_${ctx.from.id}.docx`;
      fs.writeFileSync(fileName, buffer);

      await ctx.replyWithDocument(
        { source: fileName, filename: 'Ҳисоботи_Муфассал.docx' },
        { caption: "📄 <b>Ҳуҷҷати шумо омода аст!</b>", parse_mode: 'HTML' }
      );
      
      fs.unlinkSync(fileName); 

    } catch (e) {
      console.error(e);
      ctx.reply("❌ Хатогӣ ҳангоми сохтани файли Word.");
    }
  });
});

// 6. САРШАВИИ ҚАДАМҲО
bot.hears('📝 Илова кардан', (ctx) => {
  ctx.session.step = 'add_title';
  ctx.reply("📝 Лутфан, **номи амалиёт**-ро нависед:", { parse_mode: 'Markdown' });
});

bot.action('action_delete', (ctx) => {
  ctx.answerCbQuery();
  ctx.session.step = 'delete';
  ctx.reply("🗑 Фақат рақами ID-ро барои нест кардан нависед:");
});

bot.action('action_edit', (ctx) => {
  ctx.answerCbQuery();
  ctx.session.step = 'edit_id';
  ctx.reply("✏️ Фақат рақами ID-ро барои таҳрир нависед:");
});

// 7. ҚАБУЛИ МАТН ВА ИҶРО
bot.on('text', (ctx) => {
  const text = ctx.message.text;
  const step = ctx.session.step;

  if (['📋 Рӯйхат', '📝 Илова кардан', '📄 Содирот ба Word'].includes(text)) return;

  if (step === 'add_title') {
    ctx.session.tempTitle = text.trim();
    ctx.session.step = 'add_amount';
    ctx.reply("💰 Акнун **маблағ**-ро нависед (бо рақам):", { parse_mode: 'Markdown' });
  } 
  else if (step === 'add_amount') {
    const amount = parseFloat(text);
    if (isNaN(amount)) return ctx.reply("❌ Илтимос, фақат рақам нависед:");
    const title = ctx.session.tempTitle;
    const date = new Date().toLocaleDateString('tj-TJ');
    db.run("INSERT INTO expenses (user_id, title, amount, category, date) VALUES (?, ?, ?, ?, ?)",
      [ctx.from.id, title, amount, 'Дигар', date], () => {
        ctx.reply(`✅ Сабт шуд:\n📌 Ном: ${title}\n💵 Маблағ: ${amount} смн.`);
        ctx.session.step = 'none';
    });
  } 
  else if (step === 'delete') {
    const id = parseInt(text);
    if (isNaN(id)) return ctx.reply("❌ Илтимос, рақами ID нависед:");
    db.run("DELETE FROM expenses WHERE id = ? AND user_id = ?", [id, ctx.from.id], function(err) {
      if (this.changes > 0) ctx.reply(`✅ Хароҷоти ID ${id} нест карда шуд.`);
      else ctx.reply("⚠️ Чунин ID ёфт нашуд.");
      ctx.session.step = 'none';
    });
  } 
  else if (step === 'edit_id') {
    const id = parseInt(text);
    if (isNaN(id)) return ctx.reply("❌ Илтимос, рақами ID нависед:");
    ctx.session.editId = id;
    ctx.session.step = 'edit_title';
    ctx.reply(`✏️ Акнун **номи нави амалиёт**-ро нависед:`, { parse_mode: 'Markdown' });
  } 
  else if (step === 'edit_title') {
    ctx.session.editTitle = text.trim();
    ctx.session.step = 'edit_amount';
    ctx.reply("💰 Акнун **маблағи нав**-ро нависед:");
  }
  else if (step === 'edit_amount') {
    const amount = parseFloat(text);
    if (isNaN(amount)) return ctx.reply("❌ Илтимос, фақат рақам нависед:");
    const id = ctx.session.editId;
    const newTitle = ctx.session.editTitle;
    db.run("UPDATE expenses SET title = ?, amount = ? WHERE id = ? AND user_id = ?", 
      [newTitle, amount, id, ctx.from.id], function(err) {
        if (this.changes > 0) ctx.reply(`✅ ID ${id} нав карда шуд.`);
        else ctx.reply("⚠️ Чунин ID ёфт нашуд.");
        ctx.session.step = 'none';
    });
  }
  else {
    ctx.reply("💡 Лутфан аз меню амалро интихоб кунед.");
  }
});

// 8. ЁДРАСКУНӢ ВА API
cron.schedule('0 20 * * *', () => {
  db.all("SELECT DISTINCT user_id FROM expenses", [], (err, rows) => {
    if (rows) rows.forEach(row => bot.telegram.sendMessage(row.user_id, "🔔 Ёдраскунӣ: Оё имрӯз хароҷот доштед?").catch(() => {}));
  });
});

app.get('/api/expenses', (req, res) => {
  db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY id DESC", [req.query.user_id], (err, rows) => res.json(rows || []));
});

app.post('/api/add', (req, res) => {
  const { user_id, title, amount } = req.body;
  const date = new Date().toLocaleDateString('tj-TJ');
  db.run("INSERT INTO expenses (user_id, title, amount, category, date) VALUES (?, ?, ?, ?, ?)", [user_id, title, amount, 'Дигар', date], function() {
      bot.telegram.sendMessage(user_id, `📱 Аз веб-сайт илова шуд: ${title} - ${amount} смн.`).catch(() => {});
      res.json({ success: true, id: this.lastID });
  });
});

const PORT = process.env.PORT || 8100;
app.listen(PORT, () => {
  console.log(`🌐 API Server дар порти ${PORT} фаъол шуд.`);
  bot.launch().then(() => console.log(`✅ БОТ БО СОДИРОТ БА WORD ФАЪОЛ ШУД!`));
});

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));