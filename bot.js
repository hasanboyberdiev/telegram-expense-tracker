const { Telegraf, Markup, session } = require('telegraf');
const sqlite3 = require('sqlite3').verbose();
const express = require('express');
const cors = require('cors');
const fs = require('fs');
const path = require('path');
const cron = require('node-cron');
const { Document, Packer, Paragraph, Table, TableCell, TableRow, AlignmentType, WidthType, HeadingLevel } = require('docx');
require('dotenv').config();

// 1. САНҶИШИ ТОКЕН
const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) {
  console.error("❌ Хатогӣ: BOT_TOKEN ёфт нашуд!");
  process.exit(1);
}
const bot = new Telegraf(TOKEN);

// 2. ХОТИРАИ БОТ ВА СЕРВЕРИ API
bot.use(session());
bot.use((ctx, next) => {
    ctx.session = ctx.session || { step: 'none' };
    return next();
});

const app = express();
app.use(cors());
app.use(express.json());

// ПАЙВАСТ КАРДАНИ ВЕБСАЙТ БА СЕРВЕР (Папкаи public)
app.use(express.static(path.join(__dirname, 'public')));

// 3. БАЗАИ МАЪЛУМОТ (SQLite)
const db = new sqlite3.Database('./database.db');
db.run(`CREATE TABLE IF NOT EXISTS expenses (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER,
  title TEXT,
  amount REAL,
  category TEXT DEFAULT 'Дигар',
  date TEXT
)`);

// 4. МЕНЮИ АСОСӢ (Бо суроғаи Alwaysdata)
function getKeyboard(userId) {
  const noCache = Date.now(); 
  // ЭЗОҲ: Агар логини шумо дар Alwaysdata 'shaxsi' набошад, онро иваз кунед!
  return Markup.keyboard([
    ['📝 Илова кардан', '📋 Рӯйхат'],
    ['📄 Содирот ба Word'], 
    [Markup.button.webApp('🌐 Web App -ро кушодан', `https://shaxsi.alwaysdata.net/?v=${noCache}&user_id=${userId}`)]
  ]).resize();
}

bot.start((ctx) => {
  ctx.session.step = 'none';
  ctx.reply("✨ Хуш омадед ба Назоратчии Буҷет! Амалро интихоб кунед:", getKeyboard(ctx.from.id));
});

// 5. РӮЙХАТ ВА ТУГМАҲОИ ТАҲРИР
bot.hears('📋 Рӯйхат', (ctx) => {
  ctx.session.step = 'none';
  db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY id DESC LIMIT 15", [ctx.from.id], (err, rows) => {
    if (err || rows.length === 0) return ctx.reply("📭 Рӯйхати шумо холӣ аст.", getKeyboard(ctx.from.id));
    
    let txt = "📋 <b>Рӯйхати хароҷоти охирин:</b>\n\n";
    rows.forEach(r => txt += `🆔 <b>ID: ${r.id}</b> | ${r.title} - ${r.amount} смн.\n`);
    
    ctx.reply(txt, { 
      parse_mode: 'HTML',
      ...Markup.inlineKeyboard([
        [Markup.button.callback('✏️ Таҳрир', 'action_edit'), Markup.button.callback('❌ Нест кардан', 'action_delete')]
      ])
    });
  });
});

// 6. СОДИРОТ БА WORD
bot.hears('📄 Содирот ба Word', (ctx) => {
  ctx.session.step = 'none';
  db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY date DESC", [ctx.from.id], async (err, rows) => {
    if (err || rows.length === 0) return ctx.reply("📭 Шумо ҳоло ягон хароҷот надоред.");

    try {
      let totalSum = 0;
      const tableRows = [];

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

      tableRows.push(
        new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "ҶАМЪИ УМУМӢ:", alignment: AlignmentType.RIGHT })], columnSpan: 2, shading: { fill: "D9D9D9" } }),
            new TableCell({ children: [new Paragraph({ text: totalSum.toString(), alignment: AlignmentType.CENTER })], shading: { fill: "A9D08E" } }),
            new TableCell({ children: [new Paragraph("")] }),
          ],
        })
      );

      const doc = new Document({
        sections: [{
          children: [
            new Paragraph({ text: "ҲИСОБОТИ МУФАССАЛИ ХАРОҶОТ", heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER }),
            new Paragraph({ text: "" }),
            new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }),
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

// 7. САРШАВИИ ҚАДАМҲО
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

// 8. ҚАБУЛИ МАТН ВА ИҶРОИ АМАЛҲО
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
        ctx.reply(`✅ Сабт шуд:\n📌 Ном: ${title}\n💵 Маблағ: ${amount} смн.`, getKeyboard(ctx.from.id));
        ctx.session.step = 'none';
    });
  } 
  else if (step === 'delete') {
    const id = parseInt(text);
    if (isNaN(id)) return ctx.reply("❌ Илтимос, рақами ID нависед:");
    
    db.run("DELETE FROM expenses WHERE id = ? AND user_id = ?", [id, ctx.from.id], function(err) {
      if (this.changes > 0) ctx.reply(`✅ Хароҷоти ID ${id} нест карда шуд.`, getKeyboard(ctx.from.id));
      else ctx.reply("⚠️ Чунин ID ёфт нашуд.", getKeyboard(ctx.from.id));
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
    
    db.run("UPDATE expenses SET title = ?, amount = ? WHERE id = ? AND user_id = ?", 
      [ctx.session.editTitle, amount, ctx.session.editId, ctx.from.id], function(err) {
        if (this.changes > 0) ctx.reply(`✅ ID ${ctx.session.editId} нав карда шуд.`, getKeyboard(ctx.from.id));
        else ctx.reply("⚠️ Чунин ID ёфт нашуд.", getKeyboard(ctx.from.id));
        ctx.session.step = 'none';
    });
  }
  else {
    ctx.reply("💡 Лутфан аз меню амалро интихоб кунед.", getKeyboard(ctx.from.id));
  }
});

// 9. API СЕРВЕР 
app.get('/api/expenses', (req, res) => {
  db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY id DESC LIMIT 20", [req.query.user_id], (err, rows) => res.json(rows || []));
});

app.post('/api/add', (req, res) => {
  const { user_id, title, amount } = req.body;
  const date = new Date().toLocaleDateString('tj-TJ');
  db.run("INSERT INTO expenses (user_id, title, amount, category, date) VALUES (?, ?, ?, ?, ?)", [user_id, title, amount, 'Дигар', date], function() {
      bot.telegram.sendMessage(user_id, `📱 Аз веб-сайт илова шуд: ${title} - ${amount} смн.`).catch(() => {});
      res.json({ success: true, id: this.lastID });
  });
});

// 10. ОҒОЗИ КОРИ СЕРВЕР ВА БОТ
const PORT = process.env.PORT || 8100;
app.listen(PORT, () => console.log(`🌐 API Server дар порти ${PORT} фаъол шуд.`));
bot.launch().then(() => console.log(`✅ БОТ ВА ВЕБСАЙТ БЕ ХАТО ФАЪОЛ ШУДАНД!`));

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));