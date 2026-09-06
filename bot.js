const { Telegraf, Markup, session } = require('telegraf');
const sqlite3 = require('sqlite3').verbose();
const fs = require('fs');
const { Document, Packer, Paragraph, Table, TableCell, TableRow, AlignmentType, WidthType, HeadingLevel } = require('docx');
require('dotenv').config();
const express = require('express'); 

const TOKEN = process.env.BOT_TOKEN;
if (!TOKEN) {
  console.error("❌ Хатогӣ: BOT_TOKEN ёфт нашуд!");
  process.exit(1);
}
const bot = new Telegraf(TOKEN);

bot.use(session());
bot.use((ctx, next) => {
    ctx.session = ctx.session || { step: 'none' };
    return next();
});

const db = new sqlite3.Database('./database.db');
db.run(`CREATE TABLE IF NOT EXISTS expenses (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  user_id INTEGER,
  title TEXT,
  amount REAL,
  category TEXT DEFAULT 'Дигар',
  date TEXT
)`);

// МЕНЮИ АСОСӢ (Фақат 3 тугмаи тоза)
function getKeyboard() {
  return Markup.keyboard([
    ['📝 Илова кардан', '📋 Рӯйхат'],
    ['📄 Содирот ба Word']
  ]).resize();
}

bot.start((ctx) => {
  ctx.session.step = 'none';
  ctx.reply("✨ Хуш омадед ба Назоратчии Буҷет! Амалро интихоб кунед:", getKeyboard());
});

async function sendWordDocument(userId) {
  return new Promise((resolve) => {
    db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY date DESC", [userId], async (err, rows) => {
      if (err || rows.length === 0) {
        bot.telegram.sendMessage(userId, "📭 Шумо ҳоло ягон хароҷот надоред.");
        return resolve();
      }
      try {
        let totalSum = 0;
        const tableRows = [];

        tableRows.push(new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "ID", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
            new TableCell({ children: [new Paragraph({ text: "Номи хароҷот", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
            new TableCell({ children: [new Paragraph({ text: "Маблағ (смн)", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
            new TableCell({ children: [new Paragraph({ text: "Сана", alignment: AlignmentType.CENTER })], shading: { fill: "4472C4" } }),
          ],
        }));

        rows.forEach((r) => {
          totalSum += r.amount;
          tableRows.push(new TableRow({
            children: [
              new TableCell({ children: [new Paragraph({ text: r.id.toString(), alignment: AlignmentType.CENTER })] }),
              new TableCell({ children: [new Paragraph(r.title)] }),
              new TableCell({ children: [new Paragraph({ text: r.amount.toString(), alignment: AlignmentType.CENTER })] }),
              new TableCell({ children: [new Paragraph({ text: r.date, alignment: AlignmentType.CENTER })] }),
            ],
          }));
        });

        tableRows.push(new TableRow({
          children: [
            new TableCell({ children: [new Paragraph({ text: "ҶАМЪИ УМУМӢ:", alignment: AlignmentType.RIGHT })], columnSpan: 2, shading: { fill: "D9D9D9" } }),
            new TableCell({ children: [new Paragraph({ text: totalSum.toString(), alignment: AlignmentType.CENTER })], shading: { fill: "A9D08E" } }),
            new TableCell({ children: [new Paragraph("")] }),
          ],
        }));

        const doc = new Document({
          sections: [{
            children: [
              new Paragraph({ text: "ҲИСОБОТИ МУФАССАЛИ ХАРОҶОТ", heading: HeadingLevel.HEADING_2, alignment: AlignmentType.CENTER }),
              new Paragraph({ text: "" }),
              new Table({ rows: tableRows, width: { size: 100, type: WidthType.PERCENTAGE } }),
            ],
          }],
        });

        const buffer = await Packer.toBuffer(doc);
        const fileName = `Hisobot_${userId}.docx`;
        fs.writeFileSync(fileName, buffer);

        await bot.telegram.sendDocument(userId,
          { source: fileName, filename: 'Ҳисоботи_Муфассал.docx' },
          { caption: "📄 <b>Ҳуҷҷати шумо омода аст!</b>", parse_mode: 'HTML' }
        );
        fs.unlinkSync(fileName); 
        resolve();
      } catch (e) {
        console.error(e);
        bot.telegram.sendMessage(userId, "❌ Хатогӣ ҳангоми сохтани файли Word.");
        resolve();
      }
    });
  });
}

bot.hears('📄 Содирот ба Word', async (ctx) => {
  ctx.reply("⏳ Ҳуҷҷати Word омода шуда истодааст...");
  await sendWordDocument(ctx.from.id);
});

// ҚИСМАТИ РӮЙХАТ БО ТУГМАҲОИ ДОХИЛӢ (INLINE KEYBOARD)
bot.hears('📋 Рӯйхат', (ctx) => {
  ctx.session.step = 'none';
  db.all("SELECT * FROM expenses WHERE user_id = ? ORDER BY id DESC LIMIT 15", [ctx.from.id], (err, rows) => {
    if (err || rows.length === 0) return ctx.reply("📭 Рӯйхати шумо холӣ аст.", getKeyboard());
    
    let txt = "📋 <b>Рӯйхати хароҷоти охирин:</b>\n\n";
    rows.forEach(r => txt += `🆔 <b>ID: ${r.id}</b> | ${r.title} - ${r.amount} смн.\n`);
    
    // Сохтани тугмаҳои зери хабар
    const inlineKeyboard = Markup.inlineKeyboard([
      Markup.button.callback('✏️ Таҳрир', 'action_edit'),
      Markup.button.callback('❌ Нест кардан', 'action_delete')
    ]);

    ctx.reply(txt, { parse_mode: 'HTML', ...inlineKeyboard }); 
  });
});

// АМАЛИЁТИ ТУГМАҲОИ ДОХИЛӢ
bot.action('action_edit', (ctx) => {
  ctx.session.step = 'edit_id';
  ctx.reply("✏️ Лутфан, **рақами ID**-и хароҷотро барои таҳрир нависед:", { parse_mode: 'Markdown' });
  ctx.answerCbQuery();
});

bot.action('action_delete', (ctx) => {
  ctx.session.step = 'delete_id';
  ctx.reply("❌ Лутфан, **рақами ID**-и хароҷотро нависед (метавонед аз рӯйхат бинед):", { parse_mode: 'Markdown' });
  ctx.answerCbQuery();
});

bot.hears('📝 Илова кардан', (ctx) => {
  ctx.session.step = 'add_title';
  ctx.reply("📝 Лутфан, **номи амалиёт**-ро нависед:", { parse_mode: 'Markdown' });
});

bot.on('text', (ctx) => {
  const text = ctx.message.text;
  if (['📋 Рӯйхат', '📝 Илова кардан', '📄 Содирот ба Word'].includes(text)) return;

  if (ctx.session.step === 'add_title') {
    ctx.session.tempTitle = text.trim();
    ctx.session.step = 'add_amount';
    ctx.reply("💰 Акнун **маблағ**-ро нависед (бо рақам):", { parse_mode: 'Markdown' });
  } 
  else if (ctx.session.step === 'add_amount') {
    const amount = parseFloat(text);
    if (isNaN(amount)) return ctx.reply("❌ Илтимос, фақат рақам нависед:");
    
    const date = new Date().toLocaleDateString('tj-TJ');
    db.run("INSERT INTO expenses (user_id, title, amount, category, date) VALUES (?, ?, ?, ?, ?)",
      [ctx.from.id, ctx.session.tempTitle, amount, 'Дигар', date], () => {
        ctx.reply(`✅ Сабт шуд: ${ctx.session.tempTitle} - ${amount} смн.`, getKeyboard());
        ctx.session.step = 'none';
    });
  } 
  else if (ctx.session.step === 'delete_id') {
    const id = parseInt(text);
    if (isNaN(id)) return ctx.reply("❌ Илтимос, танҳо рақам нависед.");
    db.run("DELETE FROM expenses WHERE id = ? AND user_id = ?", [id, ctx.from.id], function() {
        if (this.changes > 0) ctx.reply(`✅ Хароҷоти рақами ${id} нест карда шуд.`, getKeyboard());
        else ctx.reply("❌ Чунин ID ёфт нашуд.", getKeyboard());
        ctx.session.step = 'none';
    });
  }
  else if (ctx.session.step === 'edit_id') {
    const id = parseInt(text);
    if (isNaN(id)) return ctx.reply("❌ Илтимос, танҳо рақам нависед.");
    db.get("SELECT * FROM expenses WHERE id = ? AND user_id = ?", [id, ctx.from.id], (err, row) => {
        if (!row) {
            ctx.session.step = 'none';
            return ctx.reply("❌ Чунин ID ёфт нашуд.", getKeyboard());
        }
        ctx.session.editId = id;
        ctx.session.step = 'edit_title';
        ctx.reply(`Номи кунунӣ: ${row.title}\n✏️ **Номи навро** нависед:`, { parse_mode: 'Markdown' });
    });
  }
  else if (ctx.session.step === 'edit_title') {
    ctx.session.tempTitle = text.trim();
    ctx.session.step = 'edit_amount';
    ctx.reply("💰 Акнун **маблағи навро** нависед (бо рақам):", { parse_mode: 'Markdown' });
  }
  else if (ctx.session.step === 'edit_amount') {
    const amount = parseFloat(text);
    if (isNaN(amount)) return ctx.reply("❌ Илтимос, фақат рақам нависед:");
    db.run("UPDATE expenses SET title = ?, amount = ? WHERE id = ? AND user_id = ?",
       [ctx.session.tempTitle, amount, ctx.session.editId, ctx.from.id], function() {
       ctx.reply(`✅ Хароҷот бо муваффақият таҳрир карда шуд!`, getKeyboard());
       ctx.session.step = 'none';
    });
  } 
  else {
    ctx.reply("💡 Лутфан аз меню амалро интихоб кунед.", getKeyboard());
  }
});

// Сервери хурд
const app = express();
app.get('/', (req, res) => res.send('Бот фаъол аст!'));
const PORT = process.env.PORT || 8100;
app.listen(PORT, () => console.log(`🌐 Порти ${PORT} барои Alwaysdata банд карда шуд.`));

bot.launch().then(() => console.log(`✅ БОТ ФАЪОЛ ШУД!`));

process.once('SIGINT', () => bot.stop('SIGINT'));
process.once('SIGTERM', () => bot.stop('SIGTERM'));