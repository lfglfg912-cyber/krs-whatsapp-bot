import express from "express";
import dotenv from "dotenv";
import XLSX from "xlsx";
import crypto from "node:crypto";

dotenv.config();
const app = express();
const PORT = Number(process.env.PORT || 3000);
const acts = [];
const seen = new Set();
const apiVersion = process.env.GRAPH_API_VERSION || "v22.0";

app.use(express.json({
  limit: "2mb",
  verify: (req, res, buf) => { req.rawBody = Buffer.from(buf); }
}));

function allowed(from) {
  const list = (process.env.ALLOWED_SENDERS || "").split(",").map(x => x.trim()).filter(Boolean);
  return list.length === 0 || list.includes(String(from));
}
function verifySignature(req) {
  if (!process.env.WA_APP_SECRET) return true;
  const received = req.get("x-hub-signature-256") || "";
  const expected = "sha256=" + crypto.createHmac("sha256", process.env.WA_APP_SECRET).update(req.rawBody || Buffer.alloc(0)).digest("hex");
  try { return crypto.timingSafeEqual(Buffer.from(received), Buffer.from(expected)); } catch { return false; }
}
async function graph(path, options = {}) {
  const response = await fetch(`https://graph.facebook.com/${apiVersion}/${path}`, {
    ...options,
    headers: { Authorization: `Bearer ${process.env.WA_ACCESS_TOKEN}`, ...(options.headers || {}) }
  });
  const data = await response.json();
  if (!response.ok) throw new Error(data?.error?.message || `Meta API ${response.status}`);
  return data;
}
async function sendText(to, body) {
  if (!process.env.WA_ACCESS_TOKEN || !process.env.WA_PHONE_NUMBER_ID) {
    console.log("DEMO reply:", to, body);
    return;
  }
  await graph(`${process.env.WA_PHONE_NUMBER_ID}/messages`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      messaging_product: "whatsapp", recipient_type: "individual", to, type: "text",
      text: { preview_url: false, body }
    })
  });
}
async function getMedia(id) {
  const meta = await graph(id);
  const response = await fetch(meta.url, { headers: { Authorization: `Bearer ${process.env.WA_ACCESS_TOKEN}` } });
  if (!response.ok) throw new Error("Не удалось скачать файл из WhatsApp");
  return Buffer.from(await response.arrayBuffer());
}
function parseDate(text) {
  const m = String(text || "").match(/(\\d{1,2})[.\\-/](\\d{1,2})[.\\-/](\\d{4})/);
  return m ? `${m[3]}-${m[2].padStart(2,"0")}-${m[1].padStart(2,"0")}` : "";
}
function parseAct(buffer, filename, caption = "") {
  const wb = XLSX.read(buffer, { type: "buffer", cellDates: true });
  let farm = filename.replace(/\\.(xlsx|xls)$/i, "").replace(/^акт( приемки| приёмки)?\\s*/i, "").trim();
  let date = parseDate(filename) || parseDate(caption);
  let heads = null;
  let category = "";
  for (const sheetName of wb.SheetNames) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { header: 1, defval: "", raw: false });
    for (let i = 0; i < Math.min(rows.length, 30); i++) {
      const headers = rows[i].map(v => String(v || "").toLowerCase().trim());
      const countCol = headers.findIndex(v => /(колич|голов|кол-во|количество животных|итого голов)/i.test(v));
      const farmCol = headers.findIndex(v => /(хозяйств|поставщик|название кх|наименование кх)/i.test(v));
      const dateCol = headers.findIndex(v => /дата/i.test(v));
      const catCol = headers.findIndex(v => /(категор|вид животн|половозраст)/i.test(v));
      if (countCol < 0 && farmCol < 0 && dateCol < 0) continue;
      for (const row of rows.slice(i + 1)) {
        if (farmCol >= 0 && row[farmCol]) farm = String(row[farmCol]).trim();
        if (dateCol >= 0 && row[dateCol]) date = parseDate(row[dateCol]) || date;
        if (catCol >= 0 && row[catCol]) category = String(row[catCol]).trim();
        if (countCol >= 0 && row[countCol] !== "") {
          const n = Number(String(row[countCol]).replace(",", "."));
          if (Number.isFinite(n)) heads = n;
        }
      }
      if (heads !== null || farmCol >= 0) break;
    }
  }
  if (!farm) farm = "КХ (уточнить)";
  return { farm, date, heads, category, filename };
}
function makeReport() {
  if (!acts.length) return "Пока нет актов. Отправь Excel-файлы, затем напиши «итог».";
  const dates = [...new Set(acts.map(a => a.date).filter(Boolean))].sort();
  const dateText = dates.length ? dates.map(d => d.split("-").reverse().join(".")).join(" — ") : "дата не определена";
  const total = acts.reduce((s, a) => s + (Number.isFinite(a.heads) ? a.heads : 0), 0);
  const lines = acts.map((a, i) => `${i + 1}. КХ «${a.farm}» — ${Number.isFinite(a.heads) ? a.heads + " гол." : "количество не распознано"}${a.category ? " (" + a.category + ")" : ""}`);
  return `Закуп КРС за ${dateText}\\n\\n${lines.join("\\n")}\\n\\nВсего по распознанным актам: ${total} гол.\\nАктов: ${acts.length}\\n\\nПроверь данные: шаблоны Excel могут отличаться.`;
}
async function handleMessage(m) {
  const from = m.from;
  if (!from) return;
  if (!allowed(from)) return sendText(from, "Нет доступа к боту. Обратитесь к администратору.");
  if (m.type === "text") {
    const t = (m.text?.body || "").trim().toLowerCase();
    if (["итог", "/итог", "report", "/report"].includes(t)) return sendText(from, makeReport());
    if (["помощь", "/help", "меню"].includes(t)) return sendText(from, "Отправь Excel-акт (.xlsx/.xls), затем напиши «итог». Команда «сброс» очищает список.");
    if (["сброс", "/сброс", "очистить"].includes(t)) { acts.length = 0; seen.clear(); return sendText(from, "Список актов очищен."); }
    return sendText(from, "Отправь Excel-акт (.xlsx/.xls), а когда закончишь — напиши «итог».");
  }
  if (m.type === "document") {
    const doc = m.document || {};
    const filename = doc.filename || "akt.xlsx";
    if (!/\\.(xlsx|xls)$/i.test(filename)) return sendText(from, "Принимаю только Excel .xlsx или .xls.");
    if (seen.has(doc.id)) return sendText(from, "Этот файл уже добавлен. Напиши «итог» для сводки.");
    try {
      const act = parseAct(await getMedia(doc.id), filename, doc.caption || "");
      acts.push(act); seen.add(doc.id);
      return sendText(from, `Акт добавлен №${acts.length}\\nКХ: ${act.farm}\\nГолов: ${Number.isFinite(act.heads) ? act.heads : "не распознано"}\\nОтправь следующий акт или напиши «итог».`);
    } catch (e) {
      console.error(e);
      return sendText(from, "Не удалось прочитать Excel. Проверь файл и попробуй ещё раз.");
    }
  }
  return sendText(from, "Пока принимаю только Excel-акты и текстовую команду «итог».");
}

app.get("/webhook", (req, res) => {
  if (req.query["hub.mode"] === "subscribe" && req.query["hub.verify_token"] === process.env.WA_VERIFY_TOKEN) {
    return res.status(200).send(req.query["hub.challenge"]);
  }
  return res.sendStatus(403);
});
app.post("/webhook", (req, res) => {
  if (!verifySignature(req)) return res.sendStatus(401);
  res.sendStatus(200);
  for (const entry of req.body?.entry || []) {
    for (const change of entry.changes || []) {
      for (const message of change.value?.messages || []) {
        handleMessage(message).catch(err => console.error("Message error:", err));
      }
    }
  }
});
app.get("/", (_req, res) => res.type("text").send("KRS WhatsApp report bot is running."));
app.listen(PORT, () => console.log(`KRS bot listening on ${PORT}`));
