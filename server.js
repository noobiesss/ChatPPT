import express from "express";
import multer from "multer";
import mammoth from "mammoth";
import pdf from "pdf-parse";
import PptxGenJS from "pptxgenjs";
import rateLimit from "express-rate-limit";
import sharp from "sharp";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const KEY = process.env.GROQ_API_KEY;
const MODEL = process.env.GROQ_MODEL || "openai/gpt-oss-120b";
const UNSPLASH = process.env.UNSPLASH_ACCESS_KEY; // optional: enables real photos
if (!KEY) console.warn("Missing GROQ_API_KEY");

const app = express();
app.use(express.json({ limit: "2mb" }));
app.use(express.static(path.join(__dirname, "public")));
app.use("/api", rateLimit({ windowMs: 60_000, max: 8 }));
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

async function extractText(file) {
  const n = file.originalname.toLowerCase();
  if (n.endsWith(".docx")) return (await mammoth.extractRawText({ buffer: file.buffer })).value;
  if (n.endsWith(".pdf")) return (await pdf(file.buffer)).text;
  return file.buffer.toString("utf8");
}

const ICONS = ["lightbulb","target","users","book-open","rocket","shield","globe","clock","search","flask-conical",
  "code","database","heart","star","trending-up","settings","graduation-cap","scale","leaf","brain",
  "message-square","file-text","calendar","zap","lock","map-pin","award","puzzle","layers","camera"];

const SCHEMA = `{"reply":string,"title":string,"subtitle":string,"icon":string,"slides":[{"title":string,"bullets":[string],"notes":string,"icon":string,"image_query":string,"stat":{"value":string,"label":string}|null,"chart":{"type":"bar"|"pie"|"doughnut"|"line","title":string,"labels":[string],"values":[number]}|null}]}`;
const RULES = `"reply" = one friendly sentence on what you made or changed. "icon" must be one of: ${ICONS.join(", ")}. "image_query" = 1-3 concrete words for a stock photo. 3-4 bullets per slide, max 14 words each. Slide titles state the point; notes = 2-3 sentences for the presenter. Use "stat" or "chart" only when the source really has those numbers (max 2 each), otherwise null. Never invent data.`;

async function askGroq(text, o) {
  const system = o.deck
    ? `You edit an existing presentation. Apply the user's instruction and return ONLY the full updated JSON: ${SCHEMA}. Keep slides the user did not mention unchanged. ${RULES}`
    : `You turn coursework briefs and documents into presentation content. Return ONLY JSON: ${SCHEMA}. Exactly ${o.slides} content slides; cover every requirement and marking criterion in the source. Audience: ${o.audience}. Tone: ${o.tone}. ${RULES}`;
  const body = text.slice(0, 24000);
  const user = o.deck ? `CURRENT DECK:\n${JSON.stringify(o.deck)}\n\nINSTRUCTION AND ANY NEW MATERIAL:\n${body}` : body;
  const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: { Authorization: `Bearer ${KEY}`, "Content-Type": "application/json" },
    body: JSON.stringify({ model: MODEL, temperature: 0.4, response_format: { type: "json_object" },
      messages: [{ role: "system", content: system }, { role: "user", content: user }] }),
  });
  if (!r.ok) throw new Error(`Groq error ${r.status}: ${(await r.text()).slice(0, 200)}`);
  const d = JSON.parse((await r.json()).choices[0].message.content);
  d.slides = (d.slides || []).map(s => ({ ...s, bullets: (s.bullets || []).slice(0, 5) }));
  return d;
}

// ---------- visuals ----------
const pick = a => a[Math.floor(Math.random() * a.length)];
const PAL = [
  { dark: "12343B", light: "F4F6F7", a: "F2A900", b: "2E8B8B", text: "1E2B2E" },
  { dark: "3B1F4A", light: "FAF5F7", a: "FF6B6B", b: "8E5CC0", text: "2A1B33" },
  { dark: "0B2A4A", light: "F3F8FC", a: "19B5E8", b: "FF9F1C", text: "14233A" },
  { dark: "1F3D2B", light: "F5F8F2", a: "9BC53D", b: "E4572E", text: "1C2B21" },
  { dark: "222831", light: "F6F6F4", a: "FF7A00", b: "00ADB5", text: "222831" },
];
const FONTS = [["Georgia", "Calibri"], ["Trebuchet MS", "Calibri"], ["Cambria", "Calibri"]];

async function icon(name, color) {
  try {
    let f = path.join(__dirname, "node_modules/lucide-static/icons", `${name}.svg`);
    if (!ICONS.includes(name) || !fs.existsSync(f)) f = path.join(__dirname, "node_modules/lucide-static/icons/lightbulb.svg");
    const svg = fs.readFileSync(f, "utf8").replace(/currentColor/g, "#" + color)
      .replace(/width="24"/, 'width="256"').replace(/height="24"/, 'height="256"');
    return "image/png;base64," + (await sharp(Buffer.from(svg), { density: 300 }).png().toBuffer()).toString("base64");
  } catch { return null; }
}

async function photo(q) {
  if (!UNSPLASH || !q) return null;
  try {
    const r = await fetch(`https://api.unsplash.com/search/photos?per_page=5&orientation=landscape&query=${encodeURIComponent(q)}`,
      { headers: { Authorization: `Client-ID ${UNSPLASH}` } });
    const u = pick((await r.json()).results || [])?.urls?.small;
    if (!u) return null;
    return "image/jpeg;base64," + Buffer.from(await (await fetch(u)).arrayBuffer()).toString("base64");
  } catch { return null; }
}

const validChart = c => c && Array.isArray(c.labels) && Array.isArray(c.values) && c.labels.length >= 2 &&
  c.labels.length <= 8 && c.labels.length === c.values.length && c.values.every(v => typeof v === "number");

async function buildPptx(d) {
  const p = new PptxGenJS();
  p.layout = "LAYOUT_WIDE";
  const T = pick(PAL), [HF, BF] = pick(FONTS), bar = pick(["left", "top"]);
  const R = p.ShapeType.rect, E = p.ShapeType.ellipse;
  const X = bar === "left" ? 0.8 : 0.7, RIGHT = 12.6;
  const img = (sl, data, o) => data && sl.addImage({ data, ...o });

  // title slide
  const t = p.addSlide();
  t.background = { color: T.dark };
  for (let i = 0; i < 4; i++) {
    const s = 1.5 + Math.random() * 3.5;
    t.addShape(E, { x: Math.random() * 12, y: Math.random() * 6, w: s, h: s, fill: { color: pick([T.a, T.b]), transparency: 85 + Math.random() * 8 }, line: { type: "none" } });
  }
  t.addShape(R, { x: 0.7, y: 2.2, w: 0.12, h: 2.2, fill: { color: T.a } });
  t.addText(d.title, { x: 1.1, y: 2.1, w: 8.2, h: 1.6, fontFace: HF, fontSize: 38, bold: true, color: "FFFFFF", valign: "top" });
  t.addText(d.subtitle || "", { x: 1.1, y: 3.8, w: 8.2, h: 0.9, fontFace: BF, fontSize: 20, color: "D5E0E2", valign: "top" });
  t.addShape(E, { x: 9.6, y: 2.0, w: 3, h: 3, fill: { color: T.a } });
  img(t, await icon(d.icon, T.dark), { x: 10.2, y: 2.6, w: 1.8, h: 1.8 });

  let last = "";
  for (const [i, s] of d.slides.entries()) {
    const sl = p.addSlide(), B = s.bullets, n = B.length || 1;
    sl.background = { color: T.light };
    if (bar === "left") sl.addShape(R, { x: 0, y: 0, w: 0.25, h: 7.5, fill: { color: T.dark } });
    else sl.addShape(R, { x: 0, y: 0, w: 13.33, h: 0.2, fill: { color: T.dark } });
    sl.addText(s.title, { x: X, y: 0.4, w: 11.8, h: 1.0, fontFace: HF, fontSize: 28, bold: true, color: T.dark, valign: "top" });
    sl.addShape(R, { x: X, y: 1.45, w: 1.2, h: 0.07, fill: { color: T.a } });
    sl.addText(String(i + 1), { x: 12.2, y: 6.95, w: 0.6, h: 0.4, fontSize: 12, color: "7A8B8E", align: "right" });
    if (s.notes) sl.addNotes(s.notes);

    const list = (o) => sl.addText(B.map(b => ({ text: b, options: { bullet: { indent: 20 }, breakLine: true } })),
      { fontFace: BF, fontSize: 20, color: T.text, paraSpaceAfter: 14, valign: "top", ...o });

    let mode = validChart(s.chart) ? "chart" : s.stat?.value ? "stat"
      : pick(["icon", "cards", "steps", ...(UNSPLASH && s.image_query ? ["photo", "photo"] : [])]);
    if (mode === last && !["chart", "stat"].includes(mode)) mode = pick(["icon", "cards", "steps"].filter(m => m !== last));
    let ph = null;
    if (mode === "photo" && !(ph = await photo(s.image_query))) mode = "icon";
    last = mode;

    if (mode === "icon") {
      const cx = X + 1.9;
      sl.addShape(E, { x: cx - 1.7, y: 2.2, w: 3.4, h: 3.4, fill: { color: T.b, transparency: 75 }, line: { type: "none" } });
      sl.addShape(E, { x: cx - 1.35, y: 2.55, w: 2.7, h: 2.7, fill: { color: T.dark } });
      img(sl, await icon(s.icon, "FFFFFF"), { x: cx - 0.7, y: 3.2, w: 1.4, h: 1.4 });
      list({ x: X + 4.2, y: 2.0, w: RIGHT - X - 4.2, h: 4.7 });
    } else if (mode === "cards") {
      const g = 0.25, w = (RIGHT - X - g * (n - 1)) / n;
      for (const [k, b] of B.entries()) {
        const x = X + k * (w + g), c = k % 2 ? T.b : T.a;
        sl.addShape(R, { x, y: 2.0, w, h: 4.5, fill: { color: "FFFFFF" }, line: { color: "D9E0E2", width: 1 } });
        sl.addShape(R, { x, y: 2.0, w, h: 0.12, fill: { color: c } });
        sl.addText(String(k + 1), { x: x + 0.25, y: 2.4, w: 0.7, h: 0.7, shape: E, fill: { color: c }, align: "center", valign: "middle", fontSize: 20, bold: true, color: "FFFFFF" });
        sl.addText(b, { x: x + 0.2, y: 3.35, w: w - 0.4, h: 3, fontFace: BF, fontSize: n > 3 ? 16 : 18, color: T.text, valign: "top" });
      }
    } else if (mode === "steps") {
      const w = (RIGHT - X) / n;
      sl.addShape(p.ShapeType.line, { x: X + w / 2, y: 2.9, w: w * (n - 1), h: 0, line: { color: T.b, width: 3, dashType: "dash" } });
      for (const [k, b] of B.entries()) {
        const cx = X + k * w + w / 2;
        sl.addText(String(k + 1), { x: cx - 0.45, y: 2.45, w: 0.9, h: 0.9, shape: E, fill: { color: k % 2 ? T.b : T.a }, line: { color: T.light, width: 4 }, align: "center", valign: "middle", fontSize: 24, bold: true, color: "FFFFFF" });
        sl.addText(b, { x: cx - w / 2 + 0.15, y: 3.7, w: w - 0.3, h: 2.8, fontFace: BF, fontSize: n > 3 ? 16 : 18, color: T.text, align: "center", valign: "top" });
      }
    } else if (mode === "photo") {
      list({ x: X, y: 2.0, w: 6.4, h: 4.7 });
      sl.addShape(R, { x: 7.65, y: 2.05, w: 4.95, h: 4.65, fill: { color: T.a } });
      img(sl, ph, { x: 7.45, y: 1.85, w: 4.95, h: 4.65, sizing: { type: "cover", w: 4.95, h: 4.65 } });
    } else if (mode === "stat") {
      sl.addShape(R, { x: X, y: 1.9, w: 4.2, h: 4.7, fill: { color: T.dark } });
      img(sl, await icon(s.icon, T.a), { x: X + 0.4, y: 2.2, w: 0.8, h: 0.8 });
      sl.addText(String(s.stat.value), { x: X + 0.3, y: 3.1, w: 3.7, h: 1.6, fontFace: HF, fontSize: 54, bold: true, color: T.a, fit: "shrink", valign: "middle" });
      sl.addText(s.stat.label || "", { x: X + 0.3, y: 4.8, w: 3.6, h: 1.5, fontFace: BF, fontSize: 18, color: "FFFFFF", valign: "top" });
      list({ x: X + 4.7, y: 2.0, w: RIGHT - X - 4.7, h: 4.7 });
    } else if (mode === "chart") {
      const type = { bar: p.charts.BAR, pie: p.charts.PIE, doughnut: p.charts.DOUGHNUT, line: p.charts.LINE }[s.chart.type] || p.charts.BAR;
      sl.addChart(type, [{ name: s.chart.title || "Value", labels: s.chart.labels, values: s.chart.values }], {
        x: X, y: 1.9, w: 7, h: 4.8, chartColors: [T.a, T.b, T.dark, "8D99AE", "EF476F", "06D6A0", "FFD166", "118AB2"],
        showLegend: ["pie", "doughnut"].includes(s.chart.type), legendPos: "b", showValue: true, showTitle: !!s.chart.title, title: s.chart.title,
        titleFontSize: 14, dataLabelFontSize: 12, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelColor: T.text, valAxisLabelColor: T.text });
      list({ x: X + 7.3, y: 2.0, w: RIGHT - X - 7.3, h: 4.7, fontSize: 18 });
    }
  }
  return p.write({ outputType: "nodebuffer" });
}

app.post("/api/outline", upload.array("files", 5), async (req, res) => {
  try {
    let deck = null;
    try { deck = JSON.parse(req.body.deck); } catch {}
    let text = (req.body.text || "").trim();
    for (const f of req.files || []) text += `\n\n--- ${f.originalname} ---\n` + (await extractText(f));
    if (!deck && text.length < 30) return res.status(400).json({ error: "Add a brief, paste text, or attach a file first." });
    if (deck && !text) return res.status(400).json({ error: "Tell me what to change." });
    const slides = Math.min(Math.max(parseInt(req.body.slides) || 8, 3), 20);
    res.json(await askGroq(text, { deck, slides, audience: req.body.audience || "university lecturers", tone: req.body.tone || "clear and professional" }));
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post("/api/pptx", async (req, res) => {
  try {
    const buf = await buildPptx(req.body);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", 'attachment; filename="slides.pptx"');
    res.send(buf);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.listen(process.env.PORT || 3000, () => console.log("ChatPPT running at http://localhost:" + (process.env.PORT || 3000)));
