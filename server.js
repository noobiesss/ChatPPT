import express from "express";
import multer from "multer";
import mammoth from "mammoth";
import pdf from "pdf-parse";
import PptxGenJS from "pptxgenjs";
import rateLimit from "express-rate-limit";

const KEY = process.env.GROQ_API_KEY;
const MODEL = process.env.GROQ_MODEL || "llama-3.3-70b-versatile";
if (!KEY) console.warn("Missing GROQ_API_KEY (set it as an environment variable)");

const app = express();
app.use(express.json({ limit: "2mb" }));
app.use(express.static("public"));
app.use("/api", rateLimit({ windowMs: 60_000, max: 8 })); // protects your key
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 10 * 1024 * 1024 } });

async function extractText(file) {
  const n = file.originalname.toLowerCase();
  if (n.endsWith(".docx")) return (await mammoth.extractRawText({ buffer: file.buffer })).value;
  if (n.endsWith(".pdf")) return (await pdf(file.buffer)).text;
  return file.buffer.toString("utf8"); // txt, md
}

async function askGroq(text, { slides, audience, tone }) {
  const system = `You turn coursework briefs and documents into clear presentation outlines.
Return ONLY JSON: {"title":string,"subtitle":string,"slides":[{"title":string,"bullets":[string],"notes":string}]}.
Rules: exactly ${slides} content slides; 3-5 bullets each, max 14 words per bullet; slide titles state the point;
cover every requirement/marking criterion in the source; notes = 2-3 sentences the presenter can say.
Audience: ${audience}. Tone: ${tone}. Never invent facts not supported by the source.`;
  const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: { Authorization: `Bearer ${KEY}`, "Content-Type": "application/json" },
    body: JSON.stringify({
      model: MODEL, temperature: 0.4, response_format: { type: "json_object" },
      messages: [{ role: "system", content: system }, { role: "user", content: text.slice(0, 24000) }],
    }),
  });
  if (!r.ok) throw new Error(`Groq error ${r.status}: ${(await r.text()).slice(0, 200)}`);
  return JSON.parse((await r.json()).choices[0].message.content);
}

const C = { ink: "12343B", paper: "F4F6F7", gold: "F2A900", text: "1E2B2E" };
async function buildPptx(d) {
  const p = new PptxGenJS();
  p.layout = "LAYOUT_WIDE";
  const t = p.addSlide();
  t.background = { color: C.ink };
  t.addShape(p.ShapeType.rect, { x: 0.7, y: 2.2, w: 0.12, h: 1.9, fill: { color: C.gold } });
  t.addText(d.title, { x: 1.1, y: 2.1, w: 11, h: 1.4, fontFace: "Georgia", fontSize: 40, bold: true, color: "FFFFFF", valign: "top" });
  t.addText(d.subtitle || "", { x: 1.1, y: 3.6, w: 11, h: 0.7, fontFace: "Calibri", fontSize: 20, color: "BFD0D3" });
  d.slides.forEach((s, i) => {
    const sl = p.addSlide();
    sl.background = { color: C.paper };
    sl.addShape(p.ShapeType.rect, { x: 0, y: 0, w: 0.25, h: 7.5, fill: { color: C.ink } });
    sl.addText(s.title, { x: 0.8, y: 0.5, w: 11.7, h: 1.1, fontFace: "Georgia", fontSize: 30, bold: true, color: C.ink, valign: "top" });
    sl.addText(s.bullets.map(b => ({ text: b, options: { bullet: { indent: 22 }, breakLine: true } })),
      { x: 0.8, y: 1.9, w: 11.5, h: 4.7, fontFace: "Calibri", fontSize: 22, color: C.text, paraSpaceAfter: 14, valign: "top" });
    sl.addText(String(i + 1), { x: 12.2, y: 6.9, w: 0.6, h: 0.4, fontSize: 12, color: "7A8B8E", align: "right" });
    if (s.notes) sl.addNotes(s.notes);
  });
  return p.write({ outputType: "nodebuffer" });
}

app.post("/api/outline", upload.array("files", 5), async (req, res) => {
  try {
    let text = (req.body.text || "").trim();
    for (const f of req.files || []) text += `\n\n--- ${f.originalname} ---\n` + (await extractText(f));
    if (text.length < 30) return res.status(400).json({ error: "Add a brief, paste text, or upload a file first." });
    const slides = Math.min(Math.max(parseInt(req.body.slides) || 8, 3), 20);
    res.json(await askGroq(text, { slides, audience: req.body.audience || "university lecturers", tone: req.body.tone || "clear and professional" }));
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

app.post("/api/pptx", async (req, res) => {
  try {
    const buf = await buildPptx(req.body);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", 'attachment; filename="slides.pptx"');
    res.send(buf);
  } catch (e) { res.status(500).json({ error: e.message }); }
});

app.listen(process.env.PORT || 3000, () => console.log("Slidesmith running"));
