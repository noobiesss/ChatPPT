import "dotenv/config";
import dns from "dns";
dns.setDefaultResultOrder("ipv4first");
import express from "express";
import multer from "multer";
import mammoth from "mammoth";
import pdf from "pdf-parse";
import PptxGenJS from "pptxgenjs";
import rateLimit from "express-rate-limit";
import sharp from "sharp";
import JSZip from "jszip";
import crypto from "crypto";
import { analyzeTemplate as analyzeTpl, buildFromTemplate } from "./template.js";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const KEY = process.env.GROQ_API_KEY;
const MODEL = process.env.GROQ_MODEL || "openai/gpt-oss-120b";
const UNSPLASH = process.env.UNSPLASH_ACCESS_KEY; // optional: enables real photos
if (!KEY) console.warn("Missing GROQ_API_KEY");

const app = express();
app.use(express.json({ limit: "10mb" }));
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

const DECK = `{"title":string,"subtitle":string,"icon":string,"slides":[{"title":string,"bullets":[string],"notes":string,"icon":string,"image_query":string,"stat":{"value":string,"label":string}|null,"chart":{"type":"bar"|"pie"|"doughnut"|"line","title":string,"labels":[string],"values":[number]}|null}]}`;
const RULES = `Deck rules: "icon" must be one of: ${ICONS.join(", ")}. "image_query" = 1-3 concrete words for a stock photo. 3-4 bullets per slide, max 14 words each. Slide titles state the point; notes = 2-3 sentences for the presenter. Use "stat" or "chart" only when the source really has those numbers (max 2 each), otherwise null. Never invent data. With no source material, use well-established general knowledge and avoid specific statistics.`;

async function askGroq(o) {
  const system = `You are ChatPPT, a friendly assistant that both answers questions and builds PowerPoint decks for coursework.
Return ONLY JSON: {"reply": string, "deck": ${DECK} | null}.
Decide from the user's message:
1. Question, chat, advice or explanation (for example about the assignment, the slides, or the topic): answer helpfully in "reply" (plain text, up to about 150 words, short paragraphs), use the SOURCE MATERIAL when relevant, and set "deck" to null.
2. Request to make slides, or a pasted/attached brief with no other instruction: build a full deck with exactly ${o.slides} content slides (or the number the user asks for), covering every requirement and marking criterion in the source. Audience: ${o.audience}. Tone: ${o.tone}.
3. Request to change the CURRENT DECK: return the FULL updated deck and keep slides the user did not mention unchanged.
When you return a deck, "reply" is one friendly sentence about what you made or changed.
${RULES}`;
  const user = [
    o.source && `SOURCE MATERIAL:\n${o.source.slice(0, 20000)}`,
    o.deck && `CURRENT DECK:\n${JSON.stringify(o.deck)}`,
    o.history.length && `RECENT CHAT:\n${o.history.slice(-6).map(m => `${m.role}: ${String(m.text).slice(0, 600)}`).join("\n")}`,
    `USER MESSAGE:\n${o.text || "(files attached, no message)"}`,
  ].filter(Boolean).join("\n\n");
  const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: { Authorization: `Bearer ${KEY}`, "Content-Type": "application/json" },
    body: JSON.stringify({ model: MODEL, temperature: 0.4, response_format: { type: "json_object" },
      messages: [{ role: "system", content: system }, { role: "user", content: user }] }),
  });
  if (!r.ok) throw new Error(`Groq error ${r.status}: ${(await r.text()).slice(0, 200)}`);
  const out = JSON.parse((await r.json()).choices[0].message.content);
  const reply = String(out.reply || "").trim();
  if (!out.deck) return { reply: reply || "I'm here. Ask me a question, or tell me what slides to make.", deck: null };
  const good = s => s && s.title && Array.isArray(s.bullets) && s.bullets.some(b => String(b).trim());
  const slides = (out.deck.slides || []).filter(good).map(s => ({ ...s, bullets: s.bullets.slice(0, 5) }));
  if (!slides.length) return { reply: reply || "I couldn't build slides from that. Could you tell me more?", deck: null };
  if (o.deck && slides.length < o.deck.slides.length / 2 && !/(fewer|shorter|reduce|cut|remove|delete|only)/i.test(o.text))
    return { reply: "That change looked wrong, so I kept your deck as it was. Could you rephrase it?", deck: null };
  const base = o.deck || {};
  return { reply: reply || "Here is your deck.", deck: { ...base, ...out.deck, title: out.deck.title || base.title || "Untitled deck",
    subtitle: out.deck.subtitle ?? base.subtitle ?? "", icon: out.deck.icon || base.icon, slides } };
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

// ---------- template analysis (.pptx exported from Canva or PowerPoint) ----------
const lum = h => [0, 2, 4].reduce((s, i, k) => s + parseInt(h.slice(i, i + 2), 16) / 255 * [0.2126, 0.7152, 0.0722][k], 0);
const sat = h => { const v = [0, 2, 4].map(i => parseInt(h.slice(i, i + 2), 16)); return (Math.max(...v) - Math.min(...v)) / 255; };
const dist = (a, b) => [0, 2, 4].reduce((s, i) => s + Math.abs(parseInt(a.slice(i, i + 2), 16) - parseInt(b.slice(i, i + 2), 16)), 0);

async function analyzeTemplate(buf, name) {
  const zip = await JSZip.loadAsync(buf);
  const files = Object.keys(zip.files);
  if (!files.some(f => f.startsWith("ppt/slides/") || f.startsWith("ppt/theme/"))) throw new Error("this is not a PowerPoint file");
  let all = "";
  for (const f of files.filter(f => /^ppt\/(slides\/slide\d+|theme\/theme1)\.xml$/.test(f))) all += " " + await zip.file(f).async("string");
  const top = (re) => { const m = {}; for (const x of all.matchAll(re)) m[x[1]] = (m[x[1]] || 0) + 1; return Object.entries(m).sort((a, b) => b[1] - a[1]).map(e => e[0]); };
  const uniq = [];
  for (const c of top(/srgbClr val="([0-9A-Fa-f]{6})"/g).map(c => c.toUpperCase())) if (uniq.every(u => dist(u, c) > 60)) uniq.push(c);
  const fonts = [...new Set(top(/<a:latin typeface="([^"+][^"]*)"/g))].slice(0, 2);
  const fb = pick(PAL);
  const dark = uniq.filter(c => lum(c) < 0.35).sort((a, b) => lum(a) - lum(b))[0];
  const light = uniq.filter(c => lum(c) > 0.88).sort((a, b) => lum(b) - lum(a))[0];
  const acc = uniq.filter(c => sat(c) > 0.25 && lum(c) > 0.15 && lum(c) < 0.9).sort((a, b) => sat(b) - sat(a));
  const palette = { dark: dark || fb.dark, light: light || "F7F7F7", a: acc[0] || fb.a, b: acc[1] || acc[0] || fb.b, text: dark || "222222" };
  const imgs = [];
  const media = [];
  for (const f of files.filter(f => /^ppt\/media\/.+\.(png|jpe?g)$/i.test(f))) media.push(await zip.file(f).async("nodebuffer"));
  for (const b of media.filter(b => b.length > 40000).sort((a, b) => b.length - a.length).slice(0, 2)) {
    try { imgs.push("image/jpeg;base64," + (await sharp(b).resize({ width: 1100, withoutEnlargement: true }).jpeg({ quality: 60 }).toBuffer()).toString("base64")); } catch {}
  }
  return { name, palette, fonts, images: imgs, swatches: [palette.dark, palette.a, palette.b, palette.light] };
}

const shuffle = a => [...a].sort(() => Math.random() - 0.5);
const ALL = ["split", "cards", "rows", "timeline", "banner", "quote", "zigzag", "photo"];

async function buildPptx(d) {
  const p = new PptxGenJS();
  p.layout = "LAYOUT_WIDE";
  const th = d.theme, [rhf, rbf] = pick(FONTS);
  const T = th?.palette ? { ...th.palette } : pick(PAL), frame = pick(["side", "top", "none"]);
  const HF = th?.fonts?.[0] || rhf, BF = th?.fonts?.[1] || th?.fonts?.[0] || rbf;
  const R = p.ShapeType.rect, E = p.ShapeType.ellipse, BOX = Math.random() < 0.5 ? p.ShapeType.roundRect : R;
  const W = 13.33, H = 7.5, M = 0.7, G = 0.25;
  const LIGHT = { bg: T.light, fg: T.text, ti: T.dark, fill: { color: "FFFFFF" }, line: { color: "D9E0E2", width: 1 } };
  const DARK = { bg: T.dark, fg: "EEF3F4", ti: "FFFFFF", fill: { color: "FFFFFF", transparency: 90 }, line: { type: "none" } };
  const img = (sl, data, o) => data && sl.addImage({ data, ...o });
  const size = B => { const t = B.join("").length; return t < 140 ? 28 : t < 220 ? 25 : t < 320 ? 22 : 20; };

  // ----- title slide (2 variants) -----
  const t = p.addSlide();
  const timg = th?.images?.length ? pick(th.images) : null;
  if (timg && Math.random() < 0.7) {
    t.background = { color: T.dark };
    img(t, timg, { x: 0, y: 0, w: W, h: H, sizing: { type: "cover", w: W, h: H } });
    t.addShape(R, { x: 0, y: 0, w: W, h: H, fill: { color: T.dark, transparency: 30 } });
    t.addShape(R, { x: 0.7, y: 2.4, w: 0.12, h: 2.2, fill: { color: T.a } });
    t.addText(d.title, { x: 1.1, y: 2.3, w: 10, h: 1.6, fontFace: HF, fontSize: 40, bold: true, color: "FFFFFF", valign: "top", fit: "shrink" });
    t.addText(d.subtitle || "", { x: 1.1, y: 4.0, w: 10, h: 0.9, fontFace: BF, fontSize: 20, color: "E6EDEE", valign: "top" });
  } else if (Math.random() < 0.5) {
    t.background = { color: T.dark };
    for (let i = 0; i < 4; i++) {
      const s = 1.5 + Math.random() * 3.5;
      t.addShape(E, { x: Math.random() * 12, y: Math.random() * 6, w: s, h: s, fill: { color: pick([T.a, T.b]), transparency: 85 + Math.random() * 8 }, line: { type: "none" } });
    }
    t.addShape(R, { x: 0.7, y: 2.2, w: 0.12, h: 2.2, fill: { color: T.a } });
    t.addText(d.title, { x: 1.1, y: 2.1, w: 8.2, h: 1.6, fontFace: HF, fontSize: 38, bold: true, color: "FFFFFF", valign: "top", fit: "shrink" });
    t.addText(d.subtitle || "", { x: 1.1, y: 3.8, w: 8.2, h: 0.9, fontFace: BF, fontSize: 20, color: "D5E0E2", valign: "top" });
    t.addShape(E, { x: 9.6, y: 2.0, w: 3, h: 3, fill: { color: T.a } });
    img(t, await icon(d.icon, T.dark), { x: 10.2, y: 2.6, w: 1.8, h: 1.8 });
  } else {
    t.background = { color: T.light };
    t.addShape(R, { x: 8.4, y: 0, w: W - 8.4, h: H, fill: { color: T.dark } });
    t.addShape(E, { x: 9.3, y: 2.1, w: 3.3, h: 3.3, fill: { color: T.a } });
    img(t, await icon(d.icon, T.dark), { x: 10.05, y: 2.85, w: 1.8, h: 1.8 });
    t.addShape(R, { x: M, y: 2.0, w: 1.2, h: 0.1, fill: { color: T.a } });
    t.addText(d.title, { x: M, y: 2.3, w: 7.2, h: 2, fontFace: HF, fontSize: 40, bold: true, color: T.dark, valign: "top", fit: "shrink" });
    t.addText(d.subtitle || "", { x: M, y: 4.4, w: 7.2, h: 1, fontFace: BF, fontSize: 20, color: T.text, valign: "top" });
  }

  // ----- content slides -----
  const bag = []; let last = "";
  const choose = ok => {
    let pool = bag.filter(m => ok.includes(m) && m !== last);
    if (!pool.length) { bag.length = 0; bag.push(...shuffle(ALL)); pool = bag.filter(m => ok.includes(m) && m !== last); }
    bag.splice(bag.indexOf(pool[0]), 1);
    return pool[0];
  };

  for (const [i, s] of d.slides.entries()) {
    const sl = p.addSlide(), B = s.bullets.length ? s.bullets : [s.title], n = B.length, fz = size(B), sm = Math.max(16, fz - 3);
    const ok = ALL.filter(m => (m !== "quote" || n >= 3) && (m !== "photo" || th?.images?.length || (UNSPLASH && s.image_query)));
    let mode = validChart(s.chart) ? "chart" : s.stat?.value ? "stat" : choose(ok);
    let ph = null;
    if (mode === "photo" && !(ph = th?.images?.length ? pick(th.images) : await photo(s.image_query))) mode = "rows";
    last = mode;
    const c = ["cards", "rows", "timeline", "zigzag"].includes(mode) && Math.random() < 0.3 ? DARK : LIGHT;
    sl.background = { color: c.bg };
    if (s.notes) sl.addNotes(s.notes);
    if (mode !== "photo") sl.addText(String(i + 1), { x: W - 0.9, y: H - 0.5, w: 0.6, h: 0.35, fontSize: 12, color: "8A9A9D", align: "right" });

    const num = (k, x, y, dm, col) => sl.addText(String(k), { x, y, w: dm, h: dm, shape: E, fill: { color: col }, align: "center", valign: "middle", fontSize: dm * 24, bold: true, color: "FFFFFF" });
    const card = (x, y, w, h, o = {}) => sl.addShape(BOX, { x, y, w, h, fill: c.fill, line: c.line, rectRadius: 0.12, ...o });
    const list = (items, o, col = c.fg, f = fz) => sl.addText(items.map(b => ({ text: b, options: { bullet: { indent: 24 }, breakLine: true } })),
      { fontFace: BF, fontSize: f, color: col, paraSpaceAfter: 16, valign: "middle", ...o });
    const header = () => {
      if (frame === "side") sl.addShape(R, { x: 0, y: 0, w: 0.25, h: H, fill: { color: c === DARK ? T.a : T.dark } });
      else if (frame === "top") sl.addShape(R, { x: 0, y: 0, w: W, h: 0.2, fill: { color: c === DARK ? T.a : T.dark } });
      sl.addText(s.title, { x: M + (frame === "side" ? 0.2 : 0), y: 0.4, w: W - 2 * M, h: 1.0, fontFace: HF, fontSize: 30, bold: true, color: c.ti, valign: "top", fit: "shrink" });
      sl.addShape(R, { x: M + (frame === "side" ? 0.2 : 0), y: 1.45, w: 1.2, h: 0.07, fill: { color: T.a } });
    };
    const Y0 = 1.85, AH = H - Y0 - 0.55;

    if (mode === "split") {
      sl.addShape(R, { x: 0, y: 0, w: 4.9, h: H, fill: { color: T.dark } });
      sl.addShape(E, { x: 0.7, y: 0.7, w: 1.3, h: 1.3, fill: { color: T.a } });
      img(sl, await icon(s.icon, T.dark), { x: 0.97, y: 0.97, w: 0.76, h: 0.76 });
      sl.addText(s.title, { x: 0.6, y: 2.3, w: 3.8, h: 4.2, fontFace: HF, fontSize: 32, bold: true, color: "FFFFFF", valign: "top", fit: "shrink" });
      list(B, { x: 5.5, y: 0.6, w: 7.1, h: 6.3 });
    } else if (mode === "cards") {
      header();
      if (n === 3 || n === 5) {
        const w = (W - 2 * M - G * (n - 1)) / n;
        B.forEach((b, k) => {
          const x = M + k * (w + G);
          card(x, Y0, w, AH);
          sl.addText(String(k + 1).padStart(2, "0"), { x: x + 0.2, y: Y0 + 0.15, w: w - 0.4, h: 1.0, fontFace: HF, fontSize: 48, bold: true, color: k % 2 ? T.b : T.a });
          sl.addText(b, { x: x + 0.2, y: Y0 + 1.3, w: w - 0.4, h: AH - 1.5, fontFace: BF, fontSize: sm, color: c.fg, valign: "top" });
        });
      } else {
        const rows = Math.ceil(n / 2), w = (W - 2 * M - G) / 2, h = (AH - G * (rows - 1)) / rows;
        B.forEach((b, k) => {
          const x = M + (k % 2) * (w + G), y = Y0 + Math.floor(k / 2) * (h + G);
          card(x, y, w, h);
          sl.addShape(R, { x, y: y + 0.15, w: 0.12, h: h - 0.3, fill: { color: k % 2 ? T.b : T.a } });
          num(k + 1, x + 0.35, y + (h - 0.75) / 2, 0.75, k % 2 ? T.b : T.a);
          sl.addText(b, { x: x + 1.3, y, w: w - 1.5, h, fontFace: BF, fontSize: sm, color: c.fg, valign: "middle" });
        });
      }
    } else if (mode === "rows") {
      header();
      const rh = AH / n;
      B.forEach((b, k) => {
        const y = Y0 + k * rh;
        sl.addText(String(k + 1).padStart(2, "0"), { x: M, y, w: 1.4, h: rh, fontFace: HF, fontSize: 40, bold: true, color: k % 2 ? T.b : T.a, valign: "middle" });
        sl.addText(b, { x: M + 1.6, y, w: W - 2 * M - 1.6, h: rh, fontFace: BF, fontSize: fz - 1, color: c.fg, valign: "middle" });
        if (k < n - 1) sl.addShape(p.ShapeType.line, { x: M, y: y + rh, w: W - 2 * M, h: 0, line: { color: c === DARK ? "4A6A70" : "D9E0E2", width: 1 } });
      });
    } else if (mode === "timeline") {
      header();
      const rh = AH / n;
      sl.addShape(p.ShapeType.line, { x: M + 0.45, y: Y0 + rh / 2, w: 0, h: rh * (n - 1), line: { color: T.b, width: 3 } });
      B.forEach((b, k) => {
        const cy = Y0 + rh * (k + 0.5);
        num(k + 1, M + 0.05, cy - 0.4, 0.8, k % 2 ? T.b : T.a);
        card(M + 1.3, cy - rh / 2 + 0.08, W - 2 * M - 1.3, rh - 0.16);
        sl.addText(b, { x: M + 1.6, y: cy - rh / 2 + 0.08, w: W - 2 * M - 1.9, h: rh - 0.16, fontFace: BF, fontSize: sm + 1, color: c.fg, valign: "middle" });
      });
    } else if (mode === "banner") {
      sl.addShape(R, { x: 0, y: 0, w: W, h: 2.4, fill: { color: T.dark } });
      sl.addShape(R, { x: 0, y: 2.4, w: W, h: 0.1, fill: { color: T.a } });
      sl.addShape(E, { x: W - 2.5, y: 0.45, w: 1.5, h: 1.5, fill: { color: T.a } });
      img(sl, await icon(s.icon, T.dark), { x: W - 2.17, y: 0.78, w: 0.84, h: 0.84 });
      sl.addText(s.title, { x: M, y: 0.4, w: W - 2 * M - 2.6, h: 1.6, fontFace: HF, fontSize: 34, bold: true, color: "FFFFFF", valign: "middle", fit: "shrink" });
      const rows = Math.ceil(n / 2), w = (W - 2 * M - 0.4) / 2, h = 4.2 / rows;
      B.forEach((b, k) => {
        const x = M + (k % 2) * (w + 0.4), y = 2.85 + Math.floor(k / 2) * h;
        sl.addShape(E, { x, y: y + 0.12, w: 0.28, h: 0.28, fill: { color: k % 2 ? T.b : T.a } });
        sl.addText(b, { x: x + 0.5, y, w: w - 0.5, h: h - 0.1, fontFace: BF, fontSize: sm + 1, color: c.fg, valign: "top" });
      });
    } else if (mode === "quote") {
      sl.addShape(R, { x: 0, y: 0, w: 6.4, h: H, fill: { color: T.a } });
      sl.addText("\u201C", { x: 0.5, y: 0.1, w: 2, h: 2, fontFace: HF, fontSize: 130, bold: true, color: T.dark });
      sl.addText(B[0], { x: 0.7, y: 1.9, w: 5.1, h: 4.6, fontFace: HF, fontSize: 28, bold: true, italic: true, color: T.dark, valign: "top", fit: "shrink" });
      sl.addText(s.title, { x: 7, y: 0.5, w: 5.6, h: 1.2, fontFace: HF, fontSize: 28, bold: true, color: T.dark, valign: "top", fit: "shrink" });
      list(B.slice(1), { x: 7, y: 1.9, w: 5.6, h: 4.9 }, c.fg, Math.max(18, fz - 2));
    } else if (mode === "zigzag") {
      header();
      const rh = AH / n, w = 8.2;
      B.forEach((b, k) => {
        const x = k % 2 ? W - M - w : M, y = Y0 + k * rh;
        card(x, y + 0.05, w, rh - 0.15);
        num(k + 1, k % 2 ? x + w - 1.0 : x + 0.2, y + (rh - 0.8) / 2, 0.8, k % 2 ? T.b : T.a);
        sl.addText(b, { x: k % 2 ? x + 0.3 : x + 1.2, y: y + 0.05, w: w - 1.5, h: rh - 0.15, fontFace: BF, fontSize: sm + 1, color: c.fg, valign: "middle" });
      });
    } else if (mode === "photo") {
      sl.addShape(R, { x: 6.75, y: 0, w: 0.2, h: H, fill: { color: T.a } });
      img(sl, ph, { x: 6.95, y: 0, w: W - 6.95, h: H, sizing: { type: "cover", w: W - 6.95, h: H } });
      sl.addText(s.title, { x: M, y: 0.5, w: 5.7, h: 1.5, fontFace: HF, fontSize: 30, bold: true, color: T.dark, valign: "top", fit: "shrink" });
      sl.addShape(R, { x: M, y: 2.0, w: 1.2, h: 0.07, fill: { color: T.a } });
      list(B, { x: M, y: 2.3, w: 5.7, h: 4.7 }, c.fg, sm + 1);
    } else if (mode === "stat") {
      sl.addShape(R, { x: 0, y: 0, w: 5.2, h: H, fill: { color: T.dark } });
      img(sl, await icon(s.icon, T.a), { x: 0.7, y: 0.7, w: 0.9, h: 0.9 });
      sl.addText(String(s.stat.value), { x: 0.6, y: 1.9, w: 4.2, h: 2.6, fontFace: HF, fontSize: 66, bold: true, color: T.a, fit: "shrink", valign: "middle" });
      sl.addText(s.stat.label || "", { x: 0.6, y: 4.6, w: 4.2, h: 2.2, fontFace: BF, fontSize: 20, color: "FFFFFF", valign: "top" });
      sl.addText(s.title, { x: 5.9, y: 0.5, w: 6.8, h: 1.2, fontFace: HF, fontSize: 28, bold: true, color: T.dark, valign: "top", fit: "shrink" });
      list(B, { x: 5.9, y: 1.9, w: 6.8, h: 5 });
    } else if (mode === "chart") {
      header();
      const type = { bar: p.charts.BAR, pie: p.charts.PIE, doughnut: p.charts.DOUGHNUT, line: p.charts.LINE }[s.chart.type] || p.charts.BAR;
      card(M, Y0, 7.6, AH);
      sl.addChart(type, [{ name: s.chart.title || "Value", labels: s.chart.labels, values: s.chart.values }], {
        x: M + 0.1, y: Y0 + 0.1, w: 7.4, h: AH - 0.2, chartColors: [T.a, T.b, T.dark, "8D99AE", "EF476F", "06D6A0", "FFD166", "118AB2"],
        showLegend: ["pie", "doughnut"].includes(s.chart.type), legendPos: "b", showValue: true, showTitle: !!s.chart.title, title: s.chart.title,
        titleFontSize: 14, dataLabelFontSize: 12, catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, catAxisLabelColor: T.text, valAxisLabelColor: T.text });
      const rh = AH / n, x = M + 7.9, w = W - M - x;
      B.forEach((b, k) => {
        const y = Y0 + k * rh;
        card(x, y + 0.04, w, rh - 0.12);
        sl.addShape(R, { x, y: y + 0.04, w: 0.1, h: rh - 0.12, fill: { color: k % 2 ? T.b : T.a } });
        sl.addText(b, { x: x + 0.25, y: y + 0.04, w: w - 0.4, h: rh - 0.12, fontFace: BF, fontSize: Math.min(sm, 18), color: c.fg, valign: "middle" });
      });
    }
  }
  return p.write({ outputType: "nodebuffer" });
}

app.post("/api/outline", upload.array("files", 5), async (req, res) => {
  try {
    const J = k => { try { return JSON.parse(req.body[k]); } catch { return null; } };
    const deck = J("deck"), history = Array.isArray(J("history")) ? J("history") : [];
    const text = (req.body.text || "").trim();
    let material = "";
    for (const f of req.files || []) material += `\n\n--- ${f.originalname} ---\n` + (await extractText(f));
    if (text.length >= 200) material += "\n\n" + text; // a long pasted brief counts as source material
    if (!text && !material.trim()) return res.status(400).json({ error: "Type a message or attach a file first." });
    const source = ((req.body.source || "") + material).slice(0, 30000);
    const slides = Math.min(Math.max(parseInt(req.body.slides) || 8, 3), 20);
    const out = await askGroq({ text, source, deck, history, slides, audience: req.body.audience || "university lecturers", tone: req.body.tone || "clear and professional" });
    res.json({ ...out, source });
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// ---------- template mode: fill the user's own .pptx ----------
const TPL_DIR = path.join(__dirname, "templates");
fs.mkdirSync(TPL_DIR, { recursive: true });
const tplUpload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 100 * 1024 * 1024 } });
const photoCache = new Map();

async function getPhotos(query, count) {
  if (!UNSPLASH || !query) return [];
  let urls = photoCache.get(query);
  if (!urls) {
    const r = await fetch(`https://api.unsplash.com/search/photos?per_page=8&orientation=landscape&query=${encodeURIComponent(query)}`, { headers: { Authorization: `Client-ID ${UNSPLASH}` } });
    urls = ((await r.json()).results || []).map(x => x.urls.regular);
    photoCache.set(query, urls);
  }
  const chosen = shuffle(urls).slice(0, count);
  return (await Promise.all(chosen.map(async u => Buffer.from(await (await fetch(u)).arrayBuffer())))).filter(b => b.length);
}

async function askFills(items) {
  const system = `You write the text for the slots of a PowerPoint template. For each slide, use its content to fill EVERY slot.
Return ONLY JSON: {"slides":[{"fills":{"<slot key>":"text"}}]} with one entry per slide, in the same order.
Roles: "title" = the slide title; "label" = a 2-4 word heading; "body" = one or two short sentences.
Never exceed a slot's maxChars and aim near the length of its sample. Use only facts from the slide content: with more slots than bullets, split or expand a bullet logically; with fewer, merge. Never leave a slot empty or copy the sample text.`;
  const r = await fetch("https://api.groq.com/openai/v1/chat/completions", {
    method: "POST",
    headers: { Authorization: `Bearer ${KEY}`, "Content-Type": "application/json" },
    body: JSON.stringify({ model: MODEL, temperature: 0.3, response_format: { type: "json_object" },
      messages: [{ role: "system", content: system }, { role: "user", content: JSON.stringify(items) }] }),
  });
  if (!r.ok) throw new Error(`Groq error ${r.status}`);
  const arr = JSON.parse((await r.json()).choices[0].message.content).slides;
  if (!Array.isArray(arr) || arr.length !== items.length) throw new Error("unexpected fill response");
  return arr.map(s => s.fills || {});
}

app.post("/api/template", tplUpload.single("template"), async (req, res) => {
  try {
    if (!req.file) return res.status(400).json({ error: "No file received." });
    if (!/\.pptx$/i.test(req.file.originalname)) return res.status(400).json({ error: "Please upload a .pptx file (in Canva: Share > Download > PowerPoint)." });
    const spec = await analyzeTpl(req.file.buffer);
    const id = crypto.randomUUID();
    fs.writeFileSync(path.join(TPL_DIR, id + ".pptx"), req.file.buffer);
    fs.writeFileSync(path.join(TPL_DIR, id + ".json"), JSON.stringify(spec));
    res.json({ id, name: req.file.originalname, swatches: spec.swatches, slides: spec.slides.length });
  } catch (e) { console.error(e); res.status(400).json({ error: "Could not use that template: " + e.message }); }
});

app.post("/api/pptx", async (req, res) => {
  try {
    const tid = req.body.theme?.id;
    let buf;
    if (tid && /^[0-9a-f-]{36}$/.test(tid)) {
      const f = path.join(TPL_DIR, tid + ".pptx");
      if (!fs.existsSync(f)) return res.status(410).json({ error: "That template is no longer on the server. Please upload it again." });
      buf = await buildFromTemplate(fs.readFileSync(f), JSON.parse(fs.readFileSync(path.join(TPL_DIR, tid + ".json"), "utf8")), req.body, { askFills, getPhotos });
    } else buf = await buildPptx(req.body);
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation");
    res.setHeader("Content-Disposition", 'attachment; filename="slides.pptx"');
    res.send(buf);
  } catch (e) { console.error(e); res.status(500).json({ error: e.message }); }
});

// friendly JSON errors for upload problems (instead of a crash page)
app.use((err, req, res, next) => {
  if (err instanceof multer.MulterError)
    return res.status(400).json({ error: err.code === "LIMIT_FILE_SIZE" ? "That file is too large. Documents can be up to 10 MB and templates up to 100 MB." : err.message });
  console.error(err);
  res.status(500).json({ error: "Server error: " + err.message });
});

app.listen(process.env.PORT || 3000, () => console.log("ChatPPT running at http://localhost:" + (process.env.PORT || 3000)));