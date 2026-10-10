// Fills a user's .pptx template (e.g. from Canva): rewrites its text boxes and swaps its photos.
import JSZip from "jszip";
import sharp from "sharp";

const SP = /<p:sp>[\s\S]*?<\/p:sp>/g;
const attr = (s, re) => (s.match(re) || [])[1];
const unesc = s => s.replace(/&lt;/g, "<").replace(/&gt;/g, ">").replace(/&quot;/g, '"').replace(/&apos;/g, "'").replace(/&amp;/g, "&");
const esc = s => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
const rel = f => f.replace("slides/", "slides/_rels/") + ".rels";
const norm = t => { t = t.replace(/^\//, ""); return t.startsWith("ppt/") ? t : "ppt/" + t; };

// rough number of characters that fit in a box, and the font size that makes text fit
const cap = (w, h, sz) => Math.floor((w / (sz * 0.62 / 72)) * Math.max(1, Math.floor(h / (sz * 1.2 / 72))) * 0.9);
const fitSize = (text, w, h, sz) => {
  const longest = Math.max(...String(text).split(/\s+/).map(x => x.length), 1);
  let f = sz;
  while (f > sz * 0.35 && (cap(w, h, f) < text.length || longest * f * 0.65 / 72 > w * 0.9)) f -= 1;
  return f;
};

function parseShapes(xml) {
  const out = [];
  for (const m of xml.matchAll(SP)) {
    const raw = m[0];
    if (!raw.includes("<p:txBody>")) continue;
    const text = [...raw.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)]
      .map(p => [...p[0].matchAll(/<a:t(?: [^>]*)?>([\s\S]*?)<\/a:t>/g)].map(t => unesc(t[1])).join("")).join("\n").trim();
    const g = raw.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"/);
    if (!text || !g) continue;
    const sz = (+attr(raw, /<a:rPr\b[^>]*?\ssz="(\d+)"/) || +attr(raw, /\ssz="(\d+)"/) || 1800) / 100;
    out.push({ raw, id: attr(raw, /<p:cNvPr\b[^>]*?\bid="(\d+)"/), text, x: g[1] / 914400, y: g[2] / 914400, w: g[3] / 914400, h: g[4] / 914400, sz });
  }
  return out;
}
const lite = s => s && { id: s.id, text: s.text, w: s.w, h: s.h, sz: s.sz, x: s.x, y: s.y };

function analyzeSlide(xml) {
  const sh = parseShapes(xml).filter(s => !/^\d{1,2}$/.test(s.text));
  if (!sh.length) return { title: null, units: [] };
  const title = sh.reduce((a, b) => (b.sz > a.sz || (b.sz === a.sz && b.w * b.h > a.w * a.h)) ? b : a);
  const others = sh.filter(s => s !== title).sort((a, b) => a.y - b.y || a.x - b.x);
  const used = new Set(), units = [];
  for (const a of others) {
    if (used.has(a)) continue;
    used.add(a);
    const b = others.find(c => !used.has(c) && c.y > a.y && Math.abs(c.x - a.x) < 0.35 && c.y - (a.y + a.h) > -0.25 &&
      c.y - (a.y + a.h) < 1.0 && a.text.length <= 45 && a.h <= 1.3 && a.sz > c.sz + 0.5 && c.text.length > a.text.length);
    if (b) { used.add(b); units.push({ label: lite(a), body: lite(b) }); } else units.push({ body: lite(a) });
  }
  return { title: lite(title), units };
}

async function slideOrder(zip) {
  const pres = await zip.file("ppt/presentation.xml").async("string");
  const rels = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  const map = {};
  for (const r of rels.match(/<Relationship\b[^>]*>/g) || []) map[attr(r, /Id="([^"]+)"/)] = attr(r, /Target="([^"]+)"/);
  const order = [];
  for (const s of pres.match(/<p:sldId\b[^>]*>/g) || []) { const rid = attr(s, /r:id="([^"]+)"/); if (map[rid]) order.push({ rid, file: norm(map[rid]) }); }
  return order;
}

export async function analyzeTemplate(buf) {
  const zip = await JSZip.loadAsync(buf);
  if (!zip.file("ppt/presentation.xml")) throw new Error("this is not a PowerPoint file");
  const order = await slideOrder(zip);
  const slides = [], count = {};
  for (const o of order) {
    const xml = await zip.file(o.file).async("string");
    slides.push({ file: o.file, ...analyzeSlide(xml) });
    for (const m of xml.matchAll(/srgbClr val="([0-9A-Fa-f]{6})"/g)) count[m[1].toUpperCase()] = (count[m[1].toUpperCase()] || 0) + 1;
  }
  if (!slides.some((s, i) => i > 0 && s.units.length)) throw new Error("it has no slides with editable text boxes");
  const swatches = [];
  for (const c of Object.entries(count).sort((a, b) => b[1] - a[1]).map(e => e[0])) {
    const v = [0, 2, 4].map(i => parseInt(c.slice(i, i + 2), 16));
    if (v.reduce((a, b) => a + b, 0) < 700 && swatches.every(u => [0, 2, 4].reduce((s, i) => s + Math.abs(parseInt(u.slice(i, i + 2), 16) - v[i / 2]), 0) > 90)) swatches.push(c);
    if (swatches.length === 4) break;
  }
  return { slides, swatches };
}

export function slotsOf(s) {
  const out = [];
  if (s.title) out.push({ key: "title", role: "title", ...s.title });
  s.units.forEach((u, i) => {
    if (u.label) out.push({ key: `u${i}l`, role: "label", ...u.label });
    out.push({ key: `u${i}b`, role: "body", ...u.body });
  });
  return out.map(o => ({ ...o, maxChars: Math.min(o.role === "label" ? 40 : 400, Math.round(cap(o.w, o.h, o.sz) * 1.1)) }));
}

function setText(raw, text, w, h) {
  const body = raw.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/)[1];
  const head = body.match(/^[\s\S]*?(?=<a:p>)/)?.[0] ?? "";
  const p0 = body.match(/<a:p>[\s\S]*?<\/a:p>/)[0];
  const pPr = p0.match(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/)?.[0] ?? "";
  let rPr = p0.match(/<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/)?.[0] ?? '<a:rPr lang="en-US"/>';
  const sz = (+attr(rPr, /\ssz="(\d+)"/) || 1800) / 100;
  const fs = Math.round(fitSize(text, w, h, sz) * 100);
  rPr = /\ssz="/.test(rPr) ? rPr.replace(/\ssz="\d+"/, ` sz="${fs}"`) : rPr.replace(/^<a:rPr/, `<a:rPr sz="${fs}"`);
  const paras = String(text).split("\n").map(l => `<a:p>${pPr}<a:r>${rPr}<a:t>${esc(l)}</a:t></a:r></a:p>`).join("");
  return raw.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, () => `<p:txBody>${head}${paras}</p:txBody>`);
}

function fallbackFills(src, slide) {
  const f = { title: slide.title }, B = slide.bullets;
  src.units.forEach((u, i) => {
    const b = B[i] || "";
    if (u.label) f[`u${i}l`] = b.split(/[:\u2013\u2014-]/)[0].trim().split(" ").slice(0, 2).join(" ").slice(0, 16);
    f[`u${i}b`] = b;
  });
  return f;
}

export async function buildFromTemplate(buf, spec, deck, { askFills, getPhotos }) {
  const zip = await JSZip.loadAsync(buf);
  const S = spec.slides, order = await slideOrder(zip);
  const ctp = "[Content_Types].xml";
  let ct = await zip.file(ctp).async("string"), prel = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  let pres = await zip.file("ppt/presentation.xml").async("string");

  const last = S[S.length - 1];
  const closing = S.length > 2 && (!last.units.length || /thank/i.test(last.title?.text || "")) ? S.length - 1 : -1;
  const cands = S.map((s, i) => i).filter(i => i && i !== closing && S[i].units.length);
  const plan = [{ src: 0, kind: "title" }], uses = {};
  let prev = -1;
  for (const s of deck.slides) {
    let best = null;
    for (const i of cands) {
      const c = Math.abs(S[i].units.length - s.bullets.length) + (uses[i] || 0) * 1.5 + (i === prev ? 3 : 0) + Math.random() * 1.2;
      if (!best || c < best.c) best = { i, c };
    }
    plan.push({ src: best.i, kind: "content", slide: s });
    uses[best.i] = (uses[best.i] || 0) + 1; prev = best.i;
  }
  if (closing >= 0) plan.push({ src: closing, kind: "closing" });

  const content = plan.filter(p => p.kind === "content");
  let fills = null;
  try { fills = await askFills(content.map(p => ({ content: { title: p.slide.title, bullets: p.slide.bullets }, slots: slotsOf(S[p.src]).map(({ key, role, text, maxChars }) => ({ key, role, sample: text.slice(0, 60), maxChars })) }))); }
  catch (e) { console.error("fill request failed, using simple mapping:", e.message); }

  let num = Math.max(...Object.keys(zip.files).map(f => +(f.match(/^ppt\/slides\/slide(\d+)\.xml$/)?.[1] || 0)));
  const ids = []; let ci = 0;
  for (const [k, p] of plan.entries()) {
    const src = S[p.src], n = ++num, file = `ppt/slides/slide${n}.xml`;
    let xml = await zip.file(src.file).async("string"), rels = await zip.file(rel(src.file)).async("string");
    let f = {};
    if (p.kind === "title") { f = { title: deck.title }; src.units.forEach((u, i) => { f[`u${i}${u.label ? "l" : "b"}`] = i ? "" : (deck.subtitle || ""); if (u.label) f[`u${i}b`] = ""; }); }
    else if (p.kind === "content") { f = { ...fallbackFills(src, p.slide), ...(fills?.[ci] || {}) }; ci++; }
    if (p.kind !== "closing") {
      const shapes = Object.fromEntries(parseShapes(xml).map(s => [s.id, s]));
      for (const sl of slotsOf(src)) {
        const s = shapes[sl.id]; if (!s) continue;
        xml = xml.split(s.raw).join(setText(s.raw, String(f[sl.key] ?? ""), sl.w, sl.h));
      }
      const query = p.kind === "content" ? (p.slide.image_query || p.slide.title) : deck.title;
      const refs = (rels.match(/<Relationship\b[^>]*>/g) || []).filter(t => /Target="[^"]*\.jpe?g"/i.test(t) && xml.includes(`r:embed="${attr(t, /Id="([^"]+)"/)}"`));
      const photos = refs.length && getPhotos ? await getPhotos(query, refs.length).catch(() => []) : [];
      for (const [j, t] of refs.entries()) {
        if (!photos.length) break;
        const orig = zip.file(norm(attr(t, /Target="([^"]+)"/).replace("../", "")));
        if (!orig) continue;
        const ob = await orig.async("nodebuffer"); if (ob.length < 50000) continue;
        const m = await sharp(ob).metadata(), w = Math.min(m.width, 1600), h = Math.round(w * m.height / m.width);
        const nb = await sharp(photos[j % photos.length]).resize(w, h, { fit: "cover" }).jpeg({ quality: 82 }).toBuffer();
        zip.file(`ppt/media/tpl${n}_${j}.jpeg`, nb);
        rels = rels.replace(t, t.replace(/Target="[^"]*"/, `Target="../media/tpl${n}_${j}.jpeg"`));
      }
    }
    zip.file(file, xml); zip.file(rel(file), rels);
    ct = ct.replace("</Types>", `<Override PartName="/${file}" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>`);
    prel = prel.replace("</Relationships>", `<Relationship Id="rIdT${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${n}.xml"/></Relationships>`);
    ids.push(`<p:sldId id="${1000 + k}" r:id="rIdT${n}"/>`);
  }

  // drop the template's own slides (and their notes), then anything no longer used
  for (const o of order) {
    const r = zip.file(rel(o.file)) && await zip.file(rel(o.file)).async("string");
    for (const nt of (r || "").match(/Target="[^"]*notesSlide\d+\.xml"/g) || []) {
      const np = "ppt/notesSlides/" + nt.match(/notesSlide\d+\.xml/)[0];
      zip.remove(np); zip.remove(np.replace("notesSlides/", "notesSlides/_rels/") + ".rels");
      ct = ct.replace(new RegExp(`<Override[^>]*PartName="/${np}"[^>]*/>`), "");
    }
    zip.remove(o.file); zip.remove(rel(o.file));
    ct = ct.replace(new RegExp(`<Override[^>]*PartName="/${o.file}"[^>]*/>`), "");
    prel = prel.replace(new RegExp(`<Relationship\\b[^>]*Id="${o.rid}"[^>]*/>`), "");
  }
  pres = pres.replace(/<p:sldIdLst>[\s\S]*?<\/p:sldIdLst>/, `<p:sldIdLst>${ids.join("")}</p:sldIdLst>`)
    .replace(/<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}">[\s\S]*?<\/p:ext>/, "");
  zip.file(ctp, ct); zip.file("ppt/_rels/presentation.xml.rels", prel); zip.file("ppt/presentation.xml", pres);
  let refsAll = "";
  for (const name of Object.keys(zip.files)) if (name.endsWith(".rels")) refsAll += await zip.file(name).async("string");
  const keep = new Set([...refsAll.matchAll(/media\/([^"'/]+)"/g)].map(m => m[1]));
  for (const name of Object.keys(zip.files)) if (!zip.files[name].dir && name.startsWith("ppt/media/") && !keep.has(name.slice(10))) zip.remove(name);
  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE", compressionOptions: { level: 6 } });
}