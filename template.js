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
    const isBox = /txBox="(?:1|true)"/.test(raw);
    const text = [...raw.matchAll(/<a:p>[\s\S]*?<\/a:p>/g)]
      .map(p => [...p[0].matchAll(/<a:t(?: [^>]*)?>([\s\S]*?)<\/a:t>/g)].map(t => unesc(t[1])).join("")).join("\n").trim();
    const g = raw.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"/);
    if (!g || (!text && !isBox)) continue;
    const sz = (+attr(raw, /<a:rPr\b[^>]*?\ssz="(\d+)"/) || +attr(raw, /\ssz="(\d+)"/) || 1800) / 100;
    out.push({ raw, empty: !text, id: attr(raw, /<p:cNvPr\b[^>]*?\bid="(\d+)"/), text, x: g[1] / 914400, y: g[2] / 914400, w: g[3] / 914400, h: g[4] / 914400, sz });
  }
  return out;
}
const lite = s => s && { id: s.id, text: s.text, w: s.w, h: s.h, sz: s.sz, x: s.x, y: s.y };

const isChrome = (s, chrome, H) => (s.text && chrome.has(s.text.toLowerCase())) || (s.h < 0.06 * H && (s.y < 0.18 * H || s.y + s.h > 0.88 * H));
const styleOf = raw => { const b = raw?.match(/<p:txBody>[\s\S]*<\/p:txBody>/)?.[0] || ""; return { color: attr(b, /<a:srgbClr val="([0-9A-Fa-f]{6})"/), font: attr(b, /<a:latin typeface="([^"]+)"/) }; };

function analyzeSlide(shapes, chrome, H) {
  const sh = shapes.filter(s => !/^\d{1,2}$/.test(s.text) && !isChrome(s, chrome, H));
  const texted = sh.filter(s => s.text), pool = texted.length ? texted : sh;
  if (!pool.length) return { title: null, units: [] };
  const top = pool.reduce((a, b) => (b.y < a.y - 0.05 ? b : a));
  const big = pool.filter(s => s !== top && s.sz >= 2 * top.sz && Math.abs(s.y - top.y) < 2.5);
  const title = big.length ? big.reduce((a, b) => (b.sz > a.sz ? b : a)) : top;
  const overlap = (a, b) => Math.max(0, Math.min(a.x + a.w, b.x + b.w) - Math.max(a.x, b.x)) * Math.max(0, Math.min(a.y + a.h, b.y + b.h) - Math.max(a.y, b.y)) / Math.max(0.01, Math.min(a.w * a.h, b.w * b.h));
  const others = sh.filter(s => s !== title && overlap(s, title) < 0.4).sort((a, b) => a.y - b.y || a.x - b.x);
  const used = new Set(), units = [];
  for (const a of others) {
    if (used.has(a)) continue;
    used.add(a);
    const b = a.text && others.find(c => !used.has(c) && c.text && c.y > a.y && Math.abs(c.x - a.x) < 0.35 && c.y - (a.y + a.h) > -0.25 &&
      c.y - (a.y + a.h) < 1.0 && a.text.length <= 45 && a.h <= 1.3 && a.sz > c.sz + 0.5 && c.text.length > a.text.length);
    if (b) { used.add(b); units.push({ label: lite(a), body: lite(b) }); } else units.push({ body: lite(a) });
  }
  return { title: lite(title), units, _t: title.raw, _o: others[0]?.raw };
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
  const sz = (await zip.file("ppt/presentation.xml").async("string")).match(/<p:sldSz cx="(\d+)" cy="(\d+)"/);
  const W = sz ? sz[1] / 914400 : 13.33, H = sz ? sz[2] / 914400 : 7.5;
  const raws = [], count = {}, colors = {};
  for (const o of order) {
    const xml = await zip.file(o.file).async("string"), shapes = parseShapes(xml);
    raws.push({ file: o.file, shapes });
    for (const k of new Set(shapes.filter(s => s.text).map(s => s.text.toLowerCase()))) count[k] = (count[k] || 0) + 1;
    for (const m of xml.matchAll(/srgbClr val="([0-9A-Fa-f]{6})"/g)) colors[m[1].toUpperCase()] = (colors[m[1].toUpperCase()] || 0) + 1;
  }
  const chrome = new Set(Object.keys(count).filter(k => count[k] >= 3 && k.length <= 40));
  const slides = [], style = {};
  raws.forEach((r, i) => {
    const a = analyzeSlide(r.shapes, chrome, H);
    if (i && a._t && !style.title) style.title = styleOf(a._t);
    const body = a._o || r.shapes.find(s => s.text && !/^\d+$/.test(s.text) && s.raw !== a._t)?.raw;
    if (body && !style.body?.color) style.body = styleOf(body);
    slides.push({ file: r.file, title: a.title, units: a.units });
  });
  if (!slides.some((s, i) => i > 0 && s.title)) throw new Error("it has no slides with text boxes I can use");
  const swatches = [];
  for (const c of Object.keys(colors).sort((a, b) => colors[b] - colors[a])) {
    const v = [0, 2, 4].map(i => parseInt(c.slice(i, i + 2), 16));
    if (v.reduce((a, b) => a + b, 0) < 700 && swatches.every(u => [0, 2, 4].reduce((s, i) => s + Math.abs(parseInt(u.slice(i, i + 2), 16) - v[i / 2]), 0) > 90)) swatches.push(c);
    if (swatches.length === 4) break;
  }
  return { v: 3, W, H, style, chrome: [...chrome], slides, swatches };
}

export function slotsOf(s) {
  const out = [];
  if (s.title) {
    const below = s.units.flatMap(u => [u.label, u.body]).filter(b => b && b.y > s.title.y + 0.3 && b.x < s.title.x + s.title.w && b.x + b.w > s.title.x).map(b => b.y);
    const h = below.length ? Math.max(0.6, Math.min(s.title.h, Math.min(...below) - s.title.y - 0.05)) : s.title.h;
    out.push({ key: "title", role: "title", ...s.title, h });
  }
  s.units.forEach((u, i) => {
    if (u.label) out.push({ key: `u${i}l`, role: "label", ...u.label });
    out.push({ key: `u${i}b`, role: "body", ...u.body });
  });
  return out.map(o => ({ ...o, maxChars: Math.min(o.role === "label" ? 40 : 400, Math.round(cap(o.w, o.h, o.sz) * 1.1)) }));
}

function setText(raw, text, w, h, fsPt) {
  const body = raw.match(/<p:txBody>([\s\S]*?)<\/p:txBody>/)[1];
  const head = body.match(/^[\s\S]*?(?=<a:p[>/])/)?.[0] ?? "";
  const p0 = body.match(/<a:p>[\s\S]*?<\/a:p>|<a:p\/>/)?.[0] ?? "";
  const pPr = p0.match(/<a:pPr\b[^>]*\/>|<a:pPr\b[^>]*>[\s\S]*?<\/a:pPr>/)?.[0] ?? "";
  let rPr = p0.match(/<a:rPr\b[^>]*\/>|<a:rPr\b[^>]*>[\s\S]*?<\/a:rPr>/)?.[0]
    ?? p0.match(/<a:endParaRPr\b[^>]*\/>|<a:endParaRPr\b[^>]*>[\s\S]*?<\/a:endParaRPr>/)?.[0]?.replace(/endParaRPr/g, "rPr") ?? '<a:rPr lang="en-US"/>';
  const sz = (+attr(rPr, /\ssz="(\d+)"/) || +attr(raw, /\ssz="(\d+)"/) || 1800) / 100;
  let f = fsPt ?? fitSize(text, w, h, sz);
  if (fsPt == null && f === sz && sz < 48) while (f < sz * 1.4 && cap(w, h, f + 1) >= text.length * 2.5) f += 1; // short text in a big box: make it larger
  const fs = Math.round(f * 100);
  rPr = /\ssz="/.test(rPr) ? rPr.replace(/\ssz="\d+"/, ` sz="${fs}"`) : rPr.replace(/^<a:rPr/, `<a:rPr sz="${fs}"`);
  const pPrFit = pPr.replace(/algn="just"/, 'algn="l"').replace(/(<a:lnSpc>\s*<a:spcPts val=")(\d+)(")/, (m, a, v, c) => a + Math.round(v * f / sz) + c); // fixed line spacing must shrink with the font
  const paras = String(text).split("\n").map(l => `<a:p>${pPrFit}<a:r>${rPr}<a:t>${esc(l)}</a:t></a:r></a:p>`).join("");
  return raw.replace(/<p:txBody>[\s\S]*?<\/p:txBody>/, () => `<p:txBody>${head}${paras}</p:txBody>`);
}

// ---------- creating text boxes / pictures where the template has none ----------
const E9 = v => Math.round(v * 914400);
function topLevelBoxes(xml) {
  const out = [], re = /<(\/?)p:(grpSp|sp|pic|cxnSp|graphicFrame)>/g, xf = /<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"\/>/g;
  let m, depth = 0;
  while ((m = re.exec(xml))) {
    if (m[2] === "grpSp") { if (m[1]) { depth--; continue; } if (depth++ > 0) continue; }
    else if (m[1] || depth > 0) continue;
    xf.lastIndex = m.index;
    const g = xf.exec(xml);
    if (g) out.push({ x: g[1] / 914400, y: g[2] / 914400, w: g[3] / 914400, h: g[4] / 914400 });
  }
  return out;
}

// biggest empty rectangle on the slide (ignores full-slide backgrounds and empty text boxes)
function freeRect(xml, W, H) {
  const empties = new Set(parseShapes(xml).filter(s => s.empty).map(s => `${s.x.toFixed(2)},${s.y.toFixed(2)}`));
  const obs = topLevelBoxes(xml).filter(b => b.w * b.h < 0.45 * W * H && !empties.has(`${b.x.toFixed(2)},${b.y.toFixed(2)}`));
  const step = 0.25, cols = Math.floor(W / step), rows = Math.floor(H / step), M = 0.7;
  const blocked = (r, c) => { const x = c * step, y = r * step;
    return x < M || x + step > W - M || y < 0.8 || y + step > H - 0.7 || obs.some(o => x < o.x + o.w + 0.15 && x + step > o.x - 0.15 && y < o.y + o.h + 0.15 && y + step > o.y - 0.15); };
  const hist = new Array(cols).fill(0);
  let best = null;
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) hist[c] = blocked(r, c) ? 0 : hist[c] + 1;
    const st = [];
    for (let c = 0; c <= cols; c++) {
      const h = c < cols ? hist[c] : 0; let start = c;
      while (st.length && st[st.length - 1][1] >= h) {
        const [s, hh] = st.pop(), area = hh * (c - s);
        if (!best || area > best.area) best = { area, x: s * step, y: (r - hh + 1) * step, w: (c - s) * step, h: hh * step };
        start = s;
      }
      st.push([start, h]);
    }
  }
  return best && best.w >= 3.5 && best.h >= 2 ? best : null;
}

function addText(xml, r, paras, o) {
  const text = paras.join(" ") + " ".repeat(o.bullets ? paras.length * 25 : 0);
  const sz = Math.max(14, Math.round(fitSize(text, r.w - (o.bullets ? 0.4 : 0), r.h, o.sz)));
  const rPr = `<a:rPr lang="en-US" sz="${sz * 100}"${o.bold ? ' b="1"' : ""} dirty="0"><a:solidFill><a:srgbClr val="${o.color || "222222"}"/></a:solidFill>${o.font ? `<a:latin typeface="${esc(o.font)}"/>` : ""}</a:rPr>`;
  const pPr = o.bullets ? '<a:pPr marL="342900" indent="-342900"><a:spcAft><a:spcPts val="1400"/></a:spcAft><a:buFont typeface="Arial"/><a:buChar char="&#8226;"/></a:pPr>' : "";
  const ps = paras.map(x => `<a:p>${pPr}<a:r>${rPr}<a:t>${esc(x)}</a:t></a:r></a:p>`).join("");
  const sp = `<p:sp><p:nvSpPr><p:cNvPr id="${o.id}" name="ChatPPT text ${o.id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${E9(r.x)}" y="${E9(r.y)}"/><a:ext cx="${E9(r.w)}" cy="${E9(r.h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:normAutofit/></a:bodyPr><a:lstStyle/>${ps}</p:txBody></p:sp>`;
  return xml.replace("</p:spTree>", () => sp + "</p:spTree>");
}

function addPic(xml, r, rid, id) {
  const pic = `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="ChatPPT photo ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${E9(r.x)}" y="${E9(r.y)}"/><a:ext cx="${E9(r.w)}" cy="${E9(r.h)}"/></a:xfrm><a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  return xml.replace("</p:spTree>", () => pic + "</p:spTree>");
}


const cleanText = s => String(s ?? "").replace(/^[\s\u2022*-]+/, "").replace(/\s*[;,]+\s*$/, "").trim();

// font size for a text slot: titles only shrink; other text may grow a little in big boxes
function bestSize(text, w, h, sz, role, W = 13.33) {
  if (role === "title") return fitSize(text, w, h, sz);
  const longest = Math.max(...text.split(/\s+/).map(x => x.length), 1);
  for (let f = Math.min(Math.max(sz * 1.4, 1.3 * W), sz * 2.5); f > sz * 0.35; f -= 0.5) if (cap(w, h, f) * 0.8 >= text.length && longest * f * 0.65 / 72 <= w * 0.9) return f;
  return sz * 0.35;
}

// page numbers, brand name, menu items, contact details and leftover sample text from the template
function dropIconsNear(xml, s) {
  for (const m of xml.matchAll(SP)) {
    const r = m[0], g = r.match(/<a:off x="(-?\d+)" y="(-?\d+)"\/>\s*<a:ext cx="(\d+)" cy="(\d+)"/);
    if (!g || /<a:t(?: [^>]*)?>[^<]+<\/a:t>/.test(r)) continue;
    const x = g[1] / 914400, y = g[2] / 914400, w = g[3] / 914400, h = g[4] / 914400;
    if (w * h < 1.5 && x + w <= s.x + 0.2 && x >= s.x - 1.3 && y < s.y + s.h && y + h > s.y) xml = xml.split(r).join("");
  }
  return xml;
}

function tidy(xml, k, kind, used, deck, spec, H) {
  const chrome = new Set(spec.chrome || []), all = parseShapes(xml);
  const topBand = all.filter(s => s.text && s.h < 0.06 * H && s.y < 0.18 * H).sort((a, b) => a.x - b.x);
  for (const s of all) {
    if (!s.text || used.has(s.id) || /^\d{1,2}$/.test(s.text)) continue;
    let t = null;
    if (/^page\s*\d+$/i.test(s.text)) t = `Page ${String(k + 1).padStart(2, "0")}`;
    else if (s === topBand[0] && chrome.has(s.text.toLowerCase())) t = s.text === s.text.toUpperCase() ? deck.title.toUpperCase() : deck.title;
    else if (chrome.has(s.text.toLowerCase()) || /@|www\.|\.com\b|\+?\d[\d\s().-]{7,}|\b(street|st\.|avenue|ave\.)/i.test(s.text)) t = "";
    else if (kind !== "closing") t = "";
    if (t === "" && kind === "closing" && !chrome.has(s.text.toLowerCase())) xml = dropIconsNear(xml, s);
    if (t !== null) xml = xml.split(s.raw).join(setText(s.raw, t, s.w, s.h, t ? fitSize(t, s.w, s.h, s.sz) : s.sz));
  }
  return xml;
}


function fallbackFills(src, slide) {
  const f = { title: slide.title }, B = slide.bullets;
  src.units.forEach((u, i) => {
    const b = B[i] || "";
    if (u.label) f[`u${i}l`] = b.split(/[:\u2013\u2014-]/)[0].trim().split(" ").slice(0, 2).join(" ").slice(0, 16);
    f[`u${i}b`] = cleanText(b);
  });
  return f;
}

export async function buildFromTemplate(buf, spec, deck, { askFills, getPhotos }) {
  const zip = await JSZip.loadAsync(buf);
  const S = spec.slides, order = await slideOrder(zip);
  const ctp = "[Content_Types].xml";
  let ct = await zip.file(ctp).async("string"), prel = await zip.file("ppt/_rels/presentation.xml.rels").async("string");
  let pres = await zip.file("ppt/presentation.xml").async("string");

  const W = spec.W || 13.33, H = spec.H || 7.5, st = spec.style || {};
  const last = S[S.length - 1];
  const closing = S.length > 2 && (/thank|questions|the end/i.test(last.title?.text || "") || (!last.title && !last.units.length)) ? S.length - 1 : -1;
  const cands = S.map((s, i) => i).filter(i => i && i !== closing && S[i].title);
  if (!cands.length) throw new Error("this template has no slide with a title text box");
  const plan = [{ src: 0, kind: "title" }], uses = {};
  let prev = -1;
  for (const s of deck.slides) {
    let best = null;
    for (const i of cands) {
      const U = S[i].units.length, avg = s.bullets.reduce((a, b) => a + b.length, 0) / Math.max(1, s.bullets.length);
      const small = S[i].units.filter(u => cap(u.body.w, u.body.h, u.body.sz * 1.4) < 0.7 * avg).length;
      const c = (U ? Math.abs(U - s.bullets.length) + 2 * small / U : 2.5) + (uses[i] || 0) * 1.5 + (i === prev ? 3 : 0) + Math.random() * 1.2;
      if (!best || c < best.c) best = { i, c };
    }
    plan.push({ src: best.i, kind: "content", slide: s });
    uses[best.i] = (uses[best.i] || 0) + 1; prev = best.i;
  }
  if (closing >= 0) plan.push({ src: closing, kind: "closing" });

  const content = plan.filter(p => p.kind === "content" && S[p.src].units.length);
  let fills = null;
  try { if (content.length) fills = await askFills(content.map(p => ({ content: { title: p.slide.title, bullets: p.slide.bullets, notes: p.slide.notes }, slots: slotsOf(S[p.src]).map(({ key, role, text, maxChars }) => ({ key, role, sample: text.slice(0, 60), maxChars })) }))); }
  catch (e) { console.error("fill request failed, using simple mapping:", e.message); }

  let num = Math.max(...Object.keys(zip.files).map(f => +(f.match(/^ppt\/slides\/slide(\d+)\.xml$/)?.[1] || 0)));
  const ids = []; let ci = 0;
  for (const [k, p] of plan.entries()) {
    const src = S[p.src], n = ++num, file = `ppt/slides/slide${n}.xml`;
    let xml = await zip.file(src.file).async("string"), rels = await zip.file(rel(src.file)).async("string");
    let f = {};
    if (p.kind === "title") { f = { title: deck.title }; src.units.forEach((u, i) => { f[`u${i}${u.label ? "l" : "b"}`] = i ? "" : (deck.subtitle || ""); if (u.label) f[`u${i}b`] = ""; }); }
    else if (p.kind === "content") { const has = src.units.length > 0; f = { ...fallbackFills(src, p.slide), ...(has ? fills?.[ci] || {} : {}) }; if (has) ci++; }
    if (p.kind !== "closing") {
      const shapes = Object.fromEntries(parseShapes(xml).map(s => [s.id, s]));
      const items = slotsOf(src).map(sl => ({ sl, s: shapes[sl.id], text: cleanText(f[sl.key]) })).filter(x => x.s);
      items.forEach(x => { x.fs = x.text ? bestSize(x.text, x.sl.w, x.sl.h, x.sl.sz, x.sl.role, W) : x.sl.sz; });
      const titleFs = items.find(x => x.sl.role === "title" && x.text)?.fs;
      for (const role of ["body", "label"]) {
        const g = items.filter(x => x.sl.role === role && x.text);
        if (g.length > 1) { const u = Math.max(Math.min(...g.map(x => x.fs)), 0.6 * Math.max(...g.map(x => x.fs))); g.forEach(x => { x.fs = Math.min(x.fs, u); }); }
        if (titleFs) g.forEach(x => { x.fs = Math.min(x.fs, titleFs * 0.75); });
      }
      for (const x of items) xml = xml.split(x.s.raw).join(setText(x.s.raw, x.text, x.sl.w, x.sl.h, x.fs));
      xml = tidy(xml, k, p.kind, new Set(items.map(x => x.sl.id)), deck, spec, H);
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

      // add text boxes / a picture when the template slide has none
      const needTitle = !src.title, needBody = p.kind === "content" && !src.units.length, needSub = p.kind === "title" && !src.units.length && !!deck.subtitle;
      const extra = !refs.length && getPhotos ? await getPhotos(query, 1).catch(() => []) : [];
      if (needTitle || needBody || needSub || extra.length) {
        const R = freeRect(xml, W, H);
        if (R) {
          let title, sub, body, pic, rest = R;
          if (needTitle) { const h = p.kind === "title" ? R.h * 0.5 : Math.min(2, R.h * 0.3); title = { ...R, h }; rest = { ...R, y: R.y + h + 0.1, h: R.h - h - 0.1 }; }
          if (p.kind === "title" && (needTitle || needSub)) {
            sub = { ...rest, h: Math.min(1.5, rest.h * 0.45) };
            rest = rest.h - sub.h - 0.1 >= 2.2 ? { ...rest, y: rest.y + sub.h + 0.1, h: rest.h - sub.h - 0.1 } : null;
          }
          if (rest) {
            const wantPic = extra.length && rest.w >= 3 && rest.h >= 2.2 && rest.w * rest.h >= 8 && rest.w / rest.h <= 3;
            if (needBody && wantPic) {
              if (rest.w >= 9) { const bw = rest.w * 0.55; body = { ...rest, w: bw }; pic = { x: rest.x + bw + 0.3, y: rest.y, w: rest.w - bw - 0.3, h: rest.h }; }
              else if (rest.h >= 6) { const ph = rest.h * 0.4; pic = { ...rest, h: ph }; body = { ...rest, y: rest.y + ph + 0.2, h: rest.h - ph - 0.2 }; }
              else body = rest;
            } else if (needBody) body = rest;
            else if (wantPic) pic = rest;
          }
          let maxId = Math.max(1, ...[...xml.matchAll(/<p:cNvPr\b[^>]*?\bid="(\d+)"/g)].map(m => +m[1]));
          const tc = st.title?.color || st.body?.color, bc = st.body?.color || st.title?.color;
          if (title) xml = addText(xml, title, [p.kind === "title" ? deck.title : p.slide.title], { sz: p.kind === "title" ? 54 : 40, bold: true, color: tc, font: st.title?.font || st.body?.font, id: ++maxId });
          if (sub && deck.subtitle) xml = addText(xml, sub, [deck.subtitle], { sz: 24, color: bc, font: st.body?.font || st.title?.font, id: ++maxId });
          if (body) xml = addText(xml, body, p.slide.bullets, { sz: 28, bullets: true, color: bc, font: st.body?.font || st.title?.font, id: ++maxId });
          if (pic) {
            const pw = Math.min(1600, Math.round(pic.w * 150)), ph = Math.round(pw * pic.h / pic.w);
            zip.file(`ppt/media/tpl${n}_new.jpeg`, await sharp(extra[0]).resize(pw, ph, { fit: "cover" }).jpeg({ quality: 82 }).toBuffer());
            rels = rels.replace("</Relationships>", `<Relationship Id="rIdChatPPT${n}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/tpl${n}_new.jpeg"/></Relationships>`);
            if (!/Extension="jpe?g"/i.test(ct)) ct = ct.replace(/<Types[^>]*>/, m => m + '<Default Extension="jpeg" ContentType="image/jpeg"/>');
            xml = addPic(xml, pic, `rIdChatPPT${n}`, ++maxId);
          }
        }
      }
    }
    if (p.kind === "closing") xml = tidy(xml, k, "closing", new Set(), deck, spec, H);
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