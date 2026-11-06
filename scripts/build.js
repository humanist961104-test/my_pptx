/**
 * 덱 정의 JSON → .pptx
 *
 *   node <skill>/scripts/build.js deck.json [out.pptx] [--force]
 *
 * JSON 형식 (examples/sample_deck.json 참고):
 * {
 *   "title": "덱 제목",
 *   "output": "결과.pptx",              // 선택, 명령행 out이 우선
 *   "slides": [
 *     { "type": "cover",   ...brand.js cover() 인자 },
 *     { "type": "stats",   ..., "source": "출처" },
 *     ...
 *   ]
 * }
 * type: cover(A) | halfPanel(B) | textList(C) | stats(D) | section(E) | closing(F)  (A–F 문자도 허용)
 *
 * 빌드 전 검사 (ERROR면 중단, --force로 무시):
 *   ERROR  type 오류·필수 필드 누락, 발표자 노트 없음, 데이터 슬라이드 출처 없음,
 *          수치(%, 억, 조, 달러, 배, 만 등)가 들어간 슬라이드에 source 없음
 *   WARN   첫 장이 cover가 아님 / 마지막 장이 closing이 아님 / 5장 이상인데 내지 유형 3종 미만,
 *          항목 수 한도 초과
 */
const fs = require("fs");
const path = require("path");
const { createDeck } = require(path.join(__dirname, "brand.js"));

const ALIAS = { A: "cover", B: "halfPanel", C: "textList", D: "stats", E: "section", F: "closing" };
const REQUIRED = {
  cover: ["title"],
  halfPanel: ["title", "items"],
  textList: ["title", "items"],
  stats: ["title", "source"],
  section: ["title"],
  closing: [],
};
const LIMITS = { halfPanel: { items: 4 }, textList: { items: 4 }, stats: { points: 3, stats: 4 }, section: { items: 4 }, closing: { items: 3 }, cover: { satellites: 3 } };
// 통계로 보이는 수치: 숫자 뒤에 %, 배, 억, 조, 만, 달러, 원, 명, 건 등
const NUMERIC = /\d[\d,.]*\s*(%|퍼센트|배|억|조|만|천|달러|원|명|건|개국|\$)|\$\s*\d/;

function collectText(v, out = []) {
  if (typeof v === "string") out.push(v);
  else if (Array.isArray(v)) v.forEach((x) => collectText(x, out));
  else if (v && typeof v === "object") for (const [k, x] of Object.entries(v)) if (!["notes", "source", "icon", "image", "type"].includes(k)) collectText(x, out);
  return out;
}

function check(spec) {
  const errs = [], warns = [];
  if (!Array.isArray(spec.slides) || !spec.slides.length) errs.push("slides 배열이 비어 있음");
  const slides = spec.slides || [];
  slides.forEach((s, i) => {
    const tag = `슬라이드 ${i + 1}`;
    s.type = ALIAS[s.type] || s.type;
    if (!REQUIRED[s.type]) { errs.push(`${tag}: 알 수 없는 type "${s.type}"`); return; }
    for (const f of REQUIRED[s.type]) if (s[f] == null || s[f] === "" || (Array.isArray(s[f]) && !s[f].length)) errs.push(`${tag} (${s.type}): 필수 필드 "${f}" 없음`);
    if (!s.notes || !String(s.notes).trim()) errs.push(`${tag}: 발표자 노트(notes) 없음`);
    if (s.type !== "stats" && !s.source) {
      const hit = collectText(s).find((t) => NUMERIC.test(t));
      if (hit) errs.push(`${tag}: 수치 "${hit.slice(0, 25)}" 가 있는데 source 없음 — 출처를 넣거나 수치를 빼세요`);
    }
    for (const [f, n] of Object.entries(LIMITS[s.type] || {})) if (Array.isArray(s[f]) && s[f].length > n) warns.push(`${tag}: ${f} ${s[f].length}개 > 한도 ${n} (초과분은 잘림)`);
  });
  if (slides.length) {
    if (slides[0].type !== "cover") warns.push("첫 장이 cover(A)가 아님");
    if (slides.length > 1 && slides[slides.length - 1].type !== "closing") warns.push("마지막 장이 closing(F)이 아님");
    const inner = new Set(slides.slice(1, -1).map((s) => s.type));
    if (slides.length >= 5 && inner.size < 3) warns.push(`내지 유형 ${inner.size}종 — 5장 이상이면 3종 이상 섞기`);
  }
  return { errs, warns };
}

(async () => {
  const args = process.argv.slice(2).filter((a) => !a.startsWith("--"));
  const force = process.argv.includes("--force");
  if (!args[0]) { console.error("사용: node build.js deck.json [out.pptx] [--force]"); process.exit(2); }
  const specPath = path.resolve(args[0]);
  const spec = JSON.parse(fs.readFileSync(specPath, "utf8"));
  const baseDir = path.dirname(specPath);
  const out = path.resolve(args[1] || spec.output || specPath.replace(/\.json$/i, ".pptx"));

  const { errs, warns } = check(spec);
  warns.forEach((w) => console.warn("WARN ", w));
  errs.forEach((e) => console.error("ERROR", e));
  if (errs.length && !force) { console.error(`정의 파일 ERROR ${errs.length}건 — 빌드 중단 (--force로 무시)`); process.exit(1); }

  const deck = createDeck({ title: spec.title || "Presentation", author: spec.author });
  for (const s of spec.slides) {
    const { type, source, ...args2 } = s;
    if (args2.image) args2.image = path.resolve(baseDir, args2.image);
    if (type === "stats") args2.source = source;
    const slide = await deck[type](args2);
    // stats 외 슬라이드의 출처는 하단 캡션으로 표시
    if (type !== "stats" && source) {
      const onRed = ["cover", "section", "closing"].includes(type);
      deck.helpers.txt(slide, "출처: " + source, { x: 0.5, y: 5.25, w: 7.3, h: 0.25, fontSize: 10, color: onRed ? deck.C.WHITE : deck.C.GRAY_TEXT, wrap: false });
    }
  }
  const w = await deck.save(out);
  process.exit(w.some((x) => x.includes("넘침")) ? 3 : 0);
})().catch((e) => { console.error(e.stack || e); process.exit(1); });
