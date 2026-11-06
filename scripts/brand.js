/**
 * my_pptx brand library — Red Pitch Deck 레이아웃 6유형(A–F) + 장식 모티프 + 한글 줄바꿈.
 *
 * 의존 패키지(pptxgenjs, react, react-dom, react-icons, sharp)는 이 파일 위치가 아니라
 * **현재 작업 폴더(process.cwd())의 node_modules** 에서 찾는다. 그래서 스킬 폴더에는
 * node_modules를 둘 필요가 없다.
 *
 * 사용 예:
 *   const { createDeck } = require("<skill>/scripts/brand.js");
 *   const deck = createDeck({ title: "AI Agent의 미래" });
 *   await deck.cover({ brand: "AI Report", title: "AI Agent의\n미래", subtitle: "...", icon: "FaRobot", notes: "..." });
 *   await deck.save("out.pptx");
 */
const path = require("path");
const { createRequire } = require("module");
const req = createRequire(path.join(process.cwd(), "noop.js"));

const pptxgen = req("pptxgenjs");
const React = req("react");
const ReactDOMServer = req("react-dom/server");
const sharp = req("sharp");
const fa = req("react-icons/fa");

// ---------------------------------------------------------------- tokens (theme.json)
// 색·글꼴은 스킬 루트의 theme.json에서 읽는다. 다른 브랜드는 theme.json만 바꾸거나
// 환경변수 MY_PPTX_THEME=<경로> 로 다른 테마 파일을 지정한다.
const fs = require("fs");
const THEME_PATH = process.env.MY_PPTX_THEME || path.join(__dirname, "..", "theme.json");
const THEME = JSON.parse(fs.readFileSync(THEME_PATH, "utf8"));
const C = THEME.colors; // 키 이름(RED, RED_DEEP, ...)은 레이아웃이 역할로 사용하므로 유지한다
for (const k of ["RED", "RED_DEEP", "RED_TINT", "WHITE", "BLACK", "GRAY_TEXT", "GRAY_LIGHT", "GRAY_LINE"]) {
  if (!/^[0-9A-Fa-f]{6}$/.test(C[k] || "")) throw new Error(`theme.json colors.${k} 가 6자리 hex가 아님: ${C[k]}`);
}
const FONT_KO = THEME.fonts.ko, FONT_EN = THEME.fonts.en;
const W = 10, H = 5.625, M = 0.5;

// ---------------------------------------------------------------- 한글 줄바꿈
const isWide = (ch) => /[ᄀ-ᇿ　-〿㄰-㆏가-힣一-鿿＀-￯]/.test(ch);
function textWidthEm(s) {
  let w = 0;
  for (const ch of s) w += ch === " " ? 0.3 : isWide(ch) ? 1.0 : /[A-Z0-9%]/.test(ch) ? 0.65 : 0.55;
  return w;
}
function greedy(para, maxEm) {
  const lines = [];
  let cur = "";
  for (const word of para.split(" ")) {
    const cand = cur ? cur + " " + word : word;
    if (textWidthEm(cand) <= maxEm) { cur = cand; continue; }
    if (cur) lines.push(cur);
    // 단어 하나가 폭보다 길어도 자르지 않는다 (단어 중간 끊김 방지, 필요하면 PowerPoint가 처리)
    cur = word;
  }
  lines.push(cur);
  return lines;
}
/**
 * 상자 폭(인치)·글자 크기(pt)에 맞춰 단어 경계에서 \n을 넣는다. 기존 \n은 유지.
 * 줄 수를 늘리지 않는 범위에서 폭을 줄여 줄 길이를 고르게 맞춘다 (마지막 줄 한 단어 방지).
 */
function wrapKo(text, widthIn, fontPt) {
  const maxEm = (widthIn * 72) / fontPt * 0.95;
  return String(text).split("\n").map((para) => {
    let best = greedy(para, maxEm);
    if (best.length > 1) {
      for (let em = maxEm - 0.25; em > maxEm * 0.5; em -= 0.25) {
        const cand = greedy(para, em);
        if (cand.length !== best.length) break;
        best = cand;
      }
    }
    return best.join("\n");
  }).join("\n");
}

// ---------------------------------------------------------------- 아이콘
const iconCache = new Map();
async function iconData(name, color, size = 256) {
  const key = name + color;
  if (iconCache.has(key)) return iconCache.get(key);
  const Comp = fa[name];
  if (!Comp) throw new Error(`알 수 없는 아이콘: ${name} (react-icons/fa 이름을 쓰세요, 예: FaRobot)`);
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  const data = "image/png;base64," + buf.toString("base64");
  iconCache.set(key, data);
  return data;
}
async function waveData() {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 1000 180"><path d="M0,95 C180,20 330,165 560,110 C760,62 880,40 1000,18 L1000,180 L0,180 Z" fill="#FFFFFF"/></svg>`;
  const buf = await sharp(Buffer.from(svg), { density: 300 }).resize(2000, 360).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

// ---------------------------------------------------------------- deck
function createDeck({ title = "Presentation", author } = {}) {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = title;
  if (author) pres.author = author;
  const warnings = [];
  let slideNo = 0;

  /** 텍스트 추가: 한글이면 자동 줄바꿈, 넘침 추정 시 경고 */
  function txt(slide, text, o) {
    const size = o.fontSize || 12;
    const ko = /[가-힣]/.test(text);
    const body = ko && o.wrap !== false ? wrapKo(text, o.w - 0.02, size) : text;
    const lines = body.split("\n").length;
    const need = (lines * size * 1.2) / 72;
    if (o.h && need > o.h + 0.08) {
      warnings.push(`슬라이드 ${slideNo}: 텍스트 넘침 가능 (${lines}줄, 필요 ${need.toFixed(2)}" > 상자 ${o.h}") — "${String(text).slice(0, 20)}…"`);
    }
    const { wrap, ...rest } = o;
    slide.addText(body, { fontFace: ko ? FONT_KO : FONT_EN, margin: 0, color: C.BLACK, valign: "top", ...rest });
  }
  function rect(slide, x, y, w, h, color) {
    slide.addShape(pres.shapes.RECTANGLE, { x, y, w, h, fill: { color }, line: { color } });
  }
  function oval(slide, x, y, d, color) {
    slide.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color }, line: { color } });
  }
  async function badge(slide, x, y, d, circle, iconName) {
    oval(slide, x, y, d, circle);
    const p = d * 0.27;
    const iconColor = circle === C.WHITE ? C.RED : C.WHITE;
    slide.addImage({ data: await iconData(iconName, iconColor), x: x + p, y: y + p, w: d - 2 * p, h: d - 2 * p });
  }
  function xMark(slide, x, y, color = C.RED) {
    slide.addText("×", { x, y, w: 0.2, h: 0.2, fontFace: FONT_EN, fontSize: 12, bold: true, color, margin: 0, align: "center", valign: "middle" });
  }
  function cornerBlock(slide) {
    rect(slide, 8.6, 5.05, 1.4, 0.575, C.RED);
    rect(slide, 8.0, 5.45, 0.5, 0.12, C.BLACK);
  }
  function dotLine(slide, x, y, w) {
    slide.addShape(pres.shapes.LINE, { x, y, w, h: 0, line: { color: C.GRAY_LINE, width: 1, dashType: "sysDot" } });
  }
  function newSlide(bg, notes) {
    slideNo++;
    const s = pres.addSlide();
    s.background = { color: bg };
    if (notes) s.addNotes(notes);
    else warnings.push(`슬라이드 ${slideNo}: 발표자 노트 없음`);
    return s;
  }
  function title(slide, text, color, o = {}) {
    txt(slide, text, { x: M, y: 0.45, w: W - 2 * M, h: 0.65, fontSize: 30, bold: true, color, valign: "middle", ...o });
  }
  /** 아이콘 원 + 굵은 제목 + 설명 목록 (B·C형 공용) */
  async function iconList(slide, items, x, y0, w, step, onRed = false) {
    for (let i = 0; i < items.length; i++) {
      const it = items[i], y = y0 + i * step;
      await badge(slide, x, y + 0.05, 0.6, onRed ? C.WHITE : C.RED, it.icon || "FaCheck");
      txt(slide, it.title, { x: x + 0.8, y: y + 0.05, w: w - 0.8, h: 0.3, fontSize: 15, bold: true, color: onRed ? C.WHITE : C.BLACK, wrap: false });
      if (it.desc) txt(slide, it.desc, { x: x + 0.8, y: y + 0.37, w: w - 0.8, h: step - 0.42, fontSize: 12, color: onRed ? C.WHITE : C.GRAY_TEXT });
    }
  }
  const limit = (arr, n, name) => {
    if (arr.length > n) warnings.push(`슬라이드 ${slideNo}: ${name} 항목 ${arr.length}개 → 앞의 ${n}개만 사용`);
    return arr.slice(0, n);
  };

  return {
    pres, C, FONT_KO, FONT_EN, warnings,
    helpers: { txt, rect, oval, badge, xMark, cornerBlock, dotLine, iconData, wrapKo },

    /** A. 표지 — 전면 빨강 + 흰 웨이브 + 아이콘 클러스터 */
    async cover({ brand = "", title: t, subtitle = "", icon = "FaStar", satellites = [], notes }) {
      const s = newSlide(C.RED, notes);
      if (brand) txt(s, brand, { x: M, y: 0.35, w: 4, h: 0.4, fontSize: 14, bold: true, color: C.WHITE, wrap: false });
      txt(s, t, { x: 0.8, y: 1.15, w: 5.6, h: 2.1, fontSize: 54, bold: true, color: C.WHITE });
      if (subtitle) txt(s, subtitle, { x: 0.8, y: 3.3, w: 5.4, h: 0.45, fontSize: 18, color: C.WHITE });
      s.addImage({ data: await waveData(), x: 0, y: 3.825, w: W, h: 1.8 });
      const cx = 7.9, cy = 2.2;
      oval(s, cx - 1.6, cy - 1.6, 3.2, C.RED_DEEP);
      oval(s, cx - 1.25, cy - 1.25, 2.5, C.WHITE);
      s.addImage({ data: await iconData(icon, C.RED), x: cx - 0.7, y: cy - 0.7, w: 1.4, h: 1.4 });
      const pos = [[cx - 2.05, cy - 1.45], [cx + 1.15, cy - 1.5], [cx + 1.1, cy + 1.2]];
      const sats = limit(satellites, 3, "satellites");
      for (let i = 0; i < sats.length; i++) await badge(s, pos[i][0], pos[i][1], 0.65, C.WHITE, sats[i]);
      return s;
    },

    /** B. 하프 패널 — 한쪽 절반 빨강(또는 사진) + 반대쪽 제목·아이콘 목록 */
    async halfPanel({ side = "left", image, panelIcon = "FaStar", panelText = "", title: t, items = [], footnote, notes }) {
      const s = newSlide(C.WHITE, notes);
      const px = side === "left" ? 0 : W - 4.5;
      const cx = side === "left" ? 5.0 : M;
      if (image) s.addImage({ path: image, x: px, y: 0, w: 4.5, h: H, sizing: { type: "cover", w: 4.5, h: H } });
      else {
        rect(s, px, 0, 4.5, H, C.RED);
        oval(s, px + 1.05, 1.1, 2.4, C.WHITE);
        s.addImage({ data: await iconData(panelIcon, C.RED), x: px + 1.55, y: 1.6, w: 1.4, h: 1.4 });
        if (panelText) txt(s, panelText, { x: px + 0.5, y: 3.85, w: 3.5, h: 0.9, fontSize: 18, bold: true, color: C.WHITE, align: "center" });
        xMark(s, px + 0.35, 0.35, C.WHITE);
      }
      title(s, t, C.RED, { x: cx, w: 4.5 });
      await iconList(s, limit(items, 4, "items"), cx, 1.4, 4.5, 0.85);
      if (footnote) txt(s, footnote, { x: cx, y: 4.85, w: 3.4, h: 0.3, fontSize: 11, color: C.GRAY_TEXT });
      if (side === "left") cornerBlock(s); else xMark(s, 9.45, 0.35, C.BLACK);
      return s;
    },

    /** C. 본문 목록 — 흰 바탕, 아이콘 목록 + (선택) 오른쪽 RED_TINT 카드 */
    async textList({ title: t, items = [], card, notes }) {
      const s = newSlide(C.WHITE, notes);
      title(s, t, C.RED);
      const listW = card ? 5.5 : 9.0;
      const its = limit(items, 4, "items");
      await iconList(s, its, M, 1.4, listW, 0.85);
      for (let i = 0; i < its.length - 1; i++) dotLine(s, M, 1.4 + i * 0.85 + 0.8, listW);
      if (card) {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 6.5, y: 1.4, w: 3.0, h: 3.25, rectRadius: 0.1, fill: { color: C.RED_TINT }, line: { color: C.RED_TINT } });
        await badge(s, 6.85, 1.75, 0.8, C.RED, card.icon || "FaBullseye");
        txt(s, card.title, { x: 6.85, y: 2.8, w: 2.4, h: 0.85, fontSize: 18, bold: true });
        if (card.desc) txt(s, card.desc, { x: 6.85, y: 3.75, w: 2.4, h: 0.8, fontSize: 12, color: C.GRAY_TEXT });
      }
      xMark(s, 9.45, 0.35, C.BLACK);
      cornerBlock(s);
      return s;
    },

    /** D. 데이터 — 차트(bar/line/doughnut/pie) + (선택) 요점 + 하단 빨강 통계 바 */
    async stats({ title: t, chart, points = [], stats = [], source, notes }) {
      const s = newSlide(C.WHITE, notes);
      title(s, t, C.RED);
      if (!source) warnings.push(`슬라이드 ${slideNo}: 데이터 슬라이드에 출처(source) 없음`);
      const chartW = points.length ? 5.4 : 9.0;
      if (chart) {
        const type = { bar: pres.charts.BAR, line: pres.charts.LINE, doughnut: pres.charts.DOUGHNUT, pie: pres.charts.PIE }[chart.type || "bar"];
        const round = ["doughnut", "pie"].includes(chart.type);
        const series = chart.series || [{ name: chart.name || "값", labels: chart.labels, values: chart.values }];
        s.addChart(type, series, {
          x: M, y: 1.2, w: chartW, h: 2.65,
          chartColors: round ? [C.RED, C.RED_DEEP, C.BLACK, C.GRAY_LINE, C.GRAY_TEXT] : [C.RED, C.BLACK, C.GRAY_TEXT, C.GRAY_LINE],
          catAxisLabelFontFace: FONT_KO, valAxisLabelFontFace: FONT_EN, catAxisLabelFontSize: 10, valAxisLabelFontSize: 10,
          catAxisLabelColor: C.GRAY_TEXT, valAxisLabelColor: C.GRAY_TEXT,
          valGridLine: { color: C.GRAY_LIGHT, size: 0.5 }, catGridLine: { style: "none" },
          ...(round ? {} : { valAxisMinVal: chart.min ?? 0 }),
          ...(chart.format ? { dataLabelFormatCode: chart.format } : {}),
          showValue: !!chart.showValue, dataLabelFontSize: 10, dataLabelColor: round ? C.WHITE : C.BLACK,
          showLegend: series.length > 1 || round, legendPos: round ? "r" : "b", legendFontFace: FONT_KO, legendFontSize: 10,
          holeSize: 55, barGapWidthPct: 60,
        });
      }
      if (points.length) {
        const ps = limit(points, 3, "points");
        for (let i = 0; i < ps.length; i++) {
          const y = 1.3 + i * 0.85;
          xMark(s, 6.25, y + 0.05, C.RED);
          txt(s, ps[i].title, { x: 6.55, y, w: 2.95, h: 0.3, fontSize: 14, bold: true, wrap: false });
          if (ps[i].desc) txt(s, ps[i].desc, { x: 6.55, y: y + 0.32, w: 2.95, h: 0.5, fontSize: 11, color: C.GRAY_TEXT });
        }
      }
      if (source) txt(s, "출처: " + source, { x: M, y: 3.95, w: 9, h: 0.25, fontSize: 10, color: C.GRAY_TEXT, wrap: false });
      if (stats.length) {
        const st = limit(stats, 4, "stats");
        rect(s, 0, 4.3, W, H - 4.3, C.RED);
        const cw = W / st.length;
        st.forEach((v, i) => {
          txt(s, v.value, { x: i * cw, y: 4.4, w: cw, h: 0.6, fontSize: 32, bold: true, color: C.WHITE, align: "center", valign: "middle", wrap: false });
          txt(s, v.label, { x: i * cw + 0.2, y: 5.02, w: cw - 0.4, h: 0.45, fontSize: 11, color: C.WHITE, align: "center" });
        });
      } else cornerBlock(s);
      return s;
    },

    /** E. 섹션 — 전면 빨강, 중앙 제목 + 아이콘 3–4개 가로 배열 */
    async section({ title: t, subtitle = "", items = [], notes }) {
      const s = newSlide(C.RED, notes);
      txt(s, t, { x: M, y: items.length ? 0.5 : 1.9, w: 9, h: 0.7, fontSize: items.length ? 30 : 40, bold: true, color: C.WHITE, align: "center", valign: "middle" });
      if (subtitle) txt(s, subtitle, { x: M, y: items.length ? 1.2 : 2.75, w: 9, h: 0.45, fontSize: 14, color: C.WHITE, align: "center" });
      const its = limit(items, 4, "items");
      const n = its.length, gap = 0.3;
      const cw = n ? (9 - gap * (n - 1)) / n : 0;
      for (let i = 0; i < n; i++) {
        const it = its[i], x = M + i * (cw + gap);
        await badge(s, x + (cw - 1.1) / 2, 2.2, 1.1, C.WHITE, it.icon || "FaStar");
        txt(s, it.title, { x, y: 3.55, w: cw, h: 0.4, fontSize: 16, bold: true, color: C.WHITE, align: "center", wrap: false });
        if (it.desc) txt(s, it.desc, { x: x + 0.1, y: 4.05, w: cw - 0.2, h: 0.75, fontSize: 12, color: C.WHITE, align: "center" });
      }
      xMark(s, 0.35, 0.35, C.WHITE);
      xMark(s, 9.45, 5.1, C.WHITE);
      return s;
    },

    /** F. 마무리 — 전면 빨강, "Thank You," + 메시지 + (선택) 오른쪽 항목 3개 */
    async closing({ heading = "Thank You,", message = "", footer = "", heading2, items = [], icon = "FaHandshake", notes }) {
      const s = newSlide(C.RED, notes);
      txt(s, heading, { x: 0.6, y: 1.35, w: 4.4, h: 0.9, fontSize: 46, bold: true, color: C.WHITE, wrap: false });
      if (message) txt(s, message, { x: 0.6, y: 2.5, w: 4.4, h: 1.0, fontSize: 18, color: C.WHITE });
      if (footer) txt(s, footer, { x: 0.6, y: 4.85, w: 4, h: 0.3, fontSize: 12, bold: true, color: C.WHITE });
      if (items.length) {
        if (heading2) txt(s, heading2, { x: 5.4, y: 0.7, w: 4.1, h: 0.4, fontSize: 18, bold: true, color: C.WHITE });
        const its = limit(items, 3, "items");
        for (let i = 0; i < its.length; i++) {
          const y = 1.4 + i * 1.2;
          await badge(s, 5.4, y + 0.05, 0.75, C.WHITE, its[i].icon || "FaCheck");
          txt(s, its[i].title, { x: 6.4, y, w: 3.1, h: 0.35, fontSize: 16, bold: true, color: C.WHITE, wrap: false });
          if (its[i].desc) txt(s, its[i].desc, { x: 6.4, y: y + 0.4, w: 3.1, h: 0.7, fontSize: 12, color: C.WHITE });
        }
      } else {
        oval(s, 6.0, 1.0, 3.2, C.RED_DEEP);
        oval(s, 6.35, 1.35, 2.5, C.WHITE);
        s.addImage({ data: await iconData(icon, C.RED), x: 6.9, y: 1.9, w: 1.4, h: 1.4 });
      }
      xMark(s, 0.35, 0.35, C.WHITE);
      xMark(s, 9.45, 5.1, C.WHITE);
      return s;
    },

    /** 자유 슬라이드 — helpers로 직접 구성 */
    blank({ bg = C.WHITE, notes } = {}) { return newSlide(bg, notes); },

    async save(file) {
      await pres.writeFile({ fileName: file });
      if (warnings.length) {
        console.warn(`경고 ${warnings.length}건:`);
        warnings.forEach((w) => console.warn("  - " + w));
      }
      console.log("저장:", file);
      return warnings;
    },
  };
}

module.exports = { createDeck, wrapKo, iconData, C, FONT_KO, FONT_EN, THEME };
