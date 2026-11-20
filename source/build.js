const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const fa = require("react-icons/fa");
const gi = require("react-icons/gi");
const { renderAll } = require("./illus.js");
const { applyTheme } = require(process.env.SKILL_DIR + "/scripts/apply_theme.js");

const OUT = process.argv[2] || "deck.pptx";

const HEX = {
  dk1: "1B2633", lt1: "FFFFFF", dk2: "14273E", lt2: "E8F1F8",
  a1: "F26B1D", a2: "FFC20E", a3: "D7263D", a4: "2E9E5B", a5: "2F80ED", a6: "6B7C8F",
};
const THEME = {
  name: "Winter Road Safety",
  headFontFace: "맑은 고딕",
  bodyFontFace: "맑은 고딕",
  colors: {
    dk1: HEX.dk1, lt1: HEX.lt1, dk2: HEX.dk2, lt2: HEX.lt2,
    accent1: HEX.a1, accent2: HEX.a2, accent3: HEX.a3, accent4: HEX.a4, accent5: HEX.a5, accent6: HEX.a6,
    hlink: "2F80ED", folHlink: "6B7C8F",
  },
};

async function icon(Comp, color = "FFFFFF", size = 256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(Comp, { color: "#" + color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

(async () => {
  const img = await renderAll(__dirname + "/img");
  const pres = new pptxgen();
  pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
  pres.title = "도로관리원 제설작업 위험성평가 교육";
  pres.author = "대전광역시 도로관리소";
  pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
  const C = pres.SchemeColor;
  const COL = { navy: C.text2, ice: C.background2, white: C.background1, text: C.text1, orange: C.accent1, yellow: C.accent2, red: C.accent3, green: C.accent4, blue: C.accent5, gray: C.accent6 };

  const FOOT = "대전광역시 도로관리소 · 제설작업 위험성평가 교육";

  // ---------- Layouts ----------
  pres.defineSlideMaster({
    title: "TITLE",
    background: { color: COL.navy },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.7, y: 1.2, w: 7.6, h: 2.0, fontSize: 40, bold: true, color: COL.white, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.7, y: 3.3, w: 7.6, h: 1.0, fontSize: 18, color: COL.ice, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "SECTION",
    background: { color: COL.navy },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.9, y: 3.0, w: 11.5, h: 1.2, fontSize: 40, bold: true, color: COL.white, valign: "top", align: "left", margin: 0 }, text: "" } },
      { placeholder: { options: { name: "body", type: "body", x: 0.9, y: 4.3, w: 11.5, h: 0.9, fontSize: 18, color: COL.ice, valign: "top", align: "left", margin: 0 }, text: "" } },
    ],
  });
  pres.defineSlideMaster({
    title: "CONTENT",
    background: { color: COL.white },
    objects: [
      { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.85, fontSize: 30, bold: true, color: COL.navy, valign: "middle", align: "left", margin: 0 }, text: "" } },
      { text: { text: FOOT, options: { x: 0.6, y: 7.02, w: 8, h: 0.3, fontSize: 10, color: COL.gray, margin: 0 } } },
    ],
    slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: COL.gray, align: "right" },
  });

  // ---------- helpers ----------
  const T = (s, text, o) => s.addText(text, { isTextBox: true, fontSize: 14, color: COL.text, margin: 0, valign: "top", ...o });
  async function circleIcon(s, Comp, x, y, d, bg, name) {
    s.addShape(pres.shapes.OVAL, { x, y, w: d, h: d, fill: { color: bg }, line: { type: "none" }, objectName: name + " 원" });
    const pad = d * 0.24;
    s.addImage({ data: await icon(Comp), x: x + pad, y: y + pad, w: d - 2 * pad, h: d - 2 * pad, altText: name, objectName: name + " 아이콘" });
  }
  const card = (s, x, y, w, h, fill = COL.ice, name = "카드") =>
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.12, fill: { color: fill }, line: { type: "none" }, objectName: name });
  const bullets = (items, o = {}) => items.map((t, i) => ({ text: t, options: { bullet: { indent: 16 }, breakLine: i < items.length - 1, paraSpaceAfter: 6, ...o } }));
  const src = (s, text) => T(s, "출처: " + text, { x: 0.6, y: 6.62, w: 12.1, h: 0.3, fontSize: 10, color: COL.gray, italic: true });
  let section = "";
  const content = (title) => { const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section }); s.addText(title, { placeholder: "title" }); return s; };
  async function sectionSlide(num, title, sub, Comp, minutes, notes) {
    section = title;
    pres.addSection({ title });
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: section });
    T(s, num, { x: 0.9, y: 1.6, w: 6, h: 1.0, fontSize: 44, bold: true, color: COL.orange, fontFace: "Arial" });
    s.addText(title, { placeholder: "title" });
    s.addText(sub, { placeholder: "body" });
    await circleIcon(s, Comp, 10.3, 1.3, 1.8, COL.orange, title);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 10.35, y: 3.4, w: 1.7, h: 0.5, rectRadius: 0.25, fill: { color: COL.white, transparency: 85 }, line: { color: COL.ice, width: 1 }, objectName: "시간 배지" });
    T(s, `약 ${minutes}분`, { x: 10.35, y: 3.4, w: 1.7, h: 0.5, align: "center", valign: "middle", fontSize: 15, color: COL.white, bold: true });
    s.addNotes(notes);
    return s;
  }
  const hdr = (t) => ({ text: t, options: { bold: true, color: COL.white, fill: { color: COL.navy }, align: "center", valign: "middle", fontSize: 14 } });
  const cell = (t, o = {}) => ({ text: t, options: { fontSize: 14, valign: "middle", color: COL.text, ...o } });
  const lv = (t, c) => cell(t, { bold: true, color: COL.white, fill: { color: c }, align: "center" });

  // ===== 1. Cover =====
  section = "도입";
  pres.addSection({ title: section });
  {
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle: section });
    s.addImage({ path: img.cover, x: 0, y: 0, w: 13.333, h: 7.5, transparency: 0, altText: "눈 내리는 밤 도로 위 제설차와 신호수 그림", objectName: "표지 그림" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 8.8, h: 4.9, fill: { color: COL.navy, transparency: 20 }, line: { type: "none" }, objectName: "표지 제목 배경" });
    T(s, "2026 동절기 안전보건교육", { x: 0.7, y: 0.6, w: 6, h: 0.45, fontSize: 16, bold: true, color: COL.yellow });
    s.addText("도로관리원 제설작업\n위험성평가 교육", { placeholder: "title" });
    s.addText("유해·위험요인 발굴과 개선대책 · 근로자 의견청취", { placeholder: "body" });
    T(s, "대전광역시 도로관리소  |  교육시간 60분", { x: 0.7, y: 4.25, w: 7.6, h: 0.4, fontSize: 14, color: COL.white });
    s.addNotes("[인사·도입 1분] 오늘 교육은 제설작업 중 우리가 실제로 하는 일 — ① 제설제를 크레인으로 수급업체 차량에 실어주는 일, ② 갑작스런 한파 때 직원들이 직접 출동하는 제설작업 — 이 두 가지를 놓고 어디가 위험한지 함께 찾고, 어떻게 고칠지 정하는 시간입니다. 강의만 듣는 시간이 아니라 여러분의 의견을 듣는 시간이라는 점을 먼저 말씀드립니다.");
  }

  // ===== 2. 교육 개요 =====
  {
    const s = content("오늘 교육은 이렇게 진행합니다");
    const parts = [
      ["Ⅰ", "위험성평가 이해", "왜 하는지, 근로자는 무엇을 하는지", 8, fa.FaBookOpen, COL.blue],
      ["Ⅱ", "재해사례로 보는 겨울 위험", "실제 사고에서 배우는 교훈", 10, fa.FaExclamationTriangle, COL.red],
      ["Ⅲ", "작업① 제설제 크레인 상차", "추락·낙하·충돌 위험과 대책", 12, gi.GiCrane, COL.orange],
      ["Ⅳ", "작업② 한파 긴급출동 제설", "운전·살포·도로관리 위험과 대책", 15, fa.FaSnowflake, COL.blue],
      ["Ⅴ", "근로자 의견청취·정리", "현장 의견으로 개선대책 확정", 10, fa.FaComments, COL.green],
    ];
    let y = 1.45;
    for (const [n, t, d, m, Ic, c] of parts) {
      card(s, 0.6, y, 8.2, 0.9, COL.ice, `목차 ${n}`);
      await circleIcon(s, Ic, 0.78, y + 0.13, 0.64, c, t);
      T(s, `${n}. ${t}`, { x: 1.62, y: y + 0.1, w: 5.4, h: 0.4, fontSize: 18, bold: true, color: COL.navy });
      T(s, d, { x: 1.62, y: y + 0.5, w: 5.4, h: 0.32, fontSize: 13, color: COL.gray });
      T(s, `${m}분`, { x: 7.4, y: y + 0.2, w: 1.2, h: 0.5, fontSize: 18, bold: true, color: c, align: "right", fontFace: "Arial" });
      y += 1.05;
    }
    card(s, 9.2, 1.45, 3.5, 5.1, COL.navy, "교육 목표 카드");
    T(s, "교육 목표", { x: 9.5, y: 1.7, w: 3, h: 0.45, fontSize: 20, bold: true, color: COL.yellow });
    T(s, bullets([
      "제설작업의 주요 유해·위험요인을 안다",
      "실제 사고사례로 위험을 실감한다",
      "작업별 개선대책을 현장에서 실천한다",
      "위험을 발견하면 말하고, 함께 고친다",
    ]), { x: 9.5, y: 2.3, w: 3.0, h: 4.0, fontSize: 15, color: COL.white });
    s.addNotes("[1분] 전체 60분 구성입니다. 앞 20분은 개념과 사례, 가운데 27분은 우리 작업 두 가지의 위험요인과 대책, 마지막 10분은 여러분 의견을 듣고 개선대책을 확정합니다. 중간중간 질문 드릴 테니 편하게 말씀해 주세요.");
  }

  // ===== 3. 오늘 다룰 두 가지 작업 =====
  {
    const s = content("오늘 평가할 두 가지 작업");
    const cols = [
      { x: 0.6, c: COL.orange, Ic: gi.GiCrane, t: "작업① 제설제 상차 지원", items: ["제설제(톤백·포대)를 크레인으로 수급업체 살포차량에 실어줌", "제설제 살포는 수급업체가 자체 실시", "주요 위험: 추락, 낙하물 맞음, 끼임, 차량 충돌·협착"] },
      { x: 6.85, c: COL.blue, Ic: fa.FaSnowflake, t: "작업② 한파 시 직접 출동 제설", items: ["급작스런 한파·결빙 시 직원이 직접 출동", "운전 → 제설제 살포 → 도로관리(통제·순찰)", "주요 위험: 빙판 미끄럼, 교통사고·2차사고, 한랭질환"] },
    ];
    for (const k of cols) {
      card(s, k.x, 1.45, 5.9, 5.0, COL.ice, k.t);
      await circleIcon(s, k.Ic, k.x + 0.35, 1.75, 1.0, k.c, k.t);
      T(s, k.t, { x: k.x + 1.55, y: 1.9, w: 4.2, h: 0.7, fontSize: 21, bold: true, color: COL.navy, valign: "middle" });
      T(s, bullets(k.items, { paraSpaceAfter: 14 }), { x: k.x + 0.4, y: 3.1, w: 5.1, h: 3.2, fontSize: 18 });
    }
    s.addNotes("[2분] 오늘 평가 대상은 두 작업입니다. 첫째, 우리가 크레인으로 제설제를 수급업체 차량에 실어주는 일. 살포는 수급업체가 하지만, 상차 장소는 우리 사업장이고 우리 직원이 크레인을 다루므로 우리가 관리해야 할 위험입니다. 둘째, 갑작스런 한파 때 직원이 직접 운전해서 나가 제설제를 뿌리고 도로를 관리하는 일입니다. 질문: 지난 겨울 이 두 작업 중 '아찔했다' 싶은 순간이 있었나요?");
  }

  // ===== Ⅰ =====
  await sectionSlide("PART 1", "위험성평가 이해", "위험을 찾고 · 판단하고 · 줄이는 활동", fa.FaBookOpen, 8,
    "[섹션 전환] 먼저 위험성평가가 무엇인지, 왜 근로자의 참여가 꼭 필요한지 짧게 정리합니다.");

  {
    const s = content("위험성평가란 무엇인가요?");
    card(s, 0.6, 1.4, 12.1, 1.15, COL.navy, "정의 카드");
    T(s, [
      { text: "위험성평가 ", options: { bold: true, color: COL.yellow } },
      { text: "= 일터의 유해·위험요인을 찾아내고, 그 위험이 얼마나 큰지 판단하여, 줄이는 대책을 세우고 실행하는 활동", options: { color: COL.white } },
    ], { x: 0.9, y: 1.4, w: 11.5, h: 1.15, fontSize: 18, valign: "middle" });
    const steps = [
      ["1", "사전준비", "평가 대상 작업\n정하기", fa.FaClipboardCheck],
      ["2", "위험요인 파악", "현장 순회·\n근로자 의견", fa.FaSearch],
      ["3", "위험성 결정", "가능성 × 중대성\n크기 판단", fa.FaBalanceScale],
      ["4", "감소대책 실행", "제거·대체·공학·\n관리·보호구", fa.FaTools],
      ["5", "공유·기록", "TBM·게시로\n모두가 알기", fa.FaShareAlt],
    ];
    const w = 2.15, gap = 0.34;
    for (let i = 0; i < steps.length; i++) {
      const [n, t, d, Ic] = steps[i];
      const x = 0.6 + i * (w + gap);
      card(s, x, 2.95, w, 2.75, COL.ice, `단계 ${n}`);
      await circleIcon(s, Ic, x + (w - 0.85) / 2, 3.15, 0.85, i === 3 ? COL.orange : COL.navy, t);
      T(s, `${n}. ${t}`, { x: x + 0.1, y: 4.1, w: w - 0.2, h: 0.45, fontSize: 16, bold: true, color: COL.navy, align: "center" });
      T(s, d, { x: x + 0.1, y: 4.6, w: w - 0.2, h: 0.9, fontSize: 13, color: COL.text, align: "center" });
      if (i < steps.length - 1) s.addShape(pres.shapes.CHEVRON, { x: x + w + 0.08, y: 4.1, w: 0.18, h: 0.36, fill: { color: COL.gray }, line: { type: "none" }, flipH: false, objectName: "화살표" });
    }
    T(s, [
      { text: "법적 근거  ", options: { bold: true, color: COL.navy } },
      { text: "산업안전보건법 제36조 · 「사업장 위험성평가에 관한 지침」(고용노동부 고시)   |   ", options: { color: COL.text } },
      { text: "정기평가 연 1회 + 수시평가", options: { bold: true, color: COL.orange } },
      { text: "(장비·작업방법 변경, 사고 발생 시)", options: { color: COL.text } },
    ], { x: 0.6, y: 5.95, w: 12.1, h: 0.5, fontSize: 14, valign: "middle" });
    s.addNotes("[3분] 위험성평가는 서류 작업이 아니라 '위험을 찾고, 크기를 판단하고, 줄이는' 일입니다. 다섯 단계 중 가장 중요한 건 2단계 위험요인 파악과 4단계 감소대책 실행입니다. 대책은 순서가 있습니다: 위험 자체를 없애기(제거) → 덜 위험한 방법으로 바꾸기(대체) → 설비로 막기(공학적 대책) → 절차·교육(관리적 대책) → 마지막이 보호구입니다. 보호구는 최후의 수단이라는 점을 기억해 주세요. 정기평가는 연 1회, 새 장비를 들이거나 작업 방법이 바뀌거나 사고가 나면 수시평가를 합니다.");
  }

  {
    const s = content("위험은 현장 근로자가 가장 잘 압니다");
    s.addShape(pres.shapes.OVAL, { x: 0.8, y: 1.6, w: 4.2, h: 4.2, fill: { color: COL.ice }, line: { type: "none" }, objectName: "강조 원" });
    await circleIcon(s, fa.FaUsers, 1.65, 2.25, 2.5, COL.navy, "근로자 참여");
    T(s, "근로자 참여는\n선택이 아닌 필수", { x: 0.8, y: 4.85, w: 4.2, h: 0.9, fontSize: 18, bold: true, color: COL.navy, align: "center" });
    const rows = [
      [fa.FaSearch, COL.blue, "위험요인 찾기", "아차사고, 불편한 작업, \"이러다 다치겠다\" 싶은 순간을 말해 주세요"],
      [fa.FaLightbulb, COL.orange, "개선대책 제안", "현장을 아는 사람의 아이디어가 가장 실용적인 대책이 됩니다"],
      [fa.FaShareAlt, COL.green, "결과 공유·실천", "작업 전 TBM에서 위험과 대책을 확인하고 지킵니다"],
      [fa.FaHandPaper, COL.red, "작업중지권", "급박한 위험이 있으면 작업을 멈추고 대피할 수 있습니다 (산안법 제52조)"],
    ];
    let y = 1.5;
    for (const [Ic, c, t, d] of rows) {
      await circleIcon(s, Ic, 5.7, y, 0.8, c, t);
      T(s, t, { x: 6.7, y: y - 0.02, w: 6, h: 0.42, fontSize: 18, bold: true, color: COL.navy });
      T(s, d, { x: 6.7, y: y + 0.42, w: 6, h: 0.6, fontSize: 14 });
      y += 1.27;
    }
    s.addNotes("[3분] 2023년 개정된 위험성평가 지침은 근로자 참여를 크게 강화했습니다. 위험요인을 찾을 때, 대책을 세울 때, 결과를 공유할 때 모두 근로자가 참여해야 합니다. 그리고 꼭 기억할 것 — 급박한 위험이 있으면 근로자는 작업을 중지하고 대피할 수 있고, 이것으로 불이익을 주면 안 됩니다(산업안전보건법 제52조). 오늘 마지막 10분은 여러분의 의견을 직접 듣는 시간입니다.");
  }

  // ===== Ⅱ =====
  await sectionSlide("PART 2", "재해사례로 보는 겨울 위험", "\"설마\"가 사고가 된 실제 사례들", fa.FaExclamationTriangle, 10,
    "[섹션 전환] 이제 실제 사고 사례를 보겠습니다. 모두 언론과 공공기관이 발표한 실제 사례입니다.");

  {
    const s = content("숫자로 보는 겨울철 도로 작업의 위험");
    const stats = [
      ["1위", "화물차 적재함 등\n운반차량에서 떨어짐", "2025년 떨어짐 사망 단일 기인물 1위\n(고용노동부 통계 · CBS노컷뉴스 2026.6)", COL.red],
      ["1.5~1.8배", "결빙 노면 사고\n치사율", "마른 노면 대비\n(한국교통안전공단 분석)", COL.blue],
      ["5배↑", "2차사고\n치사율", "일반 교통사고 대비\n(한국도로공사)", COL.orange],
      ["6명 중 1명", "고속도로 사망자 중\n2차사고 희생자", "최근 5년 784명 중 138명\n(한국도로공사)", COL.navy],
    ];
    const w = 2.85, gap = 0.233;
    for (let i = 0; i < stats.length; i++) {
      const [n, l, d, c] = stats[i];
      const x = 0.6 + i * (w + gap);
      card(s, x, 1.5, w, 3.85, COL.ice, `통계 ${i + 1}`);
      T(s, n, { x: x + 0.15, y: 1.8, w: w - 0.3, h: 1.1, fontSize: n.length > 4 ? 34 : 48, bold: true, color: c, align: "center", valign: "middle", fit: "shrink" });
      T(s, l, { x: x + 0.15, y: 3.05, w: w - 0.3, h: 1.0, fontSize: 17, bold: true, color: COL.navy, align: "center" });
      T(s, d, { x: x + 0.15, y: 4.2, w: w - 0.3, h: 1.0, fontSize: 12, color: COL.gray, align: "center" });
    }
    {
      card(s, 0.6, 5.6, 12.1, 0.85, COL.navy, "통계 메시지");
      T(s, [{ text: "도로 위, 적재함 위 — ", options: { bold: true, color: COL.yellow } }, { text: "우리가 일하는 곳이 바로 이 숫자들의 현장입니다", options: { color: COL.white } }], { x: 0.9, y: 5.6, w: 11.5, h: 0.85, fontSize: 18, valign: "middle", align: "center" });
    }
    s.addNotes("[2분] 네 가지 숫자만 기억해 주세요. ① 지난해 떨어짐 사망 중 기인물 1위가 화물차 적재함 같은 운반차량입니다. 높이 2m도 안 되지만 머리부터 떨어지면 치명적입니다 — 우리 상차 작업과 직결됩니다. ② 얼어붙은 길의 사고는 마른 길보다 치사율이 1.5~1.8배 높습니다. ③ 사고 현장이나 작업 현장에서 뒤차가 덮치는 2차사고는 일반 사고보다 치사율이 5배 이상입니다. ④ 고속도로 사망자 6명 중 1명이 2차사고로 숨졌습니다. 도로 위에 서 있는 우리가 바로 그 위험에 노출돼 있습니다.");
  }

  async function caseSlide(title, tag, Ic, color, facts, causes, lessons, source, notes, imgPath) {
    const s = content(title);
    if (imgPath) {
      s.addImage({ path: imgPath, x: 0.6, y: 1.45, w: 4.6, h: 3.17, altText: tag, objectName: "사례 그림" });
    } else {
      card(s, 0.6, 1.45, 4.6, 3.17, color, "사례 아이콘 패널");
      await circleIcon(s, Ic, 1.95, 1.85, 1.9, COL.navy, tag);
    }
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y: 4.82, w: 4.6, h: 1.6, rectRadius: 0.12, fill: { color: COL.navy }, line: { type: "none" }, objectName: "사례 태그" });
    T(s, tag, { x: 0.85, y: 4.82, w: 4.1, h: 1.6, fontSize: 17, bold: true, color: COL.white, valign: "middle" });
    T(s, "사고 개요", { x: 5.6, y: 1.45, w: 7, h: 0.4, fontSize: 17, bold: true, color: COL.red });
    T(s, bullets(facts), { x: 5.6, y: 1.9, w: 7.1, h: 1.45, fontSize: 15 });
    T(s, "원인", { x: 5.6, y: 3.4, w: 7, h: 0.4, fontSize: 17, bold: true, color: COL.orange });
    T(s, bullets(causes), { x: 5.6, y: 3.85, w: 7.1, h: 1.0, fontSize: 15 });
    card(s, 5.6, 4.95, 7.1, 1.5, COL.ice, "교훈 카드");
    T(s, [{ text: "우리 현장의 교훈  ", options: { bold: true, color: COL.green, breakLine: true } }, ...bullets(lessons)], { x: 5.85, y: 5.05, w: 6.7, h: 1.35, fontSize: 15 });
    src(s, source);
    s.addNotes(notes);
  }

  await caseSlide("사례 1 · 폭설 속 도로통제 작업자, 미끄러진 버스에 치여 사망",
    "폭설 · 도로 통제 중\n눈길에 미끄러진 차량", fa.FaCarCrash, COL.yellow,
    ["2024년 11월 기록적 첫 폭설 당시 고속도로 사고 현장", "사고 처리를 위해 도로를 통제하던 30대 작업자가 눈길에 미끄러진 버스에 치여 숨짐"],
    ["눈길·빙판에서는 차량 제동거리가 크게 늘어나 작업자 앞에서 멈추지 못함", "작업자가 차량 진행 방향 정면에 노출"],
    ["눈길 차량은 \"멈추지 못한다\"고 가정하고, 차로 안이 아닌 갓길·가드레일 밖에서 통제"],
    "뉴데일리 2024.11.28 · 뉴시스 2024.11.28 보도",
    "[2분] 2024년 11월 첫 폭설 때입니다. 고속도로에서 사고를 수습하려고 도로를 통제하던 30대 작업자가, 눈길에 미끄러진 버스에 치여 숨졌습니다. 같은 이틀 동안 제설작업 중 숨진 분이 3명 더 있었습니다(무너진 구조물·쓰러진 나무에 깔림). 교훈: 눈길에서 다가오는 차는 내 앞에서 멈추지 못할 수 있습니다. 통제할 때는 차로 안에 서지 말고, 갓길이나 가드레일 바깥에서, 차량을 마주 보며, 도망갈 길을 확보한 채 신호해야 합니다.");

  await caseSlide("사례 2 · 사고수습 중 작업차량이 추돌당해 직원 사망",
    "2차사고 · 작업차량 추돌\n밀려난 차량에 작업자 충돌", fa.FaRoad, COL.orange,
    ["2024년 고속도로 사고 현장에 출동한 도로공사 차량을 뒤따르던 차량이 추돌", "밀려난 작업차량에 후속조치 중이던 50대 직원이 부딪혀 숨짐"],
    ["후방 운전자의 전방주시 태만(2차사고 원인의 76%)·졸음운전", "작업자가 작업차량 바로 앞에서 작업"],
    ["작업차량 바로 앞·옆은 \"밀려올 수 있는 위험구역\" — 차량 앞쪽 작업 금지, 충분한 간격 확보"],
    "뉴시스 2024.10.21 · 경인일보(한국도로공사 국회 제출자료) 보도",
    "[2분] 사고를 수습하러 출동한 도로공사 차량을 뒤차가 들이받았고, 그 충격으로 앞으로 밀린 작업차량에 앞에서 일하던 직원이 숨졌습니다. 안전순찰원은 최근 6년간 2차사고로 17명이 다치거나 숨졌습니다. 2차사고 원인의 90% 이상이 뒤차 운전자의 전방주시 태만과 졸음운전입니다. 즉, 우리가 아무리 조심해도 뒤에서 덮칠 수 있습니다. 그래서 작업차량(안전차량)은 작업자보다 '뒤쪽'에 두어 방패로 쓰고, 그 차량 바로 앞에는 서 있지 않습니다.");

  await caseSlide("사례 3 · 화물차 적재함에서 하역 중 떨어져 사망",
    "적재함 위 작업 · 끝단 추락\n높이 약 2m에서도 사망", fa.FaTruck, COL.orange,
    ["2023년 12월 충북 청주, 화물차에 실린 철근 다발을 지렛대로 내리던 작업자", "적재함 끝단에서 철근과 함께 떨어져 숨짐"],
    ["작업계획 미수립 · 작업지휘자 미배치", "적재함 끝단 추락방지 조치 없이 작업"],
    ["톤백 줄걸이·해체는 적재함 위가 아닌 지상에서 — 오를 때는 승강설비 사용"],
    "안전보건공단 재해사례 · CBS노컷뉴스 2026.6.22 보도",
    "[2분] 철근을 내리던 근로자가 적재함 끝에서 떨어져 숨진 사고입니다. 적재함 높이는 대개 2m 안팎인데 '이 정도쯤이야' 하다가 머리부터 떨어지면 사망에 이릅니다. 지난해 떨어짐 사망의 단일 기인물 1위가 바로 화물차 적재함입니다. 우리 제설제 상차 때도 톤백 고리를 걸거나 풀려고 적재함 위에 올라가는 경우가 있죠? 눈·얼음이 있는 겨울 적재함은 더 미끄럽습니다. 오늘 대책의 핵심: 적재함 위에 올라가지 않는 방법을 찾는 것입니다.",
    img.fall);

  await caseSlide("사례 4 · 블랙아이스 연쇄추돌, 제설 부실이 피해를 키웠다",
    "블랙아이스 · 다중추돌\n제설·순찰 공백", fa.FaSnowflake, COL.blue,
    ["2026년 1월 10일 새벽 서산영덕고속도로 남상주IC 부근, 블랙아이스 추정 다중추돌로 5명 사망", "국토교통부 감사: 제설제 살포 미흡 + 초기 대응 지연이 피해를 키움"],
    ["강설 시 순찰 공백 · 비상근무자 업무분담 미비 · 현장 대응 교육 부족(경찰 수사 지적)"],
    ["출동 기준·사전 살포·순찰 체계를 지키는 것이 도로 이용자의 생명을 지킴"],
    "SBS Biz 2026.1 · YTN·경북매일 2026.2.11 보도",
    "[2분] 올해 1월 상주 인근 고속도로에서 블랙아이스로 추정되는 연쇄추돌로 5명이 숨졌습니다. 국토부 감사 결과 제설제가 제대로 뿌려지지 않았고 초기 대응이 늦었다는 지적이 나왔고, 경찰은 순찰 공백, 비상근무자 업무 분담 미비, 현장 대응 교육 부족을 문제로 봤습니다. 우리 도로관리원의 출동 기준과 비상연락·업무분담이 왜 중요한지 보여주는 사례입니다. 해외에서도 제설차가 뒤차에 받히는 사고가 잦습니다(미국 위스콘신주에서 한 겨울 제설차 13대 피추돌, 대부분 후방추돌).");

  // ===== Ⅲ =====
  await sectionSlide("PART 3", "작업① 제설제 크레인 상차", "수급업체 살포차량에 제설제를 실어주는 작업", gi.GiCrane, 12,
    "[섹션 전환] 이제 우리 작업입니다. 첫 번째, 제설제를 크레인으로 수급업체 차량에 실어주는 작업입니다.");

  {
    const s = content("그림으로 보는 상차 작업의 위험 포인트");
    s.addImage({ path: img.crane, x: 0.6, y: 1.4, w: 7.6, h: 4.13, altText: "크레인으로 톤백을 덤프트럭에 싣는 현장과 6개 위험 지점", objectName: "상차 현장 그림" });
    const hz = [
      ["1", "적재함 위 작업자 추락", "톤백 고리 해체하려 적재함에 올라감"],
      ["2", "인양물 아래 출입", "흔들리는 톤백에 맞음·깔림"],
      ["3", "톤백 낙하", "슬링 끊어짐·훅 해지장치 불량"],
      ["4", "크레인 전도·접촉", "아웃트리거 미설치, 얼은 지반"],
      ["5", "차량 진입·후진 충돌", "수급업체 차량과 작업자 협착"],
      ["6", "빙판 미끄러짐", "야적장 바닥 결빙, 녹은 염수"],
    ];
    let y = 1.4;
    for (const [n, t, d] of hz) {
      s.addShape(pres.shapes.OVAL, { x: 8.55, y: y + 0.05, w: 0.5, h: 0.5, fill: { color: COL.red }, line: { type: "none" }, objectName: `위험 ${n}` });
      T(s, n, { x: 8.55, y: y + 0.05, w: 0.5, h: 0.5, fontSize: 16, bold: true, color: COL.white, align: "center", valign: "middle", fontFace: "Arial" });
      T(s, t, { x: 9.2, y, w: 3.6, h: 0.35, fontSize: 15, bold: true, color: COL.navy });
      T(s, d, { x: 9.2, y: y + 0.36, w: 3.6, h: 0.32, fontSize: 12, color: COL.gray });
      y += 0.86;
    }
    card(s, 0.6, 5.75, 7.6, 0.75, COL.ice, "질문 카드");
    T(s, [{ text: "함께 생각해 봅시다  ", options: { bold: true, color: COL.orange } }, { text: "우리 야적장에서 가장 자주 보는 위험은 몇 번인가요?" }], { x: 0.85, y: 5.75, w: 7.2, h: 0.75, fontSize: 15, valign: "middle" });
    s.addNotes("[3분] 그림 속 번호를 하나씩 짚어 보겠습니다. 1번, 받는 차량 적재함 위에 올라가 고리를 푸는 작업자 — 가장 흔하고 가장 위험합니다. 2번, 매달린 톤백 바로 아래 서 있는 사람. 3번, 슬링이 끊어지거나 훅에서 빠져 떨어지는 톤백(1톤 백이면 맞으면 사망). 4번, 아웃트리거를 덜 펴거나 얼어붙은 땅 위에 받쳐서 크레인이 넘어지는 경우, 붐이 전선·구조물에 닿는 경우. 5번, 수급업체 차량이 들어오고 후진하면서 사람을 치거나 끼이게 하는 경우. 6번, 야적장 바닥의 얼음과 녹은 염수에 미끄러지는 경우. 질문: 우리 야적장에서 가장 자주 보는 건 몇 번인가요? (손 들기)");
  }

  {
    const s = content("작업① 유해·위험요인과 개선대책");
    const rows = [
      [hdr("위험요인"), hdr("위험성"), hdr("개선대책 (현장 실천)")],
      [cell("적재함 위 올라가 줄걸이 해체 → 추락", { bold: true }), lv("높음", COL.red), cell("지상에서 해체 가능한 훅·슬링 사용, 부득이 오를 때 이동식 승강설비(사다리·발판) 사용, 적재함 위 이동 금지")],
      [cell("인양물 아래 출입 → 맞음·깔림", { bold: true }), lv("높음", COL.red), cell("작업반경 출입금지 구획(콘·라바콘·안전띠), 유도로프(태그라인)로 손대지 않고 방향 조정")],
      [cell("슬링 파손·훅 이탈 → 톤백 낙하", { bold: true }), lv("높음", COL.red), cell("작업 전 슬링·와이어로프·훅 해지장치 점검, 찢어진·젖어 언 톤백 사용 금지, 정격하중 준수")],
      [cell("아웃트리거 미설치·결빙 지반 → 크레인 전도", { bold: true }), lv("보통", COL.orange), cell("아웃트리거 완전 확장 + 받침판 사용, 얼음·눈 제거 후 평탄한 곳에 설치, 붐 주변 전선·구조물 확인")],
      [cell("수급업체 차량 진입·후진 → 충돌·협착", { bold: true }), lv("높음", COL.red), cell("유도자 배치·후진 신호 통일, 보행자 동선과 차량 동선 분리, 차량 정차 후 시동 끄고 고임목")],
      [cell("결빙 바닥·염수 → 넘어짐", { bold: true }), lv("보통", COL.orange), cell("작업장 바닥 제설·모래 살포, 미끄럼방지 안전화·아이젠, 야간 조명 확보")],
    ];
    s.addTable(rows, { x: 0.6, y: 1.4, w: 12.1, colW: [3.9, 1.1, 7.1], rowH: 0.74, border: { type: "solid", pt: 1, color: "C9D6E2" }, fontSize: 13, valign: "middle", margin: 0.1, objectName: "작업1 개선대책 표" });
    s.addNotes("[4분] 앞에서 본 위험을 표로 정리했습니다. 위험성 '높음'은 사망이나 중상으로 이어질 수 있고 자주 노출되는 것들입니다. 핵심 대책 세 가지만 기억합시다. 첫째, 적재함 위에 올라가지 않는다 — 지상에서 고리를 걸고 풀 수 있는 방법을 쓰고, 꼭 올라가야 하면 승강설비를 씁니다. 둘째, 매달린 짐 아래는 아무도 들어가지 않는다 — 콘으로 구역을 표시하고, 방향은 손이 아니라 유도로프로 잡습니다. 셋째, 수급업체 차량이 움직일 때는 유도자 한 명이 신호를 통일합니다. 표의 위험성 판단은 예시이니, 의견청취 시간에 여러분 판단으로 수정하겠습니다.");
  }

  {
    const s = content("크레인 작업 전 점검과 표준 신호");
    T(s, "작업 전 5분 점검", { x: 0.6, y: 1.4, w: 5.6, h: 0.45, fontSize: 20, bold: true, color: COL.navy });
    const checks = ["크레인 운전 자격·작업계획 확인", "와이어로프·슬링 손상(끊김·꼬임) 없음", "훅 해지장치 정상 작동", "아웃트리거 완전 확장·받침판 설치", "작업반경 출입금지 구획 설치", "신호수 1명 지정, 신호 방법 통일", "수급업체 운전원과 작업 순서 협의"];
    let y = 2.0;
    for (const c of checks) {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 0.6, y, w: 0.38, h: 0.38, rectRadius: 0.06, fill: { color: COL.white }, line: { color: COL.navy, width: 1.5 }, objectName: "체크박스" });
      T(s, c, { x: 1.15, y: y - 0.02, w: 5.1, h: 0.42, fontSize: 15, valign: "middle" });
      y += 0.62;
    }
    T(s, "크레인 표준 신호 (신호수 1명만)", { x: 6.7, y: 1.4, w: 6, h: 0.45, fontSize: 20, bold: true, color: COL.navy });
    const sig = [
      [fa.FaBullhorn, COL.blue, "운전자 호출", "호각 등으로 주의 집중"],
      [fa.FaArrowUp, COL.green, "올리기(권상)", "집게손가락 위로, 수평 원"],
      [fa.FaArrowDown, COL.orange, "내리기(권하)", "팔 아래로, 집게손가락 원"],
      [fa.FaHandPaper, COL.red, "정지", "한 손 들어 손바닥 펴기"],
      [fa.FaExclamationTriangle, COL.red, "비상정지", "양손 들어 크게 흔들기"],
    ];
    const w = 1.95;
    for (let i = 0; i < sig.length; i++) {
      const [Ic, c, t, d] = sig[i];
      const col = i % 3, row = Math.floor(i / 3);
      const x = 6.7 + col * (w + 0.1), yy = 2.0 + row * 2.3;
      card(s, x, yy, w, 2.1, COL.ice, `신호 ${t}`);
      await circleIcon(s, Ic, x + (w - 0.8) / 2, yy + 0.18, 0.8, c, t);
      T(s, t, { x: x + 0.1, y: yy + 1.05, w: w - 0.2, h: 0.4, fontSize: 15, bold: true, color: COL.navy, align: "center" });
      T(s, d, { x: x + 0.1, y: yy + 1.45, w: w - 0.2, h: 0.55, fontSize: 12, align: "center" });
    }
    card(s, 6.7 + 2 * (w + 0.1), 4.3, w, 2.1, COL.navy, "신호 원칙");
    T(s, "신호수는\n1명!\n비상정지는\n누구나", { x: 6.7 + 2 * (w + 0.1) + 0.1, y: 4.35, w: w - 0.2, h: 2.0, fontSize: 15, bold: true, color: COL.yellow, align: "center", valign: "middle" });
    s.addNotes("[3분] 크레인을 움직이기 전 5분 점검 항목입니다. 특히 겨울엔 슬링이 젖었다 얼면서 손상되기 쉽고, 땅이 얼어 아웃트리거가 미끄러질 수 있습니다. 신호는 반드시 정해진 한 사람만 합니다 — 여러 명이 동시에 손짓하면 운전원이 헷갈립니다. 단, 위험을 본 사람은 누구든 '비상정지' 신호를 보낼 수 있습니다. (가능하면 교육 중 정지·비상정지 신호를 함께 해 봅니다.) 수급업체와도 신호를 통일해야 합니다 — 상차장은 우리 사업장이므로 도급인으로서 안전조치·협의 책임이 있습니다(산안법 제63조·제64조).");
  }

  // ===== Ⅳ =====
  await sectionSlide("PART 4", "작업② 한파 긴급출동 제설", "직원이 직접 운전 → 제설제 살포 → 도로관리", fa.FaSnowflake, 15,
    "[섹션 전환] 두 번째 작업, 갑작스러운 한파 때 우리가 직접 나가는 제설작업입니다.");

  {
    const s = content("긴급출동 단계별 유해·위험요인");
    const steps = [
      [fa.FaClipboardCheck, COL.navy, "출동 준비", ["야간·새벽 급출동 → 서두름", "제설제 적재 중 미끄러짐", "살포기 점검 중 끼임"]],
      [fa.FaTruck, COL.blue, "운전·이동", ["블랙아이스 미끄러짐", "시야 불량·피로·졸음", "제동거리 증가 → 추돌"]],
      [fa.FaSnowflake, COL.orange, "제설제 살포", ["살포기 회전부 끼임", "염화칼슘 눈·피부 자극", "저속 주행 중 후방 추돌"]],
      [fa.FaRoad, COL.red, "도로관리", ["차량통제 중 차량 충돌", "2차사고 (뒤차 돌진)", "갓길 결빙 넘어짐"]],
    ];
    const w = 2.85, gap = 0.233;
    for (let i = 0; i < steps.length; i++) {
      const [Ic, c, t, items] = steps[i];
      const x = 0.6 + i * (w + gap);
      card(s, x, 1.45, w, 4.0, COL.ice, `단계 ${t}`);
      await circleIcon(s, Ic, x + (w - 0.95) / 2, 1.7, 0.95, c, t);
      T(s, `${i + 1}. ${t}`, { x: x + 0.15, y: 2.8, w: w - 0.3, h: 0.45, fontSize: 18, bold: true, color: COL.navy, align: "center" });
      T(s, bullets(items), { x: x + 0.25, y: 3.35, w: w - 0.4, h: 2.0, fontSize: 14 });
    }
    card(s, 0.6, 5.65, 12.1, 0.85, COL.navy, "공통 위험");
    await circleIcon(s, fa.FaThermometerQuarter, 0.8, 5.75, 0.65, COL.blue, "한랭");
    T(s, [{ text: "모든 단계 공통  ", options: { bold: true, color: COL.yellow } }, { text: "한랭질환(저체온증·동상) · 장시간 연속근무 피로 · 혼자 작업 시 사고 발견 지연", options: { color: COL.white } }], { x: 1.65, y: 5.65, w: 10.9, h: 0.85, fontSize: 15, valign: "middle" });
    s.addNotes("[3분] 긴급출동은 '준비 → 운전 → 살포 → 도로관리' 네 단계로 나눠 봤습니다. 갑작스러운 한파는 대부분 밤이나 새벽에 옵니다. 급하게 나가다 보니 점검을 건너뛰고, 피곤한 상태로 운전합니다. 살포기 회전 날개에 장갑이나 옷이 말려 들어가는 끼임 사고, 저속으로 살포하는 중 뒤차에 받히는 사고, 그리고 도로를 통제하다 차에 치이는 사고가 핵심 위험입니다. 모든 단계에 공통으로 추위, 피로, 혼자 작업하는 위험이 있습니다.");
  }

  {
    const s = content("신호수 배치와 차량통제 — 작업구간 이렇게 만듭니다");
    s.addImage({ path: img.zone, x: 0.6, y: 1.35, w: 12.1, h: 3.93, altText: "도로 작업구간 배치도: 주의구간, 완화구간, 작업구간, 종료구간", objectName: "작업구간 배치도" });
    const z = [
      ["1", "주의구간", "경고표지 단계 설치 · 신호수가 갓길에서 감속 유도", COL.yellow],
      ["2", "완화구간", "라바콘을 비스듬히 배치해 차로를 서서히 줄임 · 안전차량(방호차) 배치", COL.orange],
      ["3", "작업구간", "작업자는 안전차량 앞쪽 구획 안에서만 작업", COL.red],
      ["4", "종료구간", "콘으로 원래 차로 복귀 유도", COL.green],
    ];
    const w = 2.95, gap = 0.1;
    for (let i = 0; i < z.length; i++) {
      const [n, t, d, c] = z[i];
      const x = 0.6 + i * (w + gap);
      card(s, x, 5.4, w, 1.15, COL.ice, `구간 ${n}`);
      s.addShape(pres.shapes.OVAL, { x: x + 0.12, y: 5.5, w: 0.42, h: 0.42, fill: { color: c }, line: { type: "none" }, objectName: `구간 번호 ${n}` });
      T(s, n, { x: x + 0.12, y: 5.5, w: 0.42, h: 0.42, fontSize: 14, bold: true, color: COL.white, align: "center", valign: "middle", fontFace: "Arial" });
      T(s, t, { x: x + 0.65, y: 5.5, w: w - 0.75, h: 0.42, fontSize: 15, bold: true, color: COL.navy, valign: "middle" });
      T(s, d, { x: x + 0.12, y: 5.95, w: w - 0.24, h: 0.58, fontSize: 11, color: COL.text });
    }
    src(s, "국토교통부 「도로공사장 교통관리지침」의 작업구간 구성(주의·완화·작업·종료구간)을 쉽게 재구성");
    s.addNotes("[4분] 도로 위 작업구간은 네 구간으로 만듭니다. ① 주의구간: 운전자가 미리 알 수 있게 경고표지를 거리를 두고 여러 개 세우고, 신호수는 갓길에서 차를 마주 보며 깃발·경광봉으로 감속을 유도합니다. 신호수 뒤에는 피할 공간이 있어야 합니다. ② 완화구간: 라바콘을 비스듬히 놓아 차로를 천천히 줄이고, 작업자보다 뒤쪽에 안전차량(방호차량, 경광등·화살표 표시)을 세워 뒤차가 덮치면 차가 먼저 맞도록 합니다. 사례 2처럼 차가 밀려올 수 있으니 안전차량과 작업자 사이에 간격을 둡니다. ③ 작업구간: 작업자는 구획 안에서만, 차량을 등지지 않고 일합니다. ④ 종료구간: 콘으로 원래 차로로 안내합니다. 눈길에서는 제동거리가 길어지므로 평소보다 경고구간을 더 길게 잡습니다. 설치는 차량 진행 방향 순서대로(주의→완화→작업), 철수는 역순으로 합니다.");
  }

  {
    const s = content("빙판길 운전과 제설제 살포 안전수칙");
    s.addImage({ path: img.ice, x: 0.6, y: 1.4, w: 4.6, h: 3.17, altText: "빙판길에서 미끄러지는 차량과 결빙 주의 표지", objectName: "빙판 그림" });
    card(s, 0.6, 4.75, 4.6, 1.75, COL.navy, "블랙아이스 카드");
    T(s, [{ text: "블랙아이스 주의 구간", options: { bold: true, color: COL.yellow, breakLine: true } }, { text: "교량·고가도로 · 터널 출입구 · 그늘진 곡선부 · 새벽 시간대", options: { color: COL.white } }], { x: 0.85, y: 4.85, w: 4.1, h: 1.55, fontSize: 15, valign: "middle" });
    T(s, "운전", { x: 5.6, y: 1.4, w: 3.4, h: 0.45, fontSize: 19, bold: true, color: COL.blue });
    T(s, bullets([
      "평소 속도의 절반 이하로 감속",
      "차간거리 평소의 2배 이상",
      "급제동·급핸들·급가속 금지",
      "출발 전 경광등·후사경·타이어 체인 점검",
      "2시간 운전마다 휴식, 졸음 시 즉시 정차",
      "운전 중 휴대전화 사용 금지",
    ]), { x: 5.6, y: 1.95, w: 3.4, h: 4.5, fontSize: 14 });
    T(s, "살포", { x: 9.3, y: 1.4, w: 3.4, h: 0.45, fontSize: 19, bold: true, color: COL.orange });
    T(s, bullets([
      "살포기 막힘 제거는 반드시 엔진·동력 정지 후",
      "회전부 주변에 손·장갑·옷자락 접근 금지",
      "염화칼슘 취급 시 보안경·방수장갑 착용",
      "적재 시 포대는 2인 1조, 허리 아닌 다리로 들기",
      "저속 살포 중 경광등·후방 표시 상시 작동",
    ]), { x: 9.3, y: 1.95, w: 3.4, h: 4.5, fontSize: 14 });
    src(s, "한국교통안전공단·도로교통공단 겨울철 안전운전 자료 재구성");
    s.addNotes("[3분] 빙판길에서는 제동거리가 평소의 몇 배로 늘어납니다. 그래서 속도는 절반, 차간거리는 두 배. 핸들·브레이크는 부드럽게. 블랙아이스는 눈에 잘 보이지 않고 교량, 터널 입구, 그늘진 곳, 새벽에 많습니다. 살포 작업에서 제일 위험한 건 살포기가 막혔을 때 손으로 뚫는 행동입니다 — 반드시 동력을 끄고, 막대 같은 도구로 처리합니다. 염화칼슘은 눈에 들어가면 화상을 입을 수 있으니 보안경을 씁니다. 저속 살포 중에는 뒤차가 빨리 다가오니 경광등을 계속 켭니다.");
  }

  {
    const s = content("미끄러짐·한랭질환 예방");
    const items = [
      [fa.FaWalking, COL.blue, "넘어짐 예방", ["미끄럼방지 안전화·아이젠 착용", "차량 승하차 시 3점 접촉(두 손+한 발)", "주머니에 손 넣고 걷기 금지", "적재함·발판의 눈·얼음 먼저 제거"]],
      [fa.FaThermometerQuarter, COL.navy, "한랭질환 예방", ["방한복·방한장갑·귀마개, 여벌 양말", "젖은 옷·장갑은 즉시 교체", "1~2시간마다 따뜻한 곳에서 휴식", "따뜻한 물 자주 마시기(음주 금지)"]],
      [fa.FaFirstAid, COL.red, "이상 증상 시 즉시 조치", ["심한 떨림·말 어눌함 → 저체온증 의심", "손발 하얗고 감각 없음 → 동상 의심", "따뜻한 곳으로 옮기고 119 신고", "동상 부위 문지르지 않기"]],
    ];
    const w = 3.9, gap = 0.2;
    for (let i = 0; i < items.length; i++) {
      const [Ic, c, t, list] = items[i];
      const x = 0.6 + i * (w + gap);
      card(s, x, 1.45, w, 5.0, i === 2 ? "FDECEE" : COL.ice, t);
      await circleIcon(s, Ic, x + 0.3, 1.7, 0.85, c, t);
      T(s, t, { x: x + 1.3, y: 1.75, w: w - 1.4, h: 0.75, fontSize: 19, bold: true, color: COL.navy, valign: "middle" });
      T(s, bullets(list, { paraSpaceAfter: 12 }), { x: x + 0.3, y: 2.85, w: w - 0.5, h: 3.4, fontSize: 16 });
    }
    s.addNotes("[2분] 겨울철 넘어짐은 가장 흔한 재해입니다. 차량에서 내릴 때 뛰어내리지 말고, 두 손과 한 발 또는 두 발과 한 손 — 세 군데를 항상 잡고 디딥니다. 추위 속 장시간 작업 시 저체온증과 동상이 올 수 있습니다. 동료가 심하게 떨거나 말이 어눌해지면 저체온증을 의심하고 바로 따뜻한 곳으로 옮기고 119에 신고합니다. 동상 부위는 문지르면 조직이 더 손상됩니다. 2인 1조로 서로를 살펴 주세요.");
  }

  {
    const s = content("작업② 유해·위험요인과 개선대책");
    const rows = [
      [hdr("위험요인"), hdr("위험성"), hdr("개선대책 (현장 실천)")],
      [cell("도로 통제 중 차량 충돌·2차사고", { bold: true }), lv("높음", COL.red), cell("신호수 갓길 배치·대피공간 확보, 안전차량 후방 배치, 경고표지 거리 연장, 반사조끼·경광봉 필수")],
      [cell("빙판 운전 → 미끄러짐·추돌", { bold: true }), lv("높음", COL.red), cell("감속·차간거리 2배, 출동 전 체인·경광등 점검, 결빙 취약구간 사전 공유")],
      [cell("살포 중 저속 주행 → 후방 추돌", { bold: true }), lv("보통", COL.orange), cell("경광등·후방 화살표 상시 작동, 교통량 많은 구간은 VMS 안내·통제 병행")],
      [cell("살포기 막힘 처리 → 회전부 끼임", { bold: true }), lv("높음", COL.red), cell("동력 차단 후 작업(전원 OFF 표시), 손 대신 도구 사용, 회전부 덮개 유지")],
      [cell("제설제 적재 → 미끄러짐·요통", { bold: true }), lv("보통", COL.orange), cell("적재장 바닥 제설·모래, 포대 2인 1조·운반기구 사용, 작업 전 스트레칭")],
      [cell("야간 단독 출동 → 피로·한랭·발견 지연", { bold: true }), lv("보통", COL.orange), cell("2인 1조 출동, 무전·휴대전화 정시 보고, 연속근무 시간 제한·교대, 방한용품 지급")],
    ];
    s.addTable(rows, { x: 0.6, y: 1.4, w: 12.1, colW: [3.9, 1.1, 7.1], rowH: 0.74, border: { type: "solid", pt: 1, color: "C9D6E2" }, fontSize: 13, valign: "middle", margin: 0.08, objectName: "작업2 개선대책 표" });
    s.addNotes("[3분] 긴급출동 작업의 위험요인과 대책입니다. 위험성이 '높음'인 세 가지 — 도로 통제 중 차량 충돌, 빙판 운전, 살포기 끼임 — 은 반드시 지켜야 할 대책입니다. 특히 야간 출동은 혼자 나가지 않는 것, 정해진 시간마다 연락하는 것을 원칙으로 합시다. 이 표 역시 초안입니다. 다음 시간에 여러분 의견으로 고쳐서 확정하겠습니다.");
  }

  // ===== Ⅴ =====
  await sectionSlide("PART 5", "근로자 의견청취·정리", "여러분의 경험이 가장 좋은 개선대책입니다", fa.FaComments, 10,
    "[섹션 전환] 이제 여러분의 의견을 듣는 시간입니다.");

  {
    const s = content("위험성 판단 — 가능성 × 중대성 (3단계 판단법)");
    // matrix
    const gx = 2.0, gy = 1.85, cw = 1.55, ch = 1.15;
    const L = [["보통", "높음", "높음"], ["낮음", "보통", "높음"], ["낮음", "낮음", "보통"]];
    const colr = { "낮음": COL.green, "보통": COL.orange, "높음": COL.red };
    const yl = ["자주 (상)", "가끔 (중)", "드묾 (하)"], xl = ["경상 (소)", "휴업 (중)", "사망·중상 (대)"];
    T(s, "가능성", { x: 0.6, y: 1.4, w: 1.4, h: 0.4, fontSize: 14, bold: true, color: COL.navy, align: "center" });
    for (let r = 0; r < 3; r++) {
      T(s, yl[r], { x: 0.6, y: gy + r * ch, w: 1.3, h: ch, fontSize: 13, align: "center", valign: "middle" });
      for (let c = 0; c < 3; c++) {
        const v = L[r][c];
        s.addShape(pres.shapes.RECTANGLE, { x: gx + c * cw, y: gy + r * ch, w: cw - 0.06, h: ch - 0.06, fill: { color: colr[v] }, line: { type: "none" }, objectName: `매트릭스 ${r}${c}` });
        T(s, v, { x: gx + c * cw, y: gy + r * ch, w: cw - 0.06, h: ch - 0.06, fontSize: 18, bold: true, color: COL.white, align: "center", valign: "middle" });
      }
    }
    for (let c = 0; c < 3; c++) T(s, xl[c], { x: gx + c * cw, y: gy + 3 * ch + 0.05, w: cw - 0.06, h: 0.4, fontSize: 12, align: "center" });
    T(s, "중대성 (다치는 정도)", { x: gx, y: gy + 3 * ch + 0.45, w: 3 * cw, h: 0.4, fontSize: 14, bold: true, color: COL.navy, align: "center" });
    // legend + example
    const lg = [["높음", COL.red, "작업 전 즉시 개선 — 개선 전 작업 금지"], ["보통", COL.orange, "기한을 정해 개선, 그 전까지 관리대책"], ["낮음", COL.green, "현 상태 유지, 교육·점검으로 관리"]];
    let y = 1.5;
    for (const [t, c, d] of lg) {
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: 7.2, y, w: 1.0, h: 0.5, rectRadius: 0.08, fill: { color: c }, line: { type: "none" }, objectName: `범례 ${t}` });
      T(s, t, { x: 7.2, y, w: 1.0, h: 0.5, fontSize: 15, bold: true, color: COL.white, align: "center", valign: "middle" });
      T(s, d, { x: 8.4, y, w: 4.3, h: 0.5, fontSize: 14, valign: "middle" });
      y += 0.68;
    }
    card(s, 7.2, 3.75, 5.5, 2.7, COL.ice, "판단 예시");
    T(s, [
      { text: "판단 예시", options: { bold: true, color: COL.navy, fontSize: 17, breakLine: true } },
      { text: "적재함 위에서 톤백 고리 해체", options: { bold: true, breakLine: true } },
      { text: "가능성: 매번 출동 때 함 → 자주(상)", options: { breakLine: true } },
      { text: "중대성: 2m 추락 → 사망·중상(대)", options: { breakLine: true } },
      { text: "→ 높음: 지상 해체 방식·승강설비로 즉시 개선", options: { bold: true, color: COL.red } },
    ], { x: 7.45, y: 3.85, w: 5.1, h: 2.5, fontSize: 15, paraSpaceAfter: 4 });
    s.addNotes("[3분] 위험의 크기는 '얼마나 자주 일어날 수 있나(가능성)'와 '일어나면 얼마나 크게 다치나(중대성)'로 판단합니다. 2023년 개정 지침에 따라 복잡한 점수 계산 대신 높음·보통·낮음 3단계로 간단히 판단해도 됩니다. 예를 들어 적재함 위에서 고리를 푸는 일은 거의 매번 하고, 떨어지면 크게 다치니 '높음' — 바로 개선해야 합니다. 이제 이 기준으로 여러분이 직접 판단해 보겠습니다.");
  }

  {
    const s = content("근로자 의견청취 — 여러분의 생각을 들려주세요");
    const qs = [
      [fa.FaExclamationTriangle, COL.red, "아차사고", "지난 겨울 제설작업 중 \"큰일 날 뻔했다\" 싶었던 순간은?"],
      [gi.GiCrane, COL.orange, "상차 작업", "크레인 상차 때 가장 불편하거나 위험한 동작은?"],
      [fa.FaRoad, COL.blue, "도로 작업", "신호수·차량통제에서 부족한 장비나 인원은?"],
      [fa.FaLightbulb, COL.green, "개선 아이디어", "당장 바꾸면 좋을 것 한 가지는?"],
    ];
    for (let i = 0; i < qs.length; i++) {
      const [Ic, c, t, q] = qs[i];
      const x = 0.6 + (i % 2) * 3.15, y = 1.45 + Math.floor(i / 2) * 2.55;
      card(s, x, y, 3.0, 2.4, COL.ice, `질문 ${t}`);
      await circleIcon(s, Ic, x + 0.25, y + 0.25, 0.7, c, t);
      T(s, `Q${i + 1}. ${t}`, { x: x + 1.05, y: y + 0.3, w: 1.85, h: 0.6, fontSize: 16, bold: true, color: COL.navy, valign: "middle" });
      T(s, q, { x: x + 0.25, y: y + 1.1, w: 2.55, h: 1.2, fontSize: 14 });
    }
    T(s, "의견 기록표", { x: 7.0, y: 1.4, w: 5.7, h: 0.45, fontSize: 18, bold: true, color: COL.navy });
    const blank = () => [cell(" "), cell(" "), cell(" "), cell(" ")];
    s.addTable([
      [hdr("작업"), hdr("위험요인"), hdr("개선의견"), hdr("판단")],
      [cell("예) 상차", { color: COL.gray, italic: true }), cell("톤백 고리 풀려고 적재함 승강", { color: COL.gray, italic: true }), cell("긴 갈고리봉으로 지상 해체", { color: COL.gray, italic: true }), cell("높음", { color: COL.red, bold: true, align: "center" })],
      blank(), blank(), blank(), blank(),
    ], { x: 7.0, y: 1.9, w: 5.7, colW: [1.0, 1.85, 1.95, 0.9], rowH: 0.78, border: { type: "solid", pt: 1, color: "C9D6E2" }, fontSize: 12, valign: "middle", margin: 0.06, objectName: "의견 기록표" });
    s.addNotes("[6분] 네 가지 질문을 드리겠습니다. 정답은 없습니다. 작은 것이라도 좋습니다. 진행 팁: ① 2~3명씩 짝을 지어 2분간 이야기 → ② 조별로 한 가지씩 발표 → ③ 강사가 오른쪽 기록표(또는 화이트보드)에 적고, 함께 높음·보통·낮음을 판단합니다. 나온 의견은 위험성평가표에 반영하고, 개선 결과를 다음 TBM 때 알려드리겠습니다. 의견을 낸 근로자에게 개선 결과를 반드시 피드백하는 것이 참여를 이어가는 핵심입니다. (작성한 기록표는 위험성평가 근로자 참여 증빙으로 보관)");
  }

  {
    const s = content("출동 전 5분 TBM 체크리스트");
    const left = ["오늘 작업 내용·구간·담당 확인", "기상(기온·강설·결빙 예보) 공유", "2인 1조 편성, 비상연락·보고 시간 확인", "안전모·반사조끼·미끄럼방지 안전화·방한장구", "경광봉·깃발·라바콘·경고표지 수량"];
    const right = ["크레인: 슬링·훅·아웃트리거 점검", "차량: 체인·경광등·제동·살포기 점검", "결빙 취약구간·대피 공간 공유", "몸 상태 확인(음주·피로·감기약 복용)", "\"위험하면 멈춘다\" 작업중지 원칙 확인"];
    for (const [arr, x, t, c] of [[left, 0.6, "사람·장비", COL.navy], [right, 6.85, "작업·현장", COL.orange]]) {
      card(s, x, 1.45, 5.9, 4.95, COL.ice, `TBM ${t}`);
      T(s, t, { x: x + 0.35, y: 1.65, w: 5, h: 0.45, fontSize: 19, bold: true, color: c });
      let y = 2.3;
      for (const it of arr) {
        s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x: x + 0.35, y: y + 0.04, w: 0.4, h: 0.4, rectRadius: 0.06, fill: { color: COL.white }, line: { color: c, width: 1.5 }, objectName: "체크박스" });
        T(s, it, { x: x + 0.95, y, w: 4.75, h: 0.5, fontSize: 15, valign: "middle" });
        y += 0.78;
      }
    }
    s.addNotes("[1분] 출동 전 5분, 이 체크리스트로 TBM을 합니다. 급한 출동일수록 이 5분이 사고를 막습니다. 차량 대시보드나 대기실에 붙여두고 사용하시면 좋습니다.");
  }

  // ===== closing =====
  section = "마무리";
  pres.addSection({ title: section });
  {
    const s = pres.addSlide({ masterName: "SECTION", sectionTitle: section });
    T(s, "꼭 기억할 제설작업 안전수칙 6", { x: 0.7, y: 0.5, w: 12, h: 0.8, fontSize: 32, bold: true, color: COL.white });
    const rules = [
      [fa.FaTruck, "적재함 위에 올라가지 않는다", "고리 해체는 지상에서, 부득이하면 승강설비"],
      [gi.GiCrane, "매달린 짐 아래 들어가지 않는다", "작업반경 출입금지, 신호수는 1명"],
      [fa.FaCarCrash, "움직이는 차량 앞·뒤에 서지 않는다", "유도자 배치, 안전차량은 작업자 뒤쪽"],
      [gi.GiTrafficCone, "작업구간을 먼저 만들고 작업한다", "경고표지 → 라바콘 → 작업, 철수는 역순"],
      [fa.FaSnowflake, "빙판에선 속도 절반, 거리 두 배", "출동 전 체인·경광등 점검"],
      [fa.FaHandPaper, "위험하면 멈추고 말한다", "작업중지권 · 아차사고 공유"],
    ];
    for (let i = 0; i < rules.length; i++) {
      const [Ic, t, d] = rules[i];
      const x = 0.7 + (i % 2) * 6.1, y = 1.6 + Math.floor(i / 2) * 1.8;
      s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w: 5.85, h: 1.55, rectRadius: 0.12, fill: { color: COL.white, transparency: 90 }, line: { type: "none" }, objectName: `수칙 ${i + 1}` });
      await circleIcon(s, Ic, x + 0.25, y + 0.3, 0.95, COL.orange, t);
      T(s, t, { x: x + 1.4, y: y + 0.25, w: 4.3, h: 0.5, fontSize: 18, bold: true, color: COL.white, valign: "middle" });
      T(s, d, { x: x + 1.4, y: y + 0.8, w: 4.3, h: 0.5, fontSize: 14, color: COL.ice, valign: "middle" });
    }
    s.addNotes("[1분] 오늘 내용을 여섯 줄로 정리합니다. 함께 한 번 읽어 보겠습니다. 적재함 위에 올라가지 않는다, 매달린 짐 아래 들어가지 않는다, 움직이는 차량 앞뒤에 서지 않는다, 작업구간을 먼저 만들고 작업한다, 빙판에선 속도 절반 거리 두 배, 위험하면 멈추고 말한다.");
  }

  {
    const s = content("참고자료");
    T(s, bullets([
      "산업안전보건법 제36조(위험성평가), 제52조(작업중지), 제63·64조(도급인의 안전조치)",
      "고용노동부 「사업장 위험성평가에 관한 지침」(2023 개정) · 국토교통부 「도로공사장 교통관리지침」",
      "뉴데일리·뉴시스 2024.11.28 — 첫 폭설 이틀간 제설작업 중 사망 및 도로통제 작업자 사망",
      "뉴시스 2024.10.21 — 고속도로 안전순찰원 2차사고(6년간 1명 사망·16명 부상)",
      "경인일보 — 고속도로 2차사고 2020년 51건 → 2024년 70건 (한국도로공사 국회 제출자료)",
      "CBS노컷뉴스 2026.6.22 — 화물차 적재함 추락, 떨어짐 사망 단일 기인물 1위",
      "안전보건공단 재해사례 — 철근다발 하역작업 중 추락(2023.12 청주)",
      "SBS Biz·YTN·경북매일 2026.1~2 — 서산영덕고속도로 다중추돌 및 국토부 감사 결과",
      "뉴시스 2024.1.3 · 서울신문 2022.10.19 — 결빙 노면 사고 치사율(한국교통안전공단)",
    ], { paraSpaceAfter: 10 }), { x: 0.6, y: 1.45, w: 12.1, h: 5.2, fontSize: 15 });
    s.addNotes("교육에 인용한 사고사례와 통계의 출처입니다. 수치는 기관·기간에 따라 차이가 있으므로 공식 문서에 인용할 때는 원문을 확인하십시오.");
  }

  {
    const s = pres.addSlide({ masterName: "TITLE", sectionTitle: section });
    s.addImage({ path: img.cover, x: 0, y: 0, w: 13.333, h: 7.5, altText: "눈 내리는 밤 도로 위 제설차와 신호수 그림", objectName: "마무리 그림" });
    s.addShape(pres.shapes.RECTANGLE, { x: 0, y: 0, w: 8.8, h: 4.9, fill: { color: COL.navy, transparency: 20 }, line: { type: "none" }, objectName: "마무리 배경" });
    s.addText("감사합니다", { placeholder: "title" });
    s.addText("올겨울도 무사고로 — 위험하면 멈추고, 함께 말해 주세요", { placeholder: "body" });
    s.addNotes("질의응답 후 교육 참석자 서명, 의견 기록표 회수로 마무리합니다.");
  }

  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})().catch((e) => { console.error(e); process.exit(1); });
