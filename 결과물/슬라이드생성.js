const pptx = require("pptxgenjs");
const P = new pptx();
P.layout = "LAYOUT_WIDE";
const W = 13.333, H = 7.5;

const BG="EFEFF7", DARK="231D4F", CARD="FFFFFF", INK="1F1F2E", MUTED="5A5A66", FAINT="A9A9B8";
const C1="5B4FCF", C2="B5650A", C3="0E8C7A";
const T1="EFEEFB", T2="FAF2E8", T3="E6F3F0", TN="F4F4F9";
const F="Malgun Gothic";
const IMG="C:/Users/이성진/OneDrive - KookminUNIV/바탕 화면/공모전/code/data/그림_결과물/";
const FOOT="제1회 AI금융빅데이터플랫폼 소비데이터 공모전 · 팀 유비온";

const NAV = ["문제와 방법", "핵심 결과", "실행과 검증"];
function frame(s, active, title, sub){
  s.background = { color: BG };
  NAV.forEach((t,i)=>{
    s.addText(t, { x:0.85+i*2.35, y:0.2, w:2.3, h:0.32, isTextBox:true, margin:0,
      fontFace:F, fontSize:11, bold:i===active, color:i===active?C1:FAINT, align:"left" });
  });
  s.addShape(P.ShapeType.roundRect, { x:0.42, y:0.72, w:W-0.84, h:H-1.14, rectRadius:0.06,
    fill:{color:CARD}, line:{color:"E3E3EE", width:1},
    shadow:{ type:"outer", color:"9A9AB0", blur:14, offset:2, angle:90, opacity:0.18 } });
  s.addText(title, { x:0.85, y:1.0, w:W-1.7, h:0.6, isTextBox:true, margin:0,
    fontFace:F, fontSize:26, bold:true, color:INK, align:"left" });
  if(sub) s.addText(sub, { x:0.85, y:1.62, w:W-1.7, h:0.34, isTextBox:true, margin:0,
    fontFace:F, fontSize:13, color:MUTED, align:"left" });
  s.addText(FOOT, { x:0.42, y:H-0.38, w:W-0.84, h:0.24, isTextBox:true, margin:0,
    fontFace:F, fontSize:9, color:FAINT, align:"center" });
}
function concl(s, y, bold, rest, tint, border){
  s.addShape(P.ShapeType.roundRect, { x:0.85, y, w:W-1.7, h:0.62, rectRadius:0.08,
    fill:{color:tint||TN}, line:{color:border||"E3E3EE", width:1} });
  s.addText([{text:bold+" ", options:{bold:true, color:INK}}, {text:rest, options:{color:MUTED}}],
    { x:1.1, y, w:W-2.2, h:0.62, isTextBox:true, margin:0, fontFace:F, fontSize:12, valign:"middle" });
}
function statRow(s, items, y){
  items.forEach((d,i)=>{
    const x = 0.85 + i*2.96, w = 2.74;
    s.addShape(P.ShapeType.roundRect, { x, y, w, h:1.32, rectRadius:0.07,
      fill:{color:"FBFBFD"}, line:{color:"E9E9F1", width:1} });
    s.addText(d[0], { x:x+0.18, y:y+0.08, w:w-0.36, h:0.58, isTextBox:true, margin:0,
      fontFace:F, fontSize:29, bold:true, color:d[3], valign:"middle" });
    s.addText(d[1], { x:x+0.18, y:y+0.66, w:w-0.36, h:0.26, isTextBox:true, margin:0,
      fontFace:F, fontSize:11, bold:true, color:INK });
    s.addText(d[2], { x:x+0.18, y:y+0.91, w:w-0.36, h:0.36, isTextBox:true, margin:0,
      fontFace:F, fontSize:9.5, color:MUTED, lineSpacingMultiple:1.05 });
  });
}
function deepDive(s, title, sub, stats, img, cb, cr, tint, border, notes){
  frame(s, 2, title, sub);
  statRow(s, stats, 2.12);
  s.addImage({ path: IMG+img, x:1.52, y:3.62, w:10.3, h:2.61 });
  concl(s, 6.44, cb, cr, tint, border);
  s.addNotes(notes);
}

function frameApx(s, title, sub){
  s.background = { color: BG };
  s.addShape(P.ShapeType.roundRect, { x:0.85, y:0.2, w:1.1, h:0.34, rectRadius:0.05,
    fill:{color:"E3E1F2"} });
  s.addText("부록", { x:0.85, y:0.2, w:1.1, h:0.34, isTextBox:true, margin:0,
    fontFace:F, fontSize:10.5, bold:true, color:C1, align:"center", valign:"middle" });
  s.addShape(P.ShapeType.roundRect, { x:0.42, y:0.72, w:W-0.84, h:H-1.14, rectRadius:0.06,
    fill:{color:CARD}, line:{color:"E3E3EE", width:1},
    shadow:{ type:"outer", color:"9A9AB0", blur:14, offset:2, angle:90, opacity:0.18 } });
  s.addText(title, { x:0.85, y:1.0, w:W-1.7, h:0.55, isTextBox:true, margin:0,
    fontFace:F, fontSize:24, bold:true, color:INK });
  if(sub) s.addText(sub, { x:0.85, y:1.58, w:W-1.7, h:0.34, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, color:MUTED });
  s.addText(FOOT, { x:0.42, y:H-0.38, w:W-0.84, h:0.24, isTextBox:true, margin:0,
    fontFace:F, fontSize:9, color:FAINT, align:"center" });
}
function 표(s, head, rows, opt){
  const H0 = head.map(h=>({text:h[0], options:{bold:true, color:"FFFFFF",
    fill:{color:opt.head||C1}, align:h[1]||"left"}}));
  const B = rows.map(r=>r.map((c,i)=>({text:c,
    options:{color:(opt.bold&&opt.bold.includes(i))?INK:MUTED,
             bold:!!(opt.bold&&opt.bold.includes(i)), align:head[i][1]||"left"}})));
  s.addTable([H0].concat(B), { x:opt.x, y:opt.y, w:opt.w, colW:opt.colW,
    rowH:opt.rowH||0.38, fontFace:F, fontSize:opt.fs||10.5,
    border:{type:"solid", color:"ECECF3", pt:0.5}, fill:{color:"FFFFFF"}, valign:"middle" });
}

/* ═══ 1. 표지 ═══ */
{
const s = P.addSlide();
s.background = { color: DARK };
s.addText("BC카드 저침투 시장 공략", { x:1.2, y:2.35, w:11, h:0.95, isTextBox:true, margin:0,
  fontFace:F, fontSize:44, bold:true, color:"FFFFFF" });
s.addText("점포·인구 기반 소비 격차 분석을 통한 지역별 공략 우선순위 도출",
  { x:1.2, y:3.4, w:11, h:0.5, isTextBox:true, margin:0, fontFace:F, fontSize:19, color:"C9C4E8" });
s.addShape(P.ShapeType.roundRect, { x:1.2, y:4.35, w:6.6, h:0.62, rectRadius:0.08,
  fill:{color:"342C68"} });
s.addText("1,813개 시장 분석  →  12개 핵심 후보  →  3개 시범지역",
  { x:1.45, y:4.35, w:6.2, h:0.62, isTextBox:true, margin:0, fontFace:F, fontSize:14,
    bold:true, color:"FFFFFF", valign:"middle" });
s.addText(FOOT, { x:1.2, y:6.42, w:11, h:0.3, isTextBox:true, margin:0,
  fontFace:F, fontSize:11, color:"8F88BE" });
s.addText([{text:"팀장 ", options:{color:"8F88BE"}}, {text:"이성진", options:{bold:true, color:"FFFFFF"}},
  {text:"   ·   팀원 ", options:{color:"8F88BE"}}, {text:"염신호 · 이규상", options:{bold:true, color:"FFFFFF"}}],
  { x:1.2, y:6.76, w:11, h:0.32, isTextBox:true, margin:0, fontFace:F, fontSize:13 });
}

/* 101 목차 */
{
const s = P.addSlide();
s.background = { color: BG };
s.addShape(P.ShapeType.roundRect, { x:0.42, y:0.72, w:W-0.84, h:H-1.14, rectRadius:0.06,
  fill:{color:CARD}, line:{color:"E3E3EE", width:1},
  shadow:{ type:"outer", color:"9A9AB0", blur:14, offset:2, angle:90, opacity:0.18 } });
s.addText("목차", { x:0.95, y:1.05, w:4, h:0.5, isTextBox:true, margin:0,
  fontFace:F, fontSize:24, bold:true, color:INK });
const sec = [
  ["01","연구 개요", ["1. 연구 배경","2. 문제 정의","3. 연구 질문","4. 분석 프로세스","5. 연구 요약"], C1],
  ["02","데이터 구축 및 분석", ["1. 데이터 및 분석 범위","2. 1,813개 → 12개 후보","3. 최종 12개 후보","4. 부족분의 구성 분해","5. 기회 유형 3가지"], C2],
  ["03","결론 및 제언", ["1. 유형별 처방과 확장 규칙","2. 시범지역 3곳 Deep Dive","3. 성과 측정과 KPI","4. 분석의 신뢰성","5. 한계와 최종 제안"], C3],
];
sec.forEach((d,i)=>{
  const x = 0.95 + i*3.85, w = 3.55;
  s.addShape(P.ShapeType.roundRect, { x, y:2.0, w, h:4.2, rectRadius:0.08,
    fill:{color:"FBFBFD"}, line:{color:d[3], width:1.2} });
  s.addText(d[0], { x:x+0.3, y:2.25, w:1.0, h:0.62, isTextBox:true, margin:0,
    fontFace:F, fontSize:30, bold:true, color:d[3], valign:"middle" });
  s.addText(d[1], { x:x+0.3, y:2.92, w:w-0.6, h:0.42, isTextBox:true, margin:0,
    fontFace:F, fontSize:17, bold:true, color:INK, valign:"middle" });
  d[2].forEach((t,k)=>{
    s.addText(t, { x:x+0.3, y:3.55+k*0.48, w:w-0.6, h:0.4, isTextBox:true, margin:0,
      fontFace:F, fontSize:12, color:MUTED, valign:"middle" });
  });
});
s.addText(FOOT, { x:0.42, y:H-0.38, w:W-0.84, h:0.24, isTextBox:true, margin:0,
  fontFace:F, fontSize:9, color:FAINT, align:"center" });
}

/* 102 연구 배경 */
{
const s = P.addSlide();
frame(s, 0, "BC카드 이용 수준은 전국에서 고르지 않다",
  "마케팅 자원은 한정되어 있고, 모든 시장을 같은 비중으로 공략할 수는 없다");
s.addImage({ path: IMG+"03_연구배경.png", x:1.52, y:2.15, w:10.3, h:2.61 });
const pt = [["2.5배","시도 간 격차","침투지수 중앙값 최고 ÷ 최저",C1],
            ["1,008개","예상에 못 미치는 시장","전체 1,813개의 55.6%",C2],
            ["한정된 예산","공략 우선순위가 필요하다","어디부터 손댈 것인가",C3]];
pt.forEach((d,i)=>{
  const x = 0.95 + i*3.85, w = 3.55;
  s.addShape(P.ShapeType.roundRect, { x, y:5.0, w, h:1.2, rectRadius:0.07,
    fill:{color:"FBFBFD"}, line:{color:d[3], width:1.2} });
  s.addText(d[0], { x:x+0.22, y:5.1, w:w-0.44, h:0.48, isTextBox:true, margin:0,
    fontFace:F, fontSize:22, bold:true, color:d[3], valign:"middle" });
  s.addText(d[1], { x:x+0.22, y:5.58, w:w-0.44, h:0.28, isTextBox:true, margin:0,
    fontFace:F, fontSize:11.5, bold:true, color:INK });
  s.addText(d[2], { x:x+0.22, y:5.86, w:w-0.44, h:0.28, isTextBox:true, margin:0,
    fontFace:F, fontSize:10, color:MUTED });
});
concl(s, 6.4, "전국을 고르게 공략하는 것은 선택지가 아니다.",
  "어디가 상대적으로 약한지를 먼저 가려내야 한정된 자원을 배분할 수 있다.");
s.addNotes("시도별 중앙값 수치는 팩트시트 6절과 소수점 차이가 있어 본문에는 배수만 쓴다. 외부 인용 없이 우리 데이터로만 배경을 세웠다.");
}

/* ═══ 3. 문제 정의 ═══ */
{
const s = P.addSlide();
frame(s, 0, "단순 비교로는 시장이 작은 건지 BC가 약한 건지 알 수 없다",
  "기존 지표는 상주인구로 나누기 때문에 유동인구가 많은 지역이 실제보다 강해 보인다");
s.addShape(P.ShapeType.roundRect, { x:0.85, y:2.12, w:6.0, h:2.5, rectRadius:0.08,
  fill:{color:"FBFBFD"}, line:{color:"E9E9F1", width:1} });
s.addText("기존 접근 — 1인당 BC 소비", { x:1.1, y:2.28, w:5.5, h:0.34, isTextBox:true, margin:0,
  fontFace:F, fontSize:13, bold:true, color:INK });
s.addShape(P.ShapeType.roundRect, { x:1.1, y:2.62, w:5.5, h:0.38, rectRadius:0.05,
  fill:{color:"F2F2F7"} });
s.addText("1인당 BC 소비  =  그 지역 BC 결제액  ÷  상주인구", { x:1.25, y:2.62, w:5.2, h:0.38,
  isTextBox:true, margin:0, fontFace:F, fontSize:11.5, bold:true, color:INK, valign:"middle" });
s.addText("분자에는 그 지역에서 일하고 소비하는 사람까지 들어가는데 분모는 사는 사람뿐이라, 도심일수록 부풀려진다.",
  { x:1.1, y:3.06, w:5.5, h:0.5, isTextBox:true, margin:0, fontFace:F, fontSize:11,
    color:MUTED, lineSpacingMultiple:1.1 });
[["부산 중구","607","도심 · 유동인구 집중",C2],["용인 기흥구","60","주거지 · 상주인구 다수",C1]].forEach((d,i)=>{
  const y = 3.58 + i*0.52;
  s.addText(d[0], { x:1.1, y, w:1.9, h:0.5, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, bold:true, color:INK, valign:"middle" });
  s.addText(d[1], { x:3.0, y, w:0.9, h:0.5, isTextBox:true, margin:0,
    fontFace:F, fontSize:20, bold:true, color:d[3], align:"right", valign:"middle" });
  s.addText(d[2], { x:4.05, y, w:2.5, h:0.5, isTextBox:true, margin:0,
    fontFace:F, fontSize:10.5, color:MUTED, valign:"middle" });
});
s.addShape(P.ShapeType.roundRect, { x:7.15, y:2.12, w:5.28, h:2.5, rectRadius:0.08,
  fill:{color:T1}, line:{color:C1, width:1.2} });
s.addText("그래서 질문을 바꿨다", { x:7.4, y:2.28, w:4.8, h:0.34, isTextBox:true, margin:0,
  fontFace:F, fontSize:13, bold:true, color:C1 });
s.addText("이 정도 상권이면\nBC카드 결제가 얼마쯤\n나와야 하는가?",
  { x:7.4, y:2.72, w:4.8, h:1.15, isTextBox:true, margin:0, fontFace:F, fontSize:19,
    bold:true, color:INK, lineSpacingMultiple:1.12 });
s.addText("→ 실제값과 비교해, 예상보다 지속적으로 낮은 시장을 찾는다",
  { x:7.4, y:3.98, w:4.8, h:0.42, isTextBox:true, margin:0, fontFace:F, fontSize:12, color:MUTED });
concl(s, 4.9, "시장 규모를 통제한 뒤에야 'BC가 약한 곳'이 보인다.",
  "점포 수와 인구를 모델에 넣어 기대 수준을 만들고, 거기서 벗어난 만큼을 격차로 측정했다.");
s.addText("※ 1인당 소비 지표는 최종 판정에서 제외했으나, 그 검증 과정을 부록에 남겼다",
  { x:0.85, y:5.7, w:W-1.7, h:0.3, isTextBox:true, margin:0, fontFace:F, fontSize:10, color:FAINT });
s.addNotes("유동인구 편향: 부산 중구 607 vs 용인 기흥 60. 침투도가 아니라 주간활동인구를 재고 있었다는 발견.");
}

/* 103 연구 질문 */
{
const s = P.addSlide();
frame(s, 0, "세 가지 질문으로 좁혔다",
  "'어디가 약한가'에서 멈추지 않고 '무엇이 부족한가', '그래서 무엇을 할 것인가'까지 간다");
const q = [
  ["01","어디가 약한가", "\"이 정도 상권이면\nBC 결제가 얼마나 나와야 하나?\"",
   "점포 수와 인구로 기대 수준을 만들고\n거기서 벗어난 시장을 찾는다", C1],
  ["02","무엇이 부족한가", "\"같은 저침투라도\n구성이 서로 다르지 않은가?\"",
   "결제 건수 · 건당 금액 · 연령 · 카테고리로 분해해\n돈이 지금 어디 있는지를 가른다", C2],
  ["03","무엇을 할 것인가", "\"유형별로 무엇을,\n어느 수준까지 하면 효과적인가?\"",
   "유형마다 처방을 다르게 붙이고,\n성공·실패 판정 기준을 미리 정한다", C3],
];
q.forEach((d,i)=>{
  const x = 0.95 + i*3.85, w = 3.55;
  s.addShape(P.ShapeType.roundRect, { x, y:2.15, w, h:3.55, rectRadius:0.08,
    fill:{color:"FBFBFD"}, line:{color:d[5], width:1.2} });
  s.addText(d[0], { x:x+0.28, y:2.38, w:0.7, h:0.44, isTextBox:true, margin:0,
    fontFace:F, fontSize:22, bold:true, color:d[5], valign:"middle" });
  s.addText(d[1], { x:x+1.0, y:2.38, w:w-1.3, h:0.44, isTextBox:true, margin:0,
    fontFace:F, fontSize:15, bold:true, color:INK, valign:"middle" });
  s.addShape(P.ShapeType.roundRect, { x:x+0.25, y:2.98, w:w-0.5, h:1.3, rectRadius:0.06,
    fill:{color: i===0?T1 : i===1?T2 : T3} });
  s.addText(d[2], { x:x+0.42, y:2.98, w:w-0.84, h:1.3, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, bold:true, color:INK, valign:"middle", lineSpacingMultiple:1.15 });
  s.addText(d[3], { x:x+0.28, y:4.45, w:w-0.56, h:1.0, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, color:MUTED, lineSpacingMultiple:1.18 });
});
concl(s, 5.95, "'어디가'에서 나아가 '어떻게'와 '무엇을'까지 답하는 것이 이 분석의 목표다.",
  "세 질문이 각각 02장과 03장의 구성을 그대로 결정한다.");
s.addNotes("수상작의 연구 질문 페이지 구조를 따랐다. 세 질문이 이후 목차와 1:1로 대응한다.");
}

/* ═══ 5. 분석 프로세스 ═══ */
{
const s = P.addSlide();
frame(s, 0, "분석 프로세스 — 다섯 단계",
  "예상 수준을 만들고, 벗어난 곳을 찾고, 왜 벗어났는지 분해한 뒤, 실행 단위를 정한다");
const step = [
  ["1","예상 소비액 산출","점포 수 + 20대 이상 인구로\n업종별 회귀모형 적합",C1],
  ["2","침투지수 계산","실제 ÷ 예상 × 100\n격차금액 = 예상 − 실제",C1],
  ["3","후보 압축","침투 < 100\n→ 업종 내 상위 50%\n→ σ ≤ −1.5\n→ 격차 ≥ 20억",C2],
  ["4","원인 분해","결제 건수 / 건당 금액\n연령 · 지역 · 카테고리",C2],
  ["5","실행지역 선정","유형별 대표 1곳\n대전 서구 · 진주 · 서초",C3],
];
step.forEach((d,i)=>{
  const x = 0.85 + i*2.37, w = 2.05;
  s.addShape(P.ShapeType.roundRect, { x, y:2.35, w, h:2.75, rectRadius:0.08,
    fill:{color:"FBFBFD"}, line:{color:d[3], width:1.2} });
  s.addShape(P.ShapeType.ellipse, { x:x+w/2-0.26, y:2.58, w:0.52, h:0.52, fill:{color:d[3]} });
  s.addText(d[0], { x:x+w/2-0.26, y:2.58, w:0.52, h:0.52, isTextBox:true, margin:0,
    fontFace:F, fontSize:17, bold:true, color:"FFFFFF", align:"center", valign:"middle" });
  s.addText(d[1], { x:x+0.12, y:3.24, w:w-0.24, h:0.54, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, bold:true, color:INK, align:"center", valign:"middle" });
  s.addText(d[2], { x:x+0.12, y:3.84, w:w-0.24, h:1.1, isTextBox:true, margin:0,
    fontFace:F, fontSize:10, color:MUTED, align:"center", lineSpacingMultiple:1.15 });
  if(i<4) s.addText("→", { x:x+w+0.02, y:3.5, w:0.3, h:0.4, isTextBox:true, margin:0,
    fontFace:F, fontSize:16, bold:true, color:FAINT, align:"center", valign:"middle" });
});
concl(s, 5.5, "복잡한 계산은 3단계에 모여 있다.",
  "1~2단계로 기준선을 만들고, 3단계에서 네 조건을 모두 통과한 시장만 남긴다. 점수를 매겨 줄 세우지 않는다.");
s.addNotes("가중치를 정하지 않는다는 점이 핵심. 네 조건을 독립적으로 통과해야 후보가 된다.");
}

/* ═══ 2. Executive Summary ═══ */
{
const s = P.addSlide();
frame(s, 0, "10개 지역 12개 시장에서 예상 대비 큰 격차를 발견했다",
  "점포 수와 인구로 기대되는 BC 결제액을 산출하고, 실제와의 차이가 모델 오차를 넘어서는 시장을 찾았다");
const big = [["1,813개","분석 시장","227 시군구 × 8 업종",C1],
             ["12개","최종 후보","10개 시군구",C1],
             ["1,059억","모델 기준 소비 격차","회수 가능 매출이 아니다",C2],
             ["3곳","1차 시범지역","유형별 대표 1곳씩",C3]];
big.forEach((d,i)=>{
  const x = 0.85 + i*2.96, w = 2.74;
  s.addShape(P.ShapeType.roundRect, { x, y:2.18, w, h:1.62, rectRadius:0.08,
    fill:{color:i===2?T2:T1}, line:{color:d[3], width:1.2} });
  s.addText(d[0], { x:x+0.2, y:2.3, w:w-0.4, h:0.72, isTextBox:true, margin:0,
    fontFace:F, fontSize:36, bold:true, color:d[3], valign:"middle" });
  s.addText(d[1], { x:x+0.2, y:3.02, w:w-0.4, h:0.3, isTextBox:true, margin:0,
    fontFace:F, fontSize:12, bold:true, color:INK });
  s.addText(d[2], { x:x+0.2, y:3.32, w:w-0.4, h:0.3, isTextBox:true, margin:0,
    fontFace:F, fontSize:10, color:MUTED });
});
const qa = [["① 어디가 약한가?","10개 지역의 12개 시장 — 대전 서구 298억 · 익산 164억 · 진주 130억"],
            ["② 무엇이 부족한가?","건당 금액은 88~131로 정상이고 결제 건수가 26~69로 낮다 — 차이는 빈도에서 나온다"],
            ["③ 어떻게 공략하나?","돈이 지금 어디 있는지에 따라 지역 전체 · 업종 간 이동 · 채널 이동으로 나눠 처방한다"]];
qa.forEach((d,i)=>{
  const y = 4.18 + i*0.72;
  s.addShape(P.ShapeType.roundRect, { x:0.85, y, w:W-1.7, h:0.6, rectRadius:0.06,
    fill:{color:i%2===0?"FBFBFD":"FFFFFF"}, line:{color:"ECECF3", width:0.8} });
  s.addText(d[0], { x:1.1, y, w:2.5, h:0.6, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, bold:true, color:C1, valign:"middle" });
  s.addText(d[1], { x:3.5, y, w:W-4.6, h:0.6, isTextBox:true, margin:0,
    fontFace:F, fontSize:12, color:INK, valign:"middle" });
});
s.addText("1,059억은 모델이 예상한 소비와 실제 소비의 차이이며, 곧바로 회수 가능한 매출이 아니다.",
  { x:0.85, y:6.42, w:W-1.7, h:0.3, isTextBox:true, margin:0, fontFace:F, fontSize:10.5, color:MUTED });
s.addNotes("첫 2쪽에서 프로젝트 전체가 이해되도록 구성. 1,059억의 라벨을 반드시 '격차'로 유지한다.");
}

/* ═══ 4. 데이터 및 분석 범위 ═══ */
{
const s = P.addSlide();
frame(s, 1, "227개 시군구 × 8개 업종 = 1,813개 시장을 전부 계산했다",
  "일부를 골라 본 것이 아니라 전국을 계산한 뒤 그 안에서 후보를 선별했다");
const dat = [["BC카드 소비데이터","2026.01 ~ 06","결제금액 · 건수 · 성별 · 연령",C1],
             ["주민등록인구","행정안전부","20대 이상 성인인구",C3],
             ["상가정보 API","소상공인시장진흥공단","업종별 점포 수",C2]];
dat.forEach((d,i)=>{
  const x = 0.85 + i*3.95, w = 3.7;
  s.addShape(P.ShapeType.roundRect, { x, y:2.12, w, h:1.42, rectRadius:0.08,
    fill:{color:"FBFBFD"}, line:{color:d[3], width:1.2} });
  s.addText(d[0], { x:x+0.22, y:2.26, w:w-0.44, h:0.36, isTextBox:true, margin:0,
    fontFace:F, fontSize:14, bold:true, color:d[3], valign:"middle" });
  s.addText(d[1], { x:x+0.22, y:2.64, w:w-0.44, h:0.3, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, bold:true, color:INK });
  s.addText(d[2], { x:x+0.22, y:2.94, w:w-0.44, h:0.36, isTextBox:true, margin:0,
    fontFace:F, fontSize:10.5, color:MUTED });
});
s.addShape(P.ShapeType.roundRect, { x:0.85, y:3.8, w:W-1.7, h:1.5, rectRadius:0.08,
  fill:{color:T1}, line:{color:C1, width:1.2} });
s.addText("점포 수  +  20대 이상 인구", { x:1.2, y:3.98, w:4.3, h:0.5, isTextBox:true, margin:0,
  fontFace:F, fontSize:16, bold:true, color:INK, valign:"middle" });
s.addText("→", { x:5.6, y:3.98, w:0.5, h:0.5, isTextBox:true, margin:0,
  fontFace:F, fontSize:18, bold:true, color:C1, align:"center", valign:"middle" });
s.addText("예상 BC 소비액", { x:6.2, y:3.98, w:2.6, h:0.5, isTextBox:true, margin:0,
  fontFace:F, fontSize:16, bold:true, color:C1, valign:"middle" });
s.addText("실제 소비액  ÷  예상 소비액  ×  100  =  침투지수",
  { x:1.2, y:4.55, w:8.0, h:0.5, isTextBox:true, margin:0, fontFace:F, fontSize:16,
    bold:true, color:INK, valign:"middle" });
s.addText("100 = 예상만큼 결제됨\n50 = 예상의 절반만 결제됨",
  { x:9.5, y:3.98, w:2.9, h:1.1, isTextBox:true, margin:0, fontFace:F, fontSize:11.5,
    color:MUTED, valign:"middle", lineSpacingMultiple:1.15 });
[["분석 범위","8개 업종 · 227개 시군구 (인천 중구+동구 병합) · 내국인과 외국인 포함",MUTED],
 ["20대 이상만","미성년자는 카드 발급·소비의 주체가 아니다. 0~19세를 분모에 넣으면 학령인구가 많은 신도시가 불리해진다",C1],
 ["광주·전남 제외","상가정보 API에 두 지역이 미수록이라 분모인 점포 수를 구할 수 없다 (인구의 6.2%)",C2],
 ["법인 제외","법인카드는 소비 주체가 회사라 본사 소재지에 결제가 몰려 지역 소비 패턴을 왜곡한다",C3]]
  .forEach((d,i)=>{
  const y = 5.42 + i*0.36;
  s.addText(d[0], { x:0.85, y, w:1.65, h:0.32, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, bold:true, color:d[2], valign:"middle" });
  s.addText(d[1], { x:2.55, y, w:9.9, h:0.32, isTextBox:true, margin:0,
    fontFace:F, fontSize:10.5, color:MUTED, valign:"middle" });
});
s.addNotes("광주·전남은 상가정보 API 미수록으로 제외 — 인구의 6.2%. 한계에서 다시 언급한다.");
}

/* ═══ 6. 1,813 → 12 ═══ */
{
const s = P.addSlide();
frame(s, 1, "1,813개 시장 중 실제 공략 대상은 12개로 좁혔다",
  "점수를 합산해 순위를 낸 것이 아니라 네 조건을 모두 통과해야 남는다  ·  σ = 업종별 예측오차 대비 이탈 정도");
s.addImage({ path: IMG+"15_깔때기.png", x:0.85, y:2.25, w:6.4, h:3.20 });
const fn = [["1,813","전체 시장","",FAINT],
            ["1,008","예상보다 적게 결제","침투지수 < 100",MUTED],
            ["514","업종 내 큰 시장","시장규모 상위 50%",MUTED],
            ["16","예측오차 대비 큰 이탈","σ ≤ −1.5",C2],
            ["12","최종 공략 후보","격차 20억 이상",C1]];
fn.forEach((d,i)=>{
  const y = 2.22 + i*0.64;
  s.addText(d[0], { x:7.6, y, w:1.35, h:0.56, isTextBox:true, margin:0, fontFace:F,
    fontSize:i===4?26:19, bold:true, color:d[3], align:"right", valign:"middle" });
  s.addText(d[1], { x:9.1, y, w:2.2, h:0.56, isTextBox:true, margin:0,
    fontFace:F, fontSize:11.5, bold:i===4, color:INK, valign:"middle" });
  s.addText(d[2], { x:11.3, y, w:1.5, h:0.56, isTextBox:true, margin:0,
    fontFace:F, fontSize:10, color:MUTED, valign:"middle" });
});
s.addShape(P.ShapeType.roundRect, { x:7.6, y:5.6, w:4.75, h:0.72, rectRadius:0.07,
  fill:{color:T1}, line:{color:C1, width:1.2} });
s.addText([{text:"10개 시군구  ·  ", options:{color:INK, bold:true}},
           {text:"1,059억", options:{color:C1, bold:true, fontSize:17}}],
  { x:7.85, y:5.6, w:4.3, h:0.72, isTextBox:true, margin:0, fontFace:F, fontSize:13, valign:"middle" });
concl(s, 6.45, "12개는 저침투의 전량이 아니라 1순위다.",
  "기준을 σ −1.3까지 풀면 13곳이 더 들어온다. 실행 가능한 규모로 좁힌 결과다.");
s.addNotes("컷오프 민감도: σ -1.3~-1.5 구간에서 12개가 전부 유지된다. 조이면 줄어든다는 점도 정직하게 적는다.");
}

/* ═══ 7. 최종 12개 후보 ═══ */
{
const s = P.addSlide();
frame(s, 1, "12개 후보는 10개 시군구에 흩어져 있다",
  "두 업종이 함께 걸린 곳은 대전 서구(298억)와 익산(164억)뿐이다");
s.addImage({ path: IMG+"15_지도.png", x:1.05, y:2.1, w:2.81, h:4.25 });
const rows = [
  ["대전 서구","일반한식","63.5","208억","−1.54"],
  ["경남 진주시","슈퍼마켓","35.7","130억","−2.22"],
  ["전북 익산시","일반한식","59.7","127억","−1.75"],
  ["서울 서초구","슈퍼마켓","28.2","106억","−2.73"],
  ["대전 서구","서양음식","53.5","90억","−1.80"],
  ["세종시","슈퍼마켓","49.3","82억","−1.52"],
  ["화성 동탄구","슈퍼마켓","39.1","77억","−2.02"],
  ["전주 완산구","서양음식","59.2","66억","−1.51"],
  ["강원 춘천시","편의점","65.5","54억","−1.71"],
  ["전주 덕진구","편의점","66.7","51억","−1.64"],
  ["전북 익산시","서양음식","55.3","37억","−1.70"],
  ["충북 충주시","서양음식","57.8","30억","−1.57"],
];
const head = [["지역","업종","침투지수","격차금액","σ"]];
const body = rows.map(r=>[
  {text:r[0], options:{color:INK, bold:true}}, {text:r[1], options:{color:MUTED}},
  {text:r[2], options:{color:C2, bold:true, align:"right"}},
  {text:r[3], options:{color:C1, bold:true, align:"right"}},
  {text:r[4], options:{color:MUTED, align:"right"}}]);
s.addTable([head[0].map((h,i)=>({text:h, options:{bold:true, color:"FFFFFF", fill:{color:C1},
    align:i>=2?"right":"left"}}))].concat(body),
  { x:4.2, y:2.1, w:8.2, colW:[2.2,1.6,1.35,1.6,1.45], rowH:0.32,
    fontFace:F, fontSize:10.5, border:{type:"solid", color:"ECECF3", pt:0.5},
    fill:{color:"FFFFFF"}, valign:"middle" });
concl(s, 6.45, "격차가 큰 순서가 곧 실행 순서는 아니다.",
  "12개는 성격이 다르고, 회수 가능성도 다르다.");
s.addNotes("표의 침투지수와 격차만 강조색. 모든 숫자를 같은 무게로 보여주지 않는다.");
}

/* ═══ 8. 건수 vs 건당 ═══ */
{
const s = P.addSlide();
frame(s, 1, "차이는 건당 금액이 아니라 결제 빈도에서 나왔다",
  "금액과 같은 방식으로 결제 건수를 예측해 비교했다 (전국 평균 = 100)");
[["26~69","결제 건수","전국 평균의 3~7할 수준",C2,T2,0.85],
 ["88~131","건당 금액","정상 범위",C3,T3,4.3]].forEach(d=>{
  s.addShape(P.ShapeType.roundRect, { x:d[5], y:2.15, w:3.2, h:1.6, rectRadius:0.08,
    fill:{color:d[4]}, line:{color:d[3], width:1.2} });
  s.addText(d[0], { x:d[5]+0.22, y:2.28, w:2.76, h:0.7, isTextBox:true, margin:0,
    fontFace:F, fontSize:32, bold:true, color:d[3], valign:"middle" });
  s.addText(d[1], { x:d[5]+0.22, y:3.0, w:2.76, h:0.3, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, bold:true, color:INK });
  s.addText(d[2], { x:d[5]+0.22, y:3.3, w:2.76, h:0.3, isTextBox:true, margin:0,
    fontFace:F, fontSize:10.5, color:MUTED });
});
s.addShape(P.ShapeType.roundRect, { x:0.85, y:3.95, w:6.65, h:1.95, rectRadius:0.08,
  fill:{color:"FBFBFD"}, line:{color:"E9E9F1", width:1} });
s.addText("그래서 혜택 설계가 달라진다", { x:1.1, y:4.1, w:6.15, h:0.34, isTextBox:true, margin:0,
  fontFace:F, fontSize:13, bold:true, color:INK });
[["✕","\"5만 원 이상 결제 시 할인\"  — 금액형", MUTED],
 ["○","\"이번 달 3회 이상 결제 시 혜택\"  — 빈도형", C1]].forEach((d,i)=>{
  const y = 4.55 + i*0.6;
  s.addText(d[0], { x:1.15, y, w:0.35, h:0.5, isTextBox:true, margin:0,
    fontFace:F, fontSize:14, bold:true, color:d[2], align:"center", valign:"middle" });
  s.addText(d[1], { x:1.55, y, w:5.6, h:0.5, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, bold:i===1, color:i===1?INK:MUTED, valign:"middle" });
});
s.addImage({ path: IMG+"18_건수건당.png", x:7.75, y:2.15, w:4.65, h:3.75 });
concl(s, 6.2, "저침투의 차이는 건당 금액보다 결제 빈도에서 크게 나타났다.",
  "따라서 1회 결제 금액을 높이는 오퍼보다 반복 결제를 유도하는 오퍼를 우선 검증한다.");
s.addNotes("데이터가 보여주는 것은 건수지수 26~69, 건당지수 88~131까지다. 원인은 단정하지 않고 오퍼 설계의 우선순위로 연결한다.");
}

/* 9쪽 : 3가지 기회 유형 */
{
const s = P.addSlide();
frame(s, 1, "12개는 같은 문제가 아니다 — 돈이 지금 어디 있는지로 갈린다",
  "업종 · 카테고리 · 지역 전체를 3층으로 비교하면 세 가지로 나뉜다  (모두 1인당 BC 결제액 지수, 전국=100)");

const cards = [
  { c:C1, t:T1, no:"1", nm:"지역 전체 기회", n:"7개", won:"673억",
    line:"어디서도 BC를 덜 쓴다.\n옮겨간 업종이 없다",
    ev:[["지역지수","57~72"],["카테고리÷지역","82~103"],["회수 가능성","가장 큼"]] },
  { c:C2, t:T2, no:"2", nm:"업종 간 이동", n:"4개", won:"280억",
    line:"지역은 정상인데\n이 업종만 낮다",
    ev:[["지역지수","87~100"],["카테고리÷지역","91~105"],["회수 가능성","중간"]] },
  { c:C3, t:T3, no:"3", nm:"채널 이동", n:"1개", won:"106억",
    line:"지역은 전국 상위인데\n이 업종만 무너졌다",
    ev:[["지역지수","135"],["카테고리÷지역","67"],["회수 가능성","낮음 — 검증용"]] },
];
cards.forEach((d,i)=>{
  const x = 0.85 + i*4.05, w = 3.75;
  s.addShape(P.ShapeType.roundRect, { x, y:2.12, w, h:3.45, rectRadius:0.07,
    fill:{color:d.t}, line:{color:d.c, width:1.2} });
  s.addShape(P.ShapeType.ellipse, { x:x+0.26, y:2.38, w:0.46, h:0.46, fill:{color:d.c} });
  s.addText(d.no, { x:x+0.26, y:2.38, w:0.46, h:0.46, isTextBox:true, margin:0,
    fontFace:F, fontSize:15, bold:true, color:"FFFFFF", align:"center", valign:"middle" });
  s.addText(d.nm, { x:x+0.84, y:2.40, w:w-1.1, h:0.42, isTextBox:true, margin:0,
    fontFace:F, fontSize:15, bold:true, color:INK, valign:"middle" });

  s.addText([{ text:d.n, options:{ fontSize:34, bold:true, color:d.c } },
             { text:"   "+d.won, options:{ fontSize:19, bold:true, color:MUTED } }],
    { x:x+0.26, y:2.96, w:w-0.52, h:0.62, isTextBox:true, margin:0, fontFace:F });

  s.addText(d.line, { x:x+0.26, y:3.62, w:w-0.52, h:0.68, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, color:INK, lineSpacingMultiple:1.15 });

  d.ev.forEach((e,k)=>{
    const yy = 4.42 + k*0.35;
    s.addText(e[0], { x:x+0.26, y:yy, w:1.85, h:0.3, isTextBox:true, margin:0,
      fontFace:F, fontSize:10.5, color:MUTED });
    s.addText(e[1], { x:x+2.05, y:yy, w:w-2.31, h:0.3, isTextBox:true, margin:0,
      fontFace:F, fontSize:11, bold:true, color:d.c, align:"right" });
  });
});

s.addShape(P.ShapeType.roundRect, { x:0.85, y:5.78, w:W-1.7, h:0.82, rectRadius:0.08,
  fill:{color:"F4F4F9"}, line:{color:"E3E3EE", width:1} });
s.addText([
  { text:"격차 1,059억은 한 덩어리가 아니다. ", options:{ color:INK, bold:true } },
  { text:"유형 1은 돈이 BC 밖에 있을 가능성이 높아 새 매출이 될 수 있고, 유형 2는 이미 BC 안에서 다른 업종으로 결제되고 있어 총액은 그대로일 수 있다.", options:{ color:MUTED } }],
  { x:1.1, y:5.9, w:W-2.2, h:0.6, isTextBox:true, margin:0, fontFace:F, fontSize:12.5, valign:"middle" });
s.addText([{text:"시범 3곳은 유형별 대표성 · 격차 규모 · 진단 가능성으로 골랐다. ", options:{bold:true, color:INK}},
  {text:"대전 서구(8개 업종이 전부 100 미만) · 진주(돈이 간 곳이 특정되는 유일한 곳) · 서초(유형 ③ 단독)", options:{color:MUTED}}],
  { x:0.85, y:6.68, w:W-1.7, h:0.3, isTextBox:true, margin:0, fontFace:F, fontSize:10, valign:"middle" });
s.addNotes("분류 규칙은 지역지수 80과 카테고리÷지역 80 두 가지만 쓴다. 대형할인점 지수(유형1 21~70 / 유형2·3 93~178)는 목적지를 보는 보조 지표이지 분류 기준이 아니다.");
}

/* 10쪽 : 유형별 처방 */
{
const s = P.addSlide();
frame(s, 2, "유형이 다르면 처방도 다르다 — 그리고 확장 경로도 다르다",
  "진단에서 나눈 세 유형에 처방을 1:1로 붙인다. 시범이 성공하면 같은 유형의 나머지 후보로 넓힌다.");

const col = [
  { c:C1, t:T1, nm:"1  지역 전체 기회", tag:"7개 · 673억",
    rx:"지자체 지역단위 제휴\n업종을 타깃하지 않는다",
    of:"서로 다른 날짜 반복 결제\n금액 조건 없이 빈도만",
    bd:"저이용 · 휴면 회원에 집중",
    tg:"전 연령 (최약 86~95)" },
  { c:C2, t:T2, nm:"2  업종 간 이동", tag:"4개 · 280억",
    rx:"업종 가맹점 제휴 강화",
    of:"동일 — 단 효과는\n카테고리 합계 증가로 확인",
    bd:"목적지 미확인 3곳은 소규모 시작",
    tg:"진주만 60대+ (78)\n나머지 3곳 전 연령" },
  { c:C3, t:T3, nm:"3  채널 이동", tag:"1개 · 106억",
    rx:"진단 우선\ne커머스 제휴 관점",
    of:"소규모 한정",
    bd:"진단 후 결정",
    tg:"60대+ 제외 전 연령\n(20~50대 75~81)" },
];
const rows = [["처방","rx"],["오퍼","of"],["예산","bd"],["타깃","tg"]];
rows.forEach((r,k)=>{
  s.addText(r[0], { x:0.85, y:2.62+k*0.86, w:0.8, h:0.74, isTextBox:true, margin:0,
    fontFace:F, fontSize:11.5, bold:true, color:MUTED, valign:"middle" });
});
col.forEach((d,i)=>{
  const x = 1.78 + i*3.68, w = 3.5;
  s.addShape(P.ShapeType.roundRect, { x, y:2.05, w, h:0.52, rectRadius:0.06, fill:{color:d.c} });
  s.addText(d.nm, { x:x+0.14, y:2.05, w:w-1.35, h:0.52, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, bold:true, color:"FFFFFF", valign:"middle" });
  s.addText(d.tag, { x:x+w-1.32, y:2.05, w:1.18, h:0.52, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, bold:true, color:"FFFFFF", align:"right", valign:"middle" });
  rows.forEach((r,k)=>{
    s.addShape(P.ShapeType.roundRect, { x, y:2.62+k*0.86, w, h:0.74, rectRadius:0.05,
      fill:{color: k%2===0 ? d.t : "FBFBFD"}, line:{color:"ECECF3", width:0.8} });
    s.addText(d[r[1]], { x:x+0.16, y:2.62+k*0.86, w:w-0.32, h:0.74, isTextBox:true, margin:0,
      fontFace:F, fontSize:11, color:INK, valign:"middle", lineSpacingMultiple:1.05 });
  });
});

s.addText("확장 규칙 — 시범이 6개월 판정을 통과하면", { x:0.85, y:6.14, w:4.6, h:0.3,
  isTextBox:true, margin:0, fontFace:F, fontSize:11.5, bold:true, color:INK });
const ext = [["대전 서구 성공","유형 1 나머지 5개 · 374억",C1],
             ["진주 성공","유형 2 나머지 3개 · 150억",C2],
             ["서초","확장 대상 없음 (판정용)",C3]];
ext.forEach((e,i)=>{
  const x = 1.78 + i*3.68;
  s.addShape(P.ShapeType.roundRect, { x, y:6.48, w:3.5, h:0.48, rectRadius:0.05,
    fill:{color:"F7F7FB"}, line:{color:e[2], width:1} });
  s.addText([{text:e[0]+"  ", options:{bold:true, color:e[2]}},{text:e[1], options:{color:MUTED}}],
    { x:x+0.14, y:6.48, w:3.22, h:0.48, isTextBox:true, margin:0, fontFace:F, fontSize:10.5, valign:"middle" });
});
s.addNotes("나머지 9곳은 실행안이 아니라 확장 후보. 같은 유형의 시범이 성공 판정을 받은 뒤 같은 규칙을 적용한다.");
}

/* ═══ 11. 대전 서구 ═══ */
{
const s = P.addSlide();
deepDive(s, "대전 서구 — 지역 전체가 약한 시장",
  "8개 업종이 전부 100 미만. 특정 업종이 아니라 이용 기반 전체의 문제다.",
  [["298억","두 업종 합산 격차","일반한식 208억 · 서양음식 90억",C1],
   ["66.5","지역 전체 지수","전국 227개 중 193위",C1],
   ["70.3","대형할인점 이용","옮겨간 업종이 없다",C2],
   ["27p","연령 편차","전 연령이 고르게 약하다",MUTED]],
  "11_대전서구.png",
  "지역 단위로 묶어 집행한다.",
  "업종 캠페인이 아니라 지자체·회원사와의 지역 제휴. 두 업종을 한 캠페인으로 묶고 성과만 따로 집계한다.",
  T1, C1,
  "대형할인점 70.3이 핵심 근거 — 돈이 지역 안 다른 업종으로 간 흔적이 없다. 지역 전체 격차는 539억이지만 판정된 후보는 298억.");
}

/* ═══ 12. 진주 ═══ */
{
const s = P.addSlide();
deepDive(s, "진주 — 지역은 정상인데 슈퍼마켓만 낮은 시장",
  "대형할인점 이용은 높은 반면 슈퍼마켓 이용은 매우 낮다. 세대 차이도 후보 중 가장 뚜렷하다.",
  [["35.7","슈퍼마켓 침투지수","격차 130억 · σ −2.22",C2],
   ["88.2","지역 전체 지수","지역 이용은 정상 수준",C1],
   ["178.2","대형할인점 이용","전국 227개 중 19위",C2],
   ["78","60대+ 연령지수","20·30대는 130 — 편차 52p",C3]],
  "12_진주.png",
  "업종 단독 제휴 — 효과는 장보기 합계 증가로 확인한다.",
  "슈퍼마켓이 올라도 대형할인점이 줄면 BC 총액은 그대로다.",
  T2, C2,
  "진주는 목적지가 대형할인점으로 특정되는 유일한 곳. 완산·춘천·충주는 목적지 미상이라 추적이 먼저다. 대조군 재선정 필요.");
}

/* 13쪽 : 서초 */
{
const s = P.addSlide();
frame(s, 2, "서초 — 지역은 전국 상위인데 이 업종 하나만 무너졌다",
  "서울 25개 자치구 중 슈퍼마켓 침투지수 최저. 대형할인점 이용은 높은 반면 슈퍼마켓 이용은 매우 낮다.");

const stat = [
  ["28.2","슈퍼마켓 침투지수","서울 25개구 중 최저 (중앙값 77.3)",C3],
  ["135.1","지역 전체 지수","전국 상위 — 카드 이용 자체는 강하다",C1],
  ["168.9","대형할인점 이용","강남 115 · 송파 114",C2],
  ["27","건수지수","건당은 105로 정상 — 횟수만 부족",MUTED],
];
stat.forEach((d,i)=>{
  const x = 0.85 + i*2.96, w = 2.74;
  s.addShape(P.ShapeType.roundRect, { x, y:2.12, w, h:1.32, rectRadius:0.07,
    fill:{color:"FBFBFD"}, line:{color:"E9E9F1", width:1} });
  s.addText(d[0], { x:x+0.18, y:2.2, w:w-0.36, h:0.58, isTextBox:true, margin:0,
    fontFace:F, fontSize:29, bold:true, color:d[3], valign:"middle" });
  s.addText(d[1], { x:x+0.18, y:2.78, w:w-0.36, h:0.26, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, bold:true, color:INK });
  s.addText(d[2], { x:x+0.18, y:3.03, w:w-0.36, h:0.36, isTextBox:true, margin:0,
    fontFace:F, fontSize:9.5, color:MUTED, lineSpacingMultiple:1.05 });
});

s.addImage({ path: IMG+"20_서초_슬라이드.png", x:1.52, y:3.62, w:10.3, h:2.61 });

s.addShape(P.ShapeType.roundRect, { x:0.85, y:6.44, w:W-1.7, h:0.56, rectRadius:0.08,
  fill:{color:T3}, line:{color:C3, width:1} });
s.addText([
  { text:"서초는 돈을 버는 시범이 아니라, 되찾을 수 있는 돈인지 가리는 시범이다. ", options:{ bold:true, color:INK } },
  { text:"격차 106억의 회수 가능 여부가 불확실해 목표를 5%로 낮게 잡았다. 검증에 실패하면 공략 대상을 1,059억에서 953억으로 확정한다.", options:{ color:MUTED } }],
  { x:1.1, y:6.44, w:W-2.2, h:0.56, isTextBox:true, margin:0, fontFace:F, fontSize:12, valign:"middle" });
s.addNotes("20~50대 연령지수 75~81, 60대+만 148. 온라인 장보기 이용층과 연령 구조가 일치한다는 간접 근거까지가 한계. 확정은 BC 내부 데이터 필요.");
}

/* 14 KPI와 판정 기준 */
{
const s = P.addSlide();
frame(s, 2, "KPI는 4개로 좁혔다 — 주 지표 2개, 선행·지속 2개",
  "성과 → 부작용 방지 → 행동 변화 → 지속성 순으로 쌓았고, 목표치는 유형마다 다르다");

// 상위 목표 연결
s.addShape(P.ShapeType.roundRect, { x:0.85, y:2.02, w:W-1.7, h:0.46, rectRadius:0.06,
  fill:{color:T1}, line:{color:C1, width:1} });
s.addText([{text:"BC 상위 목표  ", options:{bold:true, color:C1}},
  {text:"한정된 예산으로 저침투 시장의 추가 결제액 발생 검증", options:{color:INK}},
  {text:"   →   ", options:{color:C1, bold:true}},
  {text:"캠페인 목표  ", options:{bold:true, color:C1}},
  {text:"유형별 시범으로 회수 가능성을 검증하고 확장 여부를 판정", options:{color:INK}}],
  { x:1.1, y:2.02, w:W-2.2, h:0.46, isTextBox:true, margin:0, fontFace:F,
    fontSize:11.5, valign:"middle" });

// KPI 표
표(s, [["구분","center"],["KPI"],["측정 방법"],["목표치"],["주기","center"]],
  [["주","분기 건수지수","업종별 회귀 예측 대비 실제 ×100 · 대조군 대비 차이",
    "유형① +5.5p  /  ② +6.4p  /  ③ +3.6p","분기"],
   ["게이트","카테고리 합계 증가","유형① 외식 · ②③ 장보기 합계의 이중차분",
    "0 초과 — 미달 시 실패 처리","분기"],
   ["선행","저이용 · 휴면 회원 재활성화율","직전 3개월 미이용 회원의 재결제 비율",
    "BC 내부 기준선 설정 필요","월"],
   ["지속","혜택 종료 후 재결제율","혜택 종료 1개월 내 재결제 비율",
    "BC 내부 기준선 설정 필요","12개월"]],
  { x:0.85, y:2.62, w:11.63, colW:[0.9,2.6,4.3,3.0,0.83], rowH:0.42, fs:10, bold:[1] });

s.addText("보조 지표 (판정에 쓰지 않음) — 침투지수 · 참여 가맹점 수 · 참여 카드 수 · 첫 결제 전환율",
  { x:0.85, y:4.76, w:7.2, h:0.28, isTextBox:true, margin:0, fontFace:F,
    fontSize:10, color:MUTED, valign:"middle" });
s.addText("목표치는 최소 탐지 기준의 1.6~4.6배 — 잡음에 묻히지 않는 수준",
  { x:8.1, y:4.76, w:4.38, h:0.28, isTextBox:true, margin:0, fontFace:F,
    fontSize:10, color:C1, align:"right", valign:"middle" });

// 유형별 판정
s.addText("6개월 판정 — 유형마다 결론이 다르다", { x:0.85, y:5.08, w:5, h:0.28,
  isTextBox:true, margin:0, fontFace:F, fontSize:11.5, bold:true, color:INK });
const col = [
  { c:C1, t:T1, nm:"① 지역 전체 기회", 여유:"여유 4.6배",
    성공:"유형 ① 나머지 5개로 확장 · 374억",
    실패:"회원사 채널 재점검 · 확장 보류" },
  { c:C2, t:T2, nm:"② 업종 간 이동", 여유:"여유 1.6배",
    성공:"카테고리 합계 증가 확인 지역만 확장 · 150억",
    실패:"확장 중단" },
  { c:C3, t:T3, nm:"③ 채널 이동", 여유:"여유 3.3배",
    성공:"채널 대체 가설 검증 → 전국 재스크리닝",
    실패:"현 가설 제외 · 1,059억 → 953억" },
];
col.forEach((d,i)=>{
  const x = 0.85 + i*3.94, w = 3.75;
  s.addShape(P.ShapeType.roundRect, { x, y:5.40, w, h:0.40, rectRadius:0.05, fill:{color:d.c} });
  s.addText(d.nm, { x:x+0.14, y:5.40, w:w-1.25, h:0.40, isTextBox:true, margin:0,
    fontFace:F, fontSize:11.5, bold:true, color:"FFFFFF", valign:"middle" });
  s.addText(d.여유, { x:x+w-1.22, y:5.40, w:1.08, h:0.40, isTextBox:true, margin:0,
    fontFace:F, fontSize:10, bold:true, color:"FFFFFF", align:"right", valign:"middle" });
  [["성공","성공",d.t],["실패","실패","FBFBFD"]].forEach((r,k)=>{
    const y = 5.84 + k*0.42;
    s.addShape(P.ShapeType.roundRect, { x, y, w, h:0.40, rectRadius:0.04,
      fill:{color:r[2]}, line:{color:"ECECF3", width:0.8} });
    s.addText([{text:r[0]+"  ", options:{bold:true, color:d.c}},
               {text:d[r[1]], options:{color:INK}}],
      { x:x+0.14, y, w:w-0.28, h:0.40, isTextBox:true, margin:0, fontFace:F,
        fontSize:9.5, valign:"middle" });
  });
});

s.addText([{text:"선행 · 지속 KPI의 목표치는 BC가 정한다. ", options:{bold:true, color:INK}},
  {text:"회원 단위 이용 이력이 있어야 기준선을 세울 수 있어, 우리는 지표와 측정 방법까지만 제시한다.", options:{color:MUTED}}],
  { x:0.85, y:6.72, w:W-1.7, h:0.28, isTextBox:true, margin:0, fontFace:F,
    fontSize:10.5, valign:"middle" });
s.addNotes("성과(건수지수) → 부작용 방지(카테고리 합계) → 행동 변화(휴면 재활성화) → 지속성(종료 후 재결제)의 계층. 파일럿 실패가 회수 불가능의 증명은 아니므로 '회수 불가'가 아니라 '현 가설 제외'로 적는다.");
}

/* ═══ 15. 강건성 검증 ═══ */
{
const s = P.addSlide();
frame(s, 2, "이 결과는 모델 하나에 의존하지 않았다",
  "학습에 쓰지 않은 지역과 기간에서도, 다른 모델에서도 같은 후보가 반복해서 나타났다");
const v = [["12 / 12","교차검증 (5-fold)","자기 자신을 빼고 예측해도 후보 유지\n과적합 폭 0.2~1.3%p",C1],
           ["12 / 12","시도 단위 홀드아웃","강원 전체를 빼고 춘천을 예측\nR² 0.82~0.95",C1],
           ["12 / 12","시간 분할","1~3월로 학습해 4~6월 예측\n순위 상관 0.965~0.998",C3],
           ["0.946","반기 재산출 σ 상관","전반기·후반기를 따로 계산해도\n같은 지역이 낮게 나온다",C3]];
v.forEach((d,i)=>{
  const x = 0.85 + (i%2)*6.05, y = 2.15 + Math.floor(i/2)*1.72;
  s.addShape(P.ShapeType.roundRect, { x, y, w:5.78, h:1.52, rectRadius:0.08,
    fill:{color:"FBFBFD"}, line:{color:d[3], width:1.2} });
  s.addText(d[0], { x:x+0.22, y:y+0.14, w:1.9, h:0.62, isTextBox:true, margin:0,
    fontFace:F, fontSize:26, bold:true, color:d[3], valign:"middle" });
  s.addText(d[1], { x:x+2.2, y:y+0.18, w:3.35, h:0.34, isTextBox:true, margin:0,
    fontFace:F, fontSize:13, bold:true, color:INK, valign:"middle" });
  s.addText(d[2], { x:x+2.2, y:y+0.56, w:3.35, h:0.7, isTextBox:true, margin:0,
    fontFace:F, fontSize:10.5, color:MUTED, lineSpacingMultiple:1.1 });
});
concl(s, 5.78, "유의성은 주장하지 않는다.",
  "1,813개를 모두 검토해 하위를 선별하는 구조라 개별 p값은 다중비교 문제로 해석할 수 없다. 대신 재현성으로 검증했다.");
s.addText("※ 시도 홀드아웃에서는 12개 모두 침투지수가 더 낮게 나와, 현재 수치가 보수적인 쪽임을 확인했다",
  { x:0.85, y:6.55, w:W-1.7, h:0.3, isTextBox:true, margin:0, fontFace:F, fontSize:10, color:FAINT });
s.addNotes("FDR 보정은 업종당 227개라는 표본 한계로 어떤 후보도 통과할 수 없다. 질문이 나오면 재현성으로 답한다.");
}

/* ═══ 16. 한계 + 최종 제안 ═══ */
{
const s = P.addSlide();
s.background = { color: DARK };
s.addText("한계와 최종 제안", { x:0.9, y:0.6, w:11.5, h:0.5, isTextBox:true, margin:0,
  fontFace:F, fontSize:15, bold:true, color:"9B93CC" });
s.addShape(P.ShapeType.roundRect, { x:0.9, y:1.25, w:5.6, h:3.4, rectRadius:0.08,
  fill:{color:"2E2760"}, line:{color:"473E8C", width:1} });
s.addText("이번 분석에서 알 수 없는 것", { x:1.2, y:1.45, w:5.0, h:0.36, isTextBox:true, margin:0,
  fontFace:F, fontSize:14, bold:true, color:"FFFFFF" });
["타사 카드 이용 여부","현금 · 온라인 결제","소득 · 가구 구성","2026년 이전 장기 추세","시군구보다 작은 상권 단위"]
  .forEach((t,i)=>{
  s.addText("·  "+t, { x:1.2, y:1.92+i*0.42, w:5.0, h:0.36, isTextBox:true, margin:0,
    fontFace:F, fontSize:12.5, color:"D6D1EE", valign:"middle" });
});
s.addShape(P.ShapeType.roundRect, { x:1.2, y:4.02, w:5.0, h:0.46, rectRadius:0.06,
  fill:{color:"3B327A"} });
s.addText("1,059억은 회수 가능한 매출이 아니라 모델 기준 격차다",
  { x:1.35, y:4.02, w:4.7, h:0.46, isTextBox:true, margin:0, fontFace:F, fontSize:11,
    bold:true, color:"FFFFFF", valign:"middle" });

s.addText("전국 모든 시장을 공략하는 것이 아니라\nBC카드가 상대적으로 약한 시장을 찾아\n지역과 업종에 맞는 방식으로 검증한다",
  { x:7.0, y:1.5, w:5.4, h:2.0, isTextBox:true, margin:0, fontFace:F, fontSize:20,
    bold:true, color:"FFFFFF", lineSpacingMultiple:1.28 });
s.addShape(P.ShapeType.roundRect, { x:7.0, y:3.75, w:5.4, h:0.9, rectRadius:0.08,
  fill:{color:"2E2760"}, line:{color:"473E8C", width:1} });
s.addText("연간 스크리닝으로 매년 재산출한다", { x:7.25, y:3.75, w:4.9, h:0.9, isTextBox:true,
  margin:0, fontFace:F, fontSize:13, bold:true, color:"C9C4E8", valign:"middle" });

const flow = ["1,813개 시장","12개 후보","3개 시범","성과 검증","연간 스크리닝"];
flow.forEach((t,i)=>{
  const x = 0.9 + i*2.33;
  s.addShape(P.ShapeType.roundRect, { x, y:5.25, w:2.05, h:0.62, rectRadius:0.06,
    fill:{color: i===4 ? "5B4FCF" : "342C68"} });
  s.addText(t, { x, y:5.25, w:2.05, h:0.62, isTextBox:true, margin:0, fontFace:F,
    fontSize:11.5, bold:true, color:"FFFFFF", align:"center", valign:"middle" });
  if(i<4) s.addText("→", { x:x+2.05, y:5.25, w:0.28, h:0.62, isTextBox:true, margin:0,
    fontFace:F, fontSize:14, bold:true, color:"8F88BE", align:"center", valign:"middle" });
});
s.addText(FOOT, { x:0.9, y:6.55, w:11.5, h:0.3, isTextBox:true, margin:0,
  fontFace:F, fontSize:10, color:"8F88BE", align:"center" });
s.addNotes("마지막은 메시지 하나로 닫는다. 표지와 같은 어두운 배경으로 샌드위치 구조를 만든다.");
}

/* 201 데이터 출처 */
{
const s = P.addSlide();
frameApx(s, "부록 1. 데이터 출처",
  "분석에 사용한 모든 외부 데이터와 기준 시점");
표(s, [["데이터"],["제공처"],["기준 시점","center"],["용도"]],
  [["BC카드 소비데이터","대회 제공","2026.01~06","결제금액 · 건수 (분자)"],
   ["소상공인시장진흥공단_상가(상권)정보","공공데이터포털","2026.06","업종별 점포 수 (분모)"],
   ["행정안전부_생활_대규모점포 조회서비스","공공데이터포털","2026.08","대형마트 — 상가정보 미수록분 보완"],
   ["주민등록인구현황 (시군구·성별·연령)","행정안전부","2026.01~06","20대 이상 성인인구 (분모)"],
   ["카드사별 월별 이용실적 (표2 회원은행 포함)","여신금융협회","2026.01~06","BC 전국 점유율 기준선 — 해석 참고용"],
   ["시군구 경계 GeoJSON","southkorea-maps (통계청 2018 기준)","—","지도 시각화"]],
  { x:0.85, y:2.1, w:11.63, colW:[3.9,2.5,1.35,3.88], rowH:0.42 });
const note = [
  ["시점 불일치", "대규모점포만 2026년 8월 기준이고 나머지는 6월이다", C2],
  ["지역 제외", "상가정보에 광주·전남이 누락돼 분석에서 제외했다 (인구의 6.2%)", C2],
];
note.forEach((d,i)=>{
  const y = 5.12 + i*0.42;
  s.addText(d[0], { x:0.85, y, w:1.5, h:0.38, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, bold:true, color:d[2], valign:"middle" });
  s.addText(d[1], { x:2.4, y, w:10.0, h:0.38, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, color:MUTED, valign:"middle" });
});
s.addNotes("여신금융협회 데이터는 최종 지표 산출에 쓰이지 않았으므로 '해석 참고용'으로 표시한다.");
}

/* 202 업종 매칭 */
{
const s = P.addSlide();
frameApx(s, "부록 2. BC 업종과 상가정보 업종의 매칭",
  "서로 다른 두 데이터를 잇는 기준 — 대회 설명서의 TPBUZ_DT_NM을 근거로 삼았다");
표(s, [["BC 업종"],["상가정보 업종 코드"]],
  [["편의점","G20405"],
   ["슈퍼마켓","G20404"],
   ["일반한식","I201"],
   ["일식회집","I203"],
   ["중국음식","I202"],
   ["서양음식","I204 + I212 + I21003 + I21004 + I21005"],
   ["스넥","I21007 + I21006"],
   ["제과점","I21001 + I21002 + I21008"],
   ["대형할인점","(상가정보 미수록) 대규모점포 API · BZSTAT_SE_NM = 대형마트"]],
  { x:0.85, y:2.1, w:7.3, colW:[1.9,5.4], rowH:0.40, bold:[0] });
s.addShape(P.ShapeType.roundRect, { x:8.45, y:2.1, w:3.98, h:1.9, rectRadius:0.08,
  fill:{color:T2}, line:{color:C2, width:1.2} });
s.addText("한식 3종이 한 코드에 묶인다", { x:8.7, y:2.26, w:3.5, h:0.32, isTextBox:true, margin:0,
  fontFace:F, fontSize:12.5, bold:true, color:C2 });
s.addText("일반한식 · 갈비전문점 · 한정식은 대회 설명서가 동일해 상가정보 I201 하나에 함께 대응된다. 다만 갈비전문점(87억) · 한정식(12억)은 전국 한식 결제의 0.2%에 불과해, 일반한식에 합산해 다시 계산해도 침투지수 순위상관이 0.9999로 결과가 바뀌지 않는다.",
  { x:8.7, y:2.62, w:3.5, h:1.25, isTextBox:true, margin:0, fontFace:F, fontSize:10.5,
    color:INK, lineSpacingMultiple:1.16 });
s.addShape(P.ShapeType.roundRect, { x:8.45, y:4.15, w:3.98, h:1.75, rectRadius:0.08,
  fill:{color:"FBFBFD"}, line:{color:"E9E9F1", width:1} });
s.addText("대형할인점도 제외", { x:8.7, y:4.31, w:3.5, h:0.32, isTextBox:true, margin:0,
  fontFace:F, fontSize:12.5, bold:true, color:INK });
s.addText("상권이 행정구역을 넘어 점포 수·면적 모두 소비를 설명하지 못했다 (R² 0.47). BC 소비의 10.2%를 포기했음을 밝힌다. 다만 카테고리 합계 판정에는 포함한다.",
  { x:8.7, y:4.67, w:3.5, h:1.1, isTextBox:true, margin:0, fontFace:F, fontSize:10.5,
    color:MUTED, lineSpacingMultiple:1.16 });
concl(s, 6.2, "최종 분석 업종은 8개다.",
  "편의점 · 슈퍼마켓 · 일반한식 · 일식회집 · 중국음식 · 서양음식 · 스넥 · 제과점");
s.addNotes("두 데이터를 어떻게 연결했는지는 반드시 나오는 질문이다. 코드 단위까지 밝혀 둔다.");
}

/* 203 지표 정의 */
{
const s = P.addSlide();
frameApx(s, "부록 3. 지표 정의",
  "본문에서 쓴 모든 지표의 산식과 읽는 법");
표(s, [["지표"],["산식"],["읽는 법"]],
  [["침투지수","실제 BC 소비 ÷ 모델예측 × 100","100 = 예상만큼 결제됨. 50 = 예상의 절반"],
   ["모델예측","업종별 회귀 적합값 exp(β₀ + β₁·log 점포수 + β₂·log 성인인구)","그 상권 규모라면 나와야 할 결제액"],
   ["격차금액","모델예측 − 실제 소비 (6개월 합)","회수 가능 매출이 아니라 예상과의 차이"],
   ["시장 백분위","모델예측의 업종 내 백분위","100에 가까울수록 큰 시장"],
   ["σ (표준화 이탈지표)","log(침투지수÷100) ÷ 업종별 5-fold CV 예측오차","업종별 CV 예측오차 대비 하방 이탈 정도. −1.5 = 예측오차의 1.5배 이상 낮음"],
   ["건수지수","log(건수) ~ log(점포수)+log(인구) 예측 대비 실제 × 100","얼마나 자주 결제하나"],
   ["건당지수","지역 건당금액 ÷ 전국 동업종 건당금액 × 100","한 번에 얼마나 쓰나"],
   ["연령지수","지역 연령별 1인당 결제액 ÷ 전국 동업종 연령별 1인당 × 100","내국인 기준. 연령 간 상대 차이만 해석"]],
  { x:0.85, y:2.1, w:11.63, colW:[1.55,5.35,4.73], rowH:0.40, bold:[0] });
const pt = [["재현성","교차검증 폴드를 난수가 아니라 순차 분할로 고정해 매번 같은 결과가 나온다",C1],
            ["연령 분해만 내국인","분모가 주민등록인구라 외국인을 포함하면 20~30대 전국 기준선이 부풀려진다",C2]];
pt.forEach((d,i)=>{
  const y = 5.92 + i*0.42;
  s.addText(d[0], { x:0.85, y, w:2.1, h:0.4, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, bold:true, color:d[2], valign:"middle" });
  s.addText(d[1], { x:3.0, y, w:9.4, h:0.4, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, color:MUTED, valign:"middle" });
});
s.addNotes("본문에서 용어를 줄인 대신 정의는 부록에 모았다. σ는 순위 지표이지 가설검정 통계량이 아니다.");
}

/* 204 로그 변환과 회귀식 */
{
const s = P.addSlide();
frameApx(s, "부록 4. 왜 로그를 씌웠나",
  "세 변수 모두 지역 간 최대 19만 배 차이 나는 극단적 우편향 분포다");
s.addImage({ path: IMG+"12_로그변환.png", x:0.85, y:2.05, w:7.3, h:2.53 });
표(s, [["업종"],["R²","right"],["CV 오차","right"]],
  [["편의점","0.952","0.247"],["서양음식","0.920","0.348"],["일반한식","0.912","0.295"],
   ["중국음식","0.898","0.342"],["제과점","0.895","0.418"],["스넥","0.880","0.385"],
   ["일식회집","0.834","0.682"],["슈퍼마켓","0.833","0.465"]],
  { x:8.45, y:2.05, w:3.98, colW:[1.7,1.14,1.14], rowH:0.33, fs:10, bold:[0] });
s.addText("가중평균 R² 0.891", { x:8.45, y:5.05, w:3.98, h:0.3, isTextBox:true, margin:0,
  fontFace:F, fontSize:11.5, bold:true, color:C1, align:"center" });
const why = [["분포 치우침 제거","BC 소비 왜도 3.45 → −0.51",C1],
             ["대도시 쏠림 방지","상위 10% 지역의 오차 배율 2.3~4.2배 → 0.7~1.8배",C1],
             ["계수가 탄력성이 된다","소비 = A × 점포수^b × 인구^c 를 선형으로 푼다",MUTED],
             ["잔차가 곧 침투지수","exp(잔차) × 100 = 침투지수",MUTED]];
why.forEach((d,i)=>{
  const y = 4.70 + i*0.40;
  s.addText("·  "+d[0], { x:0.85, y, w:2.9, h:0.38, isTextBox:true, margin:0,
    fontFace:F, fontSize:11, bold:true, color:d[2], valign:"middle" });
  s.addText(d[1], { x:3.75, y, w:4.5, h:0.38, isTextBox:true, margin:0,
    fontFace:F, fontSize:10.5, color:MUTED, valign:"middle" });
});
concl(s, 6.38, "업종마다 따로 적합했다.",
  "8개를 한 식으로 묶으면 R²가 0.768로 떨어진다 — 업종마다 점포 규모와 객단가가 다르기 때문이다.");
s.addNotes("R² 비교는 종속변수가 달라 로그 변환의 근거로 쓰지 않는다. 근거는 대도시 오차 쏠림이다.");
}

P.writeFile({ fileName: "최종본.pptx" }).then(f=>console.log("생성:", f));
