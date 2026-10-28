// 국기원 태권도연구 투고 양식 원고 생성 (docx)
// 사용: NODE_PATH=/opt/node22/lib/node_modules node build_manuscript.js <out.docx>
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, BorderStyle,
  AlignmentType, SectionType, PageNumber, Footer, FootnoteReferenceRun, VerticalAlign,
  ShadingType, PageBreak, LineRuleType,
} = require("docx");

const OUT = process.argv[2] || "manuscript.docx";

// ---------------------------------------------------------------- 서식 상수 (국기원 기타 지침)
const MJ = { ascii: "Times New Roman", hAnsi: "Times New Roman", eastAsia: "바탕", cs: "Times New Roman" };   // 명조
const GT = { ascii: "Arial", hAnsi: "Arial", eastAsia: "돋움", cs: "Arial" };                               // 고딕
const pt = (x) => Math.round(x * 2);          // half-points
const LS = (pct) => Math.round(240 * pct / 100); // 줄간격 %
const MM = (x) => Math.round(x * 56.6929);       // mm → dxa

const PAGE_W = 11906, MARGIN = MM(20);
const BODY_W = PAGE_W - 2 * MARGIN;           // 1단 폭
const COL_GAP = MM(7);
const COL_W = Math.floor((BODY_W - COL_GAP) / 2);  // 2단 한 단 폭

// ---------------------------------------------------------------- 인라인 마크업
// 【...】 = 저자 확인 필요(노란 형광), *...* = 이탤릭, ~x~ = 아래첨자, ^x^ = 위첨자
function runs(text, o = {}) {
  const font = o.font || MJ, size = pt(o.size || 10);
  const parts = text.split(/(【[^】]*】|\*[^*]+\*|~[^~]+~|\^[^^]+\^)/).filter((s) => s !== "");
  return parts.flatMap((s) => {
    const base = { font, size, bold: o.bold, italics: o.italics, color: o.color, highlight: o.highlight ? "yellow" : undefined };
    if (s.startsWith("【")) {
      const hb = { ...base, highlight: "yellow" };
      return [new TextRun({ ...hb, text: "【" }), ...runs(s.slice(1, -1), { ...o, highlight: true }), new TextRun({ ...hb, text: "】" })];
    }
    if (s.startsWith("*") && s.endsWith("*") && s.length > 2) return new TextRun({ ...base, text: s.slice(1, -1), italics: true });
    if (s.startsWith("~") && s.endsWith("~")) return new TextRun({ ...base, text: s.slice(1, -1), subScript: true });
    if (s.startsWith("^") && s.endsWith("^")) return new TextRun({ ...base, text: s.slice(1, -1), superScript: true });
    return new TextRun({ ...base, text: s });
  });
}

// 본문: 명조 10p, 줄간격 160%, 첫 줄 10pt 들여쓰기, 양쪽 정렬
const body = (t) => new Paragraph({
  children: runs(t), alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LS(160), lineRule: LineRuleType.AUTO, before: 0, after: 0 },
  indent: { firstLine: 200 },
});
// 장 제목: 명조 15p 굵게
const h1 = (t) => new Paragraph({
  children: runs(t, { size: 15, bold: true }), keepNext: true,
  spacing: { before: 300, after: 120, line: LS(150) },
});
// 절 제목: 고딕 11p 굵게
const h2 = (t) => new Paragraph({
  children: runs(t, { font: GT, size: 11, bold: true }), keepNext: true,
  spacing: { before: 160, after: 80, line: LS(150) },
});
// 항 제목: 명조 10p 굵게
const h3 = (t) => new Paragraph({
  children: runs(t, { bold: true }), keepNext: true,
  spacing: { before: 100, after: 40, line: LS(160) },
});
const center = (t, o = {}) => new Paragraph({
  children: runs(t, o), alignment: AlignmentType.CENTER,
  spacing: { line: LS(o.ls || 150), before: o.before || 0, after: o.after || 0 },
});
const blank = (h = 120) => new Paragraph({ children: [], spacing: { before: 0, after: h } });

// ---------------------------------------------------------------- 표 (가로선만)
const NONE = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const THICK = { style: BorderStyle.SINGLE, size: 10, color: "000000" };
const THIN = { style: BorderStyle.SINGLE, size: 4, color: "000000" };

function table({ widths, header, rows, size = 8.5, align }) {
  const total = widths.reduce((a, b) => a + b, 0);
  const mk = (cellText, w, isHead, isFirstRow, isLastRow, colIdx, span) => new TableCell({
    width: { size: w, type: WidthType.DXA },
    columnSpan: span || 1,
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 40, bottom: 40, left: 60, right: 60 },
    borders: {
      left: NONE, right: NONE,
      top: isHead ? THICK : NONE,
      bottom: isHead ? THIN : (isLastRow ? THICK : NONE),
    },
    children: String(cellText).split("\n").map((line) => new Paragraph({
      children: runs(line, { size, bold: isHead, font: isHead ? GT : MJ }),
      alignment: isHead ? AlignmentType.CENTER : ((align && align[colIdx]) || (colIdx === 0 ? AlignmentType.LEFT : AlignmentType.CENTER)),
      spacing: { line: LS(125), before: 0, after: 0 },
      keepNext: !isLastRow,   // 작은 표가 단·쪽 사이에서 갈라지지 않도록
    })),
  });
  const headRow = new TableRow({
    tableHeader: true,
    children: header.map((h, i) => mk(h, widths[i], true, true, false, i)),
  });
  const bodyRows = rows.map((r, ri) => {
    const cells = [];
    let wi = 0;
    r.forEach((c, ci) => {
      const span = (typeof c === "object" && c !== null && c.span) ? c.span : 1;
      const txt = (typeof c === "object" && c !== null) ? c.text : c;
      const w = widths.slice(wi, wi + span).reduce((a, b) => a + b, 0);
      cells.push(mk(txt, w, false, false, ri === rows.length - 1, ci, span));
      wi += span;
    });
    return new TableRow({ children: cells, cantSplit: true });
  });
  return new Table({
    width: { size: total, type: WidthType.DXA }, columnWidths: widths,
    alignment: AlignmentType.CENTER,
    rows: [headRow, ...bodyRows],
  });
}
// 표 제목: 고딕 9p, 줄간격 150%, 표 위 왼쪽
const tcap = (t) => new Paragraph({
  children: runs(t, { font: GT, size: 9, bold: true }), keepNext: true,
  spacing: { before: 200, after: 60, line: LS(150) },
});
const tnote = (t) => new Paragraph({
  children: runs(t, { size: 8 }), spacing: { before: 40, after: 160, line: LS(130) },
});

// ---------------------------------------------------------------- 참고문헌 (명조 9p, 150%, 15pt 내어쓰기)
const ref = (t) => new Paragraph({
  children: runs(t, { size: 9 }), alignment: AlignmentType.LEFT,
  spacing: { line: LS(150), before: 0, after: 40 },
  indent: { left: 300, hanging: 300 },
});

// ================================================================ 내용
const TITLE_KO = "12주 고강도 기능성 인터벌 트레이닝에 따른 대학 태권도 겨루기 선수의 최대산소섭취량과 기능적 움직임 변화";
const TITLE_EN = "Changes in Maximal Oxygen Uptake and Functional Movement Following 12 Weeks of High-Intensity Functional Interval Training in Collegiate Taekwondo Sparring Athletes";

const ABS_KO = [
  ["목적", "이 연구는 12주간의 고강도 기능성 인터벌 트레이닝(H-IFIT)에 따른 대학 태권도 겨루기 선수의 최대산소섭취량(VO~2~max)과 기능적 움직임의 변화를 규명하는 데 목적이 있다."],
  ["방법", "대학 태권도 겨루기 선수 25명을 실험군(*n* = 13)과 통제군(*n* = 12)으로 나누어, 실험군은 12주간 주 3회, 회당 60분의 H-IFIT를 실시하고 통제군은 정규훈련을 실시하였다. VO~2~max는 트레드밀 점증부하검사로, 기능적 움직임은 기능적 움직임 검사(FMS) 총점으로 측정하였다. 자료는 2(집단)×2(시기) 반복측정 분산분석과 대응표본 *t*검정으로 분석하였으며, 사전값을 공변량으로 한 공분산분석과 비모수 검정을 보조적으로 실시하였다."],
  ["결과", "실험군의 VO~2~max는 【사전 M±SD】에서 【사후 M±SD】 mL/kg/min으로, FMS 총점은 【사전 M±SD】점에서 【사후 M±SD】점으로 【변화 방향】하였다. 집단과 시기 간 상호작용은 VO~2~max에서 【유의/비유의(*p*)】, FMS 총점에서 【유의/비유의(*p*)】하였으며(η~p~^2^ = 【 】), 통제군은 【결과】."],
  ["결론", "【재분석 결과에 맞춰 기재. 예: H-IFIT는 대학 태권도 겨루기 선수의 유산소성 체력과 기본 움직임 수행을 향상시키는 훈련 방법으로 활용될 수 있다.】"],
];
const KW_KO = "고강도 기능성 인터벌 트레이닝, 태권도 겨루기, 최대산소섭취량, 기능적 움직임 검사, 대학 선수";

const ABS_EN = [
  ["Purpose", "This study examined changes in maximal oxygen uptake (VO~2~max) and functional movement following 12 weeks of high-intensity functional interval training (H-IFIT) in collegiate taekwondo sparring athletes."],
  ["Methods", "Twenty-five collegiate taekwondo sparring athletes were assigned to an experimental group (*n* = 13) or a control group (*n* = 12). The experimental group performed H-IFIT three times per week, 60 minutes per session, for 12 weeks, whereas the control group continued regular training. VO~2~max was measured with a graded treadmill exercise test, and functional movement was assessed with the Functional Movement Screen (FMS). Data were analyzed with two-way (group × time) repeated-measures ANOVA and paired *t*-tests, with baseline-adjusted ANCOVA and nonparametric tests as sensitivity analyses."],
  ["Results", "In the experimental group, VO~2~max changed from 【pre M ± SD】 to 【post M ± SD】 mL/kg/min, and the FMS composite score changed from 【pre M ± SD】 to 【post M ± SD】 points. The group × time interaction was 【significant/non-significant】 for VO~2~max (【F, p】) and 【significant/non-significant】 for the FMS composite score (【F, p】). The control group 【results】."],
  ["Conclusion", "【Write according to the re-analysis, e.g., H-IFIT may be used as a training method to improve aerobic fitness and fundamental movement patterns in collegiate taekwondo sparring athletes.】"],
];
const KW_EN = "high-intensity functional interval training, taekwondo sparring, maximal oxygen uptake, Functional Movement Screen, collegiate athletes";

const INTRO = [
  "태권도 겨루기는 짧은 시간 동안 폭발적인 발차기와 스텝이 반복되고, 그 사이에 상대적으로 낮은 강도의 이동과 회복이 이어지는 간헐적 고강도 종목이다. 모의 경기에서 에너지 기여도를 분석한 Campos 등(2012)은 고강도 공격 동작에는 무산소성 비젖산 시스템이 크게 기여하지만, 경기 전체로 보면 유산소성 대사의 기여가 상당하다고 보고하였다. 국제 수준 태권도 선수의 신체적·생리적 특성을 정리한 Bridge 등(2014)도 선수 간 최대산소섭취량(VO~2~max)에 차이가 있으나, 경기의 대사적 요구를 감당하고 연속되는 경기 사이의 회복을 돕기 위해서는 중등도 이상의 심폐체력이 필요하다고 하였다. 따라서 겨루기 선수의 체력 훈련은 고강도 동작의 반복 수행 능력과 함께 유산소성 체력을 효율적으로 향상시키는 방향으로 구성될 필요가 있다.",
  "고강도 인터벌 트레이닝(high-intensity interval training, HIIT)은 짧은 고강도 운동과 회복을 반복하는 훈련 방식으로, 건강한 성인을 대상으로 한 메타분석에서 지속적 지구성 훈련과 함께 VO~2~max를 크게 향상시키는 방법으로 보고되었다(Milanović et al., 2015). 태권도 선수를 대상으로 한 중재연구에서도 인터벌 훈련을 고강도 지속 달리기와 비교한 연구(Monks et al., 2017), 반복 질주 훈련과 발차기 기반 반복 기술 훈련을 정규훈련에 추가하여 비교한 연구(Ouergui et al., 2020), 정규훈련에 주 2회의 인터벌 훈련을 8주간 추가하여 VO~2~max의 증가를 보고한 연구(Batra & Zatoń, 2016), 짧은 질주 인터벌과 기술 특이적 인터벌을 비교한 연구(Song & Sheykhlouvand, 2024)가 수행되어, 인터벌 형식의 훈련이 태권도 선수의 체력 향상에 활용될 수 있음이 제시되었다. 그러나 이들 연구는 대부분 달리기나 발차기 위주의 인터벌을 8주 이하의 비교적 짧은 기간 동안 적용하였으며, 다양한 기능적 동작을 결합한 12주 이상의 프로그램 효과는 상대적으로 덜 검토되었다.",
  "고강도 기능성 훈련(high-intensity functional training, HIFT)은 다관절의 기능적 동작을 다양한 형태로 구성하여 높은 강도로 수행하는 훈련으로, 정해진 운동 형태를 반복하는 HIIT와 구별되는 체력 향상 방법으로 정의되었다(Feito et al., 2018). 이 연구에서 적용한 고강도 기능성 인터벌 트레이닝(high-intensity functional interval training, H-IFIT)은 점프·런지와 같은 하지 동작, 몸통 안정화 동작, 방향 전환 및 발차기 동작을 인터벌 구조로 조합한 프로그램이다. 이러한 구성은 심폐계에 대한 고강도 자극과 함께 기본 움직임 패턴의 수행에도 영향을 줄 수 있다.",
  "기능적 움직임 검사(Functional Movement Screen, FMS)는 일곱 가지 기본 동작을 0–3점으로 평가하여 움직임 패턴의 제한과 좌우 비대칭을 확인하는 도구이다(Cook et al., 2014a, 2014b). 프로 미식축구 선수에서 FMS 총점 14점 이하가 심각한 부상 위험의 증가와 관련되었다는 보고(Kiesel et al., 2007) 이후 선수 현장에서 널리 사용되고 있으나, 체계적 문헌고찰에서는 총점만으로 부상을 예측하는 근거가 충분하지 않다고 하였다(Moran et al., 2017). 한편 FMS 점수는 훈련을 통해 변화할 수 있으며(Kiesel et al., 2011), 선수 대상 운동중재 연구를 종합한 메타분석에서는 코어 안정성, 저항, 기능성 훈련 등 여러 형태의 훈련이 FMS 총점을 향상시키는 것으로 나타났다(Maleki et al., 2025). 태권도 분야에서도 태권도 전공자에게 12주간 HIIT를 적용하여 FMS 변화를 분석한 연구가 보고되었다(김현륜, 이상일, 2026).",
  "겨루기 선수는 한 발 지지 상태에서 발차기를 반복하고 빠르게 방향을 전환해야 하므로, 심폐체력과 함께 하지와 몸통의 기본 움직임 능력이 경기 수행의 바탕이 된다고 볼 수 있다. 그러나 대학 태권도 겨루기 선수에게 기능적 동작 기반의 인터벌 프로그램을 적용하고, 심폐체력(VO~2~max)과 기능적 움직임(FMS)의 변화를 정규훈련 집단과 함께 비교한 연구는 부족하다. 이에 이 연구는 12주간의 H-IFIT에 따른 대학 태권도 겨루기 선수의 VO~2~max와 FMS 총점의 변화를 정규훈련 집단과 비교하여 규명하고, 겨루기 선수의 체력 훈련 프로그램 구성을 위한 기초자료를 제공하는 데 목적이 있다. 이를 위해 '12주간의 H-IFIT에 참여한 실험군은 정규훈련만 실시한 통제군에 비해 VO~2~max와 FMS 총점이 더 크게 향상될 것이다'라는 가설을 설정하였다【연구 계획 당시의 가설과 일치하는지 확인】.",
];

const M1 = [
  "이 연구의 대상자는 【지역】 소재 【대학교】 태권도부에 소속된 대학 태권도 겨루기 선수 25명이었다. 대상자는 실험군 13명과 통제군 12명으로 【배정 방법: 무작위 배정(방법 기재) 또는 소속·훈련 일정에 따른 배정】하였다. 선정 기준은 【예: 선수 경력 5년 이상, 최근 6개월 이내 근골격계 부상이 없는 자】였으며, 【제외 기준】에 해당하는 자는 제외하였다. 표본 크기는 【G*Power 산출 근거(효과크기, α, 검정력) 또는 '사전 산출 없음'】. 중재 기간 중 탈락자는 【 】명이었으며, 최종 분석 대상은 실험군 【13】명, 통제군 【12】명이었다.",
  "모든 대상자에게 연구의 목적과 절차, 예상되는 위험 및 참여 철회의 권리를 설명하고 서면 동의를 받았다. 이 연구는 헬싱키 선언의 윤리 원칙에 따라 【기관명】 기관생명윤리위원회의 승인(승인번호: 【 】)을 받아 수행하였다. 대상자의 일반적 특성은 <표 1>과 같다.",
];
const M2 = [
  "실험군은 12주간 주 3회, 회당 60분의 H-IFIT에 참여하였다. 각 회기는 준비운동 10분, 본운동 40분, 정리운동 10분으로 구성하였다(표 2). 본운동은 하지 폭발력·이동 동작(스쿼트 점프, 점핑 런지, 버피, 스텝업), 몸통 안정화·회전 동작(플랭크, 사이드 플랭크, 밴드 회전, 메디신볼 슬램), 태권도 특이 이동·발차기 동작(셔틀런, 방향 전환, 스텝과 발차기, 교대 돌려차기) 가운데 회기마다 6–10종목을 선택하여 서킷 형태로 실시하였다. 종목당 운동 시간은 30–40초, 세트 수는 3–4세트였으며, 운동 대 휴식 비율은 1:1에서 2:1로 점진적으로 높였다【주차별 진행 기준: 예) 1–4주 1:1, 5–8주 1.5:1, 9–12주 2:1】.",
  "운동 강도는 최대심박수(HRmax)의 85–95%와 자각운동강도(RPE) 15–18을 목표로 하였으며, 【심박수 측정 장비와 HRmax 산출 방법】. 프로그램은 【지도자 자격·경력】이 지도하였고, 실험군의 프로그램 출석률은 【 】%였다. 통제군은 같은 기간 동안 【정규훈련 내용, 주당 횟수·시간】의 정규훈련을 실시하였으며, 실험군은 【정규훈련에 H-IFIT를 추가하였다 / 정규훈련 일부를 H-IFIT로 대체하였다】. H-IFIT는 【정규훈련 전/후 또는 별도 시간】에 실시하였다.",
];
const M3a = "VO~2~max는 트레드밀(【제조사, 모델, 국가】)과 호흡가스 분석기(【제조사, 모델, 국가】)를 이용한 점증부하검사로 측정하였다. 검사는 【프로토콜: 예) Bruce 프로토콜】로 실시하였고, 호흡가스는 【산출 방식: 예) 30초 평균】으로 분석하였다. 대상자가 더 이상 운동을 지속할 수 없을 때 검사를 종료하였으며, 【판정 기준: 예) 호흡교환율 1.10 이상, 연령 예측 최대심박수의 95% 이상, RPE 17 이상 중 2개 이상 충족】한 경우 최대 수준에 도달한 것으로 판단하였다(Midgley et al., 2007). 사전 검사는 중재 시작 전 【 】일 이내, 사후 검사는 중재 종료 후 【 】일 이내에 동일한 시간대와 조건에서 실시하였으며, 대상자에게는 검사 전 24시간 동안 고강도 운동과 카페인 섭취를 제한하도록 하였다【실제 통제 내용 확인】.";
const M3b = "FMS는 Cook 등(2014a, 2014b)이 제시한 절차에 따라 Deep Squat, Hurdle Step, In-line Lunge, Shoulder Mobility, Active Straight Leg Raise, Trunk Stability Push-up, Rotary Stability의 7개 항목으로 평가하였다. 각 항목은 0–3점으로 채점하였고, 좌우를 평가하는 항목은 낮은 쪽 점수를 항목 점수로 하였으며, 어깨 충돌·척추 신전·척추 굴곡 배제 검사(clearing test)에서 통증이 나타나면 해당 항목을 0점으로 처리하였다. 이 연구에서는 7개 항목의 합(0–21점)인 총점을 분석에 사용하였다. 평가는 【FMS 자격을 갖춘 평가자 n명】이 사전·사후 모두 【동일하게, 집단 배정을 모르는 상태에서】 실시하였으며, 각 항목은 【3회 시도 중 최고 점수】를 기록하였다【평가자 간 신뢰도(ICC 또는 kappa) 기재】. FMS는 VO~2~max 검사와 【다른 날 / 검사 전】에 측정하였다.";
const M4 = "자료는 SPSS 25.0(IBM Corp., Armonk, NY, USA)을 이용하여 분석하였으며, 모든 측정값은 평균과 표준편차로 제시하였다. 자료의 정규성은 Shapiro-Wilk 검정으로 확인하였다. 집단 간 사전 동질성은 연속변인의 경우 독립표본 *t*검정으로, 성별은 Fisher의 정확검정으로 확인하였다. 집단과 측정 시기에 따른 VO~2~max와 FMS 총점의 변화는 2(집단)×2(시기) 반복측정 분산분석으로 분석하였고, 상호작용이 유의한 경우 집단별 사전·사후 차이를 대응표본 *t*검정으로 확인하였다. 사전값에 집단 간 차이가 있을 가능성을 고려하여 사전값을 공변량으로 한 공분산분석을 추가로 실시하였으며, 정규성 가정이 충족되지 않은 변인은 Wilcoxon 부호순위검정과 Mann-Whitney *U* 검정(변화량) 결과를 함께 확인하였다. 효과크기는 부분 에타제곱(η~p~^2^)으로 제시하였다(Lakens, 2013). 통계적 유의수준은 α = .05로 설정하였다.";

const R0 = "두 집단의 일반적 특성과 사전 VO~2~max, FMS 총점은 【유의한 차이가 없었다 / ○○에서 유의한 차이가 있었다】(표 1). 【차이가 있는 경우: 따라서 사전값을 공변량으로 한 공분산분석 결과를 함께 제시하였다】.";
const R1 = "VO~2~max의 변화는 <표 3>과 같다. 실험군의 VO~2~max는 사전 【M±SD】 mL/kg/min에서 사후 【M±SD】 mL/kg/min으로 【변화량】 mL/kg/min(【변화율】%) 【증가/감소】하였으며(*t* = 【 】, *p* 【 】), 통제군은 사전 【M±SD】 mL/kg/min에서 사후 【M±SD】 mL/kg/min으로 【변화 기술】(*t* = 【 】, *p* 【 】). 집단과 시기 간 상호작용 효과는 【유의하였다/유의하지 않았다】(*F*(1, 【df】) = 【 】, *p* 【 】, η~p~^2^ = 【 】). 사전값을 공변량으로 한 공분산분석에서도 사후 VO~2~max의 집단 간 차이는 【유의하였다/유의하지 않았다】(*F*(1, 【df】) = 【 】, *p* 【 】)【비모수 검정 결과가 다르면 함께 기재】.";
const R2 = "FMS 총점의 변화는 <표 3>과 같다. 실험군의 FMS 총점은 사전 【M±SD】점에서 사후 【M±SD】점으로 【변화량】점 【증가/감소】하였으며(*t* = 【 】, *p* 【 】), 통제군은 사전 【M±SD】점에서 사후 【M±SD】점으로 【변화 기술】(*t* = 【 】, *p* 【 】). 집단과 시기 간 상호작용 효과는 【유의하였다/유의하지 않았다】(*F*(1, 【df】) = 【 】, *p* 【 】, η~p~^2^ = 【 】). 사전값을 공변량으로 한 공분산분석에서도 사후 FMS 총점의 집단 간 차이는 【유의하였다/유의하지 않았다】(*F*(1, 【df】) = 【 】, *p* 【 】)【비모수 검정 결과가 다르면 함께 기재】.";

const DISC = [
  "이 연구는 12주간의 H-IFIT에 따른 대학 태권도 겨루기 선수의 VO~2~max와 기능적 움직임의 변화를 정규훈련 집단과 비교하였다. 그 결과 【재분석 결과 요약: 예) 실험군에서 VO~2~max와 FMS 총점이 유의하게 증가하였고, 두 변인 모두 집단과 시기 간 상호작용이 유의하였다】.",
  "실험군의 VO~2~max는 【변화량】 mL/kg/min(【변화율】%) 【증가/감소】하였다. 실험군의 사전 VO~2~max(【 】 mL/kg/min)는 태권도 선수에서 보고된 값과 비교할 때 【 】 수준이었다(Bridge et al., 2014). 인터벌 훈련은 VO~2~max 향상에 효과적인 방법으로 보고되었으며(Milanović et al., 2015), 태권도 선수에게 인터벌 훈련을 적용한 선행연구에서도 유산소성 체력의 향상이 보고되었다(Monks et al., 2017; Song & Sheykhlouvand, 2024). 특히 정규훈련에 주 2회의 인터벌 훈련을 8주간 추가한 ITF 태권도 수련자의 VO~2~max는 50.13 mL/kg/min에서 53.16 mL/kg/min으로 증가하였다(Batra & Zatoń, 2016). 이 연구의 변화량은 이러한 선행연구와 비교하여 【비슷한/더 큰/더 작은】 수준이었다. 다만 선행연구는 대부분 8주 이하의 달리기 또는 기술 기반 인터벌을 적용하였으므로, 운동 형태와 기간, 비교 조건이 다른 이 연구와 효과의 크기를 직접 비교하는 데에는 한계가 있다.",
  "H-IFIT는 전신 동작을 짧은 휴식과 함께 반복하는 구조로, 운동 중 심박수를 높은 수준으로 유지하게 하여 심폐계에 반복적인 고강도 자극을 준다. 이러한 자극은 일회박출량 증가와 같은 중추적 적응과 골격근의 산소 이용 능력 향상과 같은 말초적 적응을 통해 VO~2~max를 높일 수 있으나, 이 연구에서는 이러한 생리적 변인을 직접 측정하지 않았으므로 기전은 추후 검증이 필요하다. 또한 이 연구의 VO~2~max는 체중당 상대값이므로 체중 변화도 결과에 영향을 줄 수 있다【사전·사후 체중 및 절대 VO~2~max 결과 기재】. 【사전값 차이가 있었다면: 두 집단의 사전 VO~2~max에 차이가 있었으나, 사전값을 공변량으로 통제한 분석에서도 결과가 유지되었다/달라졌다】. 한편 실험군은 【정규훈련에 H-IFIT를 추가로 수행하였으므로】 실험군의 변화에는 프로그램 구성뿐 아니라 총 훈련량의 차이가 함께 작용했을 수 있으며, 이 결과는 정규훈련에 H-IFIT를 【추가】하였을 때의 효과로 해석하는 것이 적절하다.",
  "실험군의 FMS 총점은 【변화량】점 【증가/감소】하였다. 사전 총점(【 】점)은 대학 선수의 평균으로 보고된 14.8점(Pollen et al., 2021)과 비교할 때 【 】 수준이었다. 선수 대상 운동중재 연구를 종합한 메타분석에서 기능성 훈련은 1.74점, 코어 안정성 훈련은 2.89점의 FMS 총점 향상을 보였으며(Maleki et al., 2025), 이 연구의 변화량은 【이 범위와 비교하여 기재】. 태권도 전공자에게 12주간 HIIT를 적용한 김현륜과 이상일(2026)도 FMS의 향상을 보고하였다. 다만 FMS 점수는 검사 동작과 채점 기준에 익숙해지는 것만으로도 높아질 수 있으며(Frost et al., 2015), 총점의 최소탐지변화량은 평가자 간 2.1점, 평가자 내 2.5점으로 보고되었다(Teyhen et al., 2012). 따라서 통제군의 변화와 함께 비교한 집단 간 차이를 중심으로 해석해야 하며, 집단 평균의 변화를 개별 선수의 의미 있는 변화로 해석할 때에는 측정오차를 함께 고려해야 한다.",
  "FMS 총점 14점 이하가 부상 위험 증가와 관련된다는 보고가 있으나(Kiesel et al., 2007), 체계적 문헌고찰에서는 총점만으로 부상을 예측하는 근거가 제한적이었다(Moran et al., 2017). 이 연구는 부상 발생을 추적하지 않았으므로, FMS 점수의 변화를 부상 예방이나 경기력 향상으로 직접 연결하여 해석하는 데에는 신중해야 한다.",
  "이 연구의 제한점은 다음과 같다. 첫째, 한 대학의 소수 선수를 대상으로 하였고 【무작위 배정을 하지 않아】 결과를 일반화하는 데 한계가 있다. 둘째, 두 집단의 총 훈련량을 동일하게 통제하지 못하였다【확인】. 셋째, 【FMS 평가자의 눈가림 또는 신뢰도에 관한 제한점】. 넷째, FMS는 총점만 분석하였으므로 어떤 기본 동작이 변화하였는지는 확인하지 못하였다. 다섯째, 중재 종료 후 효과의 지속성과 실제 경기력 및 부상 발생과의 관련성을 확인하지 못하였다. 후속 연구에서는 훈련량을 통제한 무작위 대조 설계와 추적 관찰을 통해 이를 보완할 필요가 있다.",
];
const CONC = [
  "【재분석 결과에 맞춰 기재. 예: 이 연구에서 12주간의 H-IFIT에 참여한 대학 태권도 겨루기 선수는 정규훈련 집단에 비해 VO~2~max와 FMS 총점에서 더 큰 향상을 보였다.】",
  "【결과에 따른 현장 적용 제언 기재】. 후속 연구에서는 훈련량을 통제한 무작위 대조 설계로 효과를 검증하고, 경기력 및 부상 발생과의 관련성을 확인할 필요가 있다.",
];

const REFS = [
  "김현륜, 이상일 (2026). 12주간 고강도 인터벌 트레이닝(HIIT)이 태권도 전공자의 기능적 움직임 평가(FMS)에 미치는 영향. *대한무도학회지, 28*(1), 123-134.",
  "Batra, A., & Zatoń, M. (2016). Effect of high intensity interval training on cardiopulmonary function in Taekwon-do ITF athletes. *Journal of Combat Sports and Martial Arts, 7*, 73-79. https://doi.org/10.5604/20815735.1225636",
  "Bridge, C. A., Ferreira da Silva Santos, J., Chaabène, H., Pieter, W., & Franchini, E. (2014). Physical and physiological profiles of taekwondo athletes. *Sports Medicine, 44*(6), 713-733. https://doi.org/10.1007/s40279-014-0159-9",
  "Campos, F. A. D., Bertuzzi, R., Dourado, A. C., Santos, V. G. F., & Franchini, E. (2012). Energy demands in taekwondo athletes during combat simulation. *European Journal of Applied Physiology, 112*(4), 1221-1228. https://doi.org/10.1007/s00421-011-2071-4",
  "Cook, G., Burton, L., Hoogenboom, B. J., & Voight, M. (2014a). Functional movement screening: The use of fundamental movements as an assessment of function - Part 1. *International Journal of Sports Physical Therapy, 9*(3), 396-409.",
  "Cook, G., Burton, L., Hoogenboom, B. J., & Voight, M. (2014b). Functional movement screening: The use of fundamental movements as an assessment of function - Part 2. *International Journal of Sports Physical Therapy, 9*(4), 549-563.",
  "Feito, Y., Heinrich, K. M., Butcher, S. J., & Poston, W. S. C. (2018). High-intensity functional training (HIFT): Definition and research implications for improved fitness. *Sports, 6*(3), 76. https://doi.org/10.3390/sports6030076",
  "Frost, D. M., Beach, T. A. C., Callaghan, J. P., & McGill, S. M. (2015). FMS scores change with performers' knowledge of the grading criteria: Are general whole-body movement screens capturing \"dysfunction\"? *Journal of Strength and Conditioning Research, 29*(11), 3037-3044.",
  "Kiesel, K., Plisky, P., & Butler, R. (2011). Functional movement test scores improve following a standardized off-season intervention program in professional football players. *Scandinavian Journal of Medicine & Science in Sports, 21*(2), 287-292. https://doi.org/10.1111/j.1600-0838.2009.01038.x",
  "Kiesel, K., Plisky, P. J., & Voight, M. L. (2007). Can serious injury in professional football be predicted by a preseason functional movement screen? *North American Journal of Sports Physical Therapy, 2*(3), 147-158.",
  "Lakens, D. (2013). Calculating and reporting effect sizes to facilitate cumulative science: A practical primer for t-tests and ANOVAs. *Frontiers in Psychology, 4*, 863. https://doi.org/10.3389/fpsyg.2013.00863",
  "Maleki, A. A., Mousavi, S. H., Biabangard, M. A., & Minoonejad, H. (2025). Influence of exercise interventions on functional movement screen scores in athletes: A systematic review and meta-analysis. *Scientific Reports, 15*, 26335. https://doi.org/10.1038/s41598-025-12371-2",
  "Midgley, A. W., McNaughton, L. R., Polman, R., & Marchant, D. (2007). Criteria for determination of maximal oxygen uptake: A brief critique and recommendations for future research. *Sports Medicine, 37*(12), 1019-1028.",
  "Milanović, Z., Sporiš, G., & Weston, M. (2015). Effectiveness of high-intensity interval training (HIT) and continuous endurance training for VO~2~max improvements: A systematic review and meta-analysis of controlled trials. *Sports Medicine, 45*(10), 1469-1481. https://doi.org/10.1007/s40279-015-0365-0",
  "Monks, L., Seo, M.-W., Kim, H.-B., Jung, H. C., & Song, J. K. (2017). High-intensity interval training and athletic performance in Taekwondo athletes. *The Journal of Sports Medicine and Physical Fitness, 57*(10), 1252-1260. https://doi.org/10.23736/S0022-4707.17.06853-0",
  "Moran, R. W., Schneiders, A. G., Mason, J., & Sullivan, S. J. (2017). Do Functional Movement Screen (FMS) composite scores predict subsequent injury? A systematic review with meta-analysis. *British Journal of Sports Medicine, 51*(23), 1661-1669. https://doi.org/10.1136/bjsports-2016-096938",
  "Ouergui, I., Messaoudi, H., Chtourou, H., Wagner, M. O., Bouassida, A., Bouhlel, E., . . . Engel, F. A. (2020). Repeated sprint training vs. repeated high-intensity technique training in adolescent taekwondo athletes: A randomized controlled trial. *International Journal of Environmental Research and Public Health, 17*(12), 4506. https://doi.org/10.3390/ijerph17124506",
  "Pollen, T. R., Keitt, F., & Trojian, T. H. (2021). Do normative composite scores on the Functional Movement Screen differ across high school, collegiate, and professional athletes? A critical review. *Clinical Journal of Sport Medicine, 31*(1), 91-102.",
  "Song, Y., & Sheykhlouvand, M. (2024). A comparative analysis of high-intensity technique-specific intervals and short sprint interval training in taekwondo athletes: Effects on cardiorespiratory fitness and anaerobic power. *Journal of Sports Science and Medicine, 23*(3), 672-683. https://doi.org/10.52082/jssm.2024.672",
  "Teyhen, D. S., Shaffer, S. W., Lorenson, C. L., Halfpap, J. P., Donofry, D. F., Walker, M. J., . . . Childs, J. D. (2012). The Functional Movement Screen: A reliability study. *Journal of Orthopaedic & Sports Physical Therapy, 42*(6), 530-540. https://doi.org/10.2519/jospt.2012.3838",
];

// ---------------------------------------------------------------- 표
const T1 = table({
  widths: [1240, 1300, 1300, 760],
  header: ["변인", "실험군 (*n* = 13)", "통제군 (*n* = 12)", "*p*"],
  rows: [
    ["연령 (세)", "【 】", "【 】", "【 】"],
    ["신장 (cm)", "【 】", "【 】", "【 】"],
    ["체중 (kg)", "【 】", "【 】", "【 】"],
    ["선수 경력 (년)", "【 】", "【 】", "【 】"],
    ["성별 (남/여, 명)", "【 】", "【 】", "【 】"],
    ["VO~2~max (mL/kg/min)", "【 】", "【 】", "【 】"],
    ["FMS 총점 (점)", "【 】", "【 】", "【 】"],
  ],
})
const T2 = table({
  widths: [880, 2020, 1700],
  align: [AlignmentType.LEFT, AlignmentType.LEFT, AlignmentType.LEFT],
  header: ["구분", "운동 내용", "시간·운동량 / 강도"],
  rows: [
    ["준비운동", "가벼운 조깅, 동적 스트레칭, 고관절·발목 가동성 운동", "10분\nRPE 7–10"],
    ["본운동", "A: 스쿼트 점프, 점핑 런지, 버피, 스텝업\nB: 플랭크, 사이드 플랭크, 밴드 회전, 메디신볼 슬램\nC: 셔틀런, 방향 전환, 스텝과 발차기, 교대 돌려차기",
      "40분: 회기당 6–10종목, 종목당 30–40초, 3–4세트, 운동:휴식 1:1→2:1\nHRmax 85–95%, RPE 15–18"],
    ["정리운동", "저강도 걷기, 정적 스트레칭", "10분\nRPE 6–9"],
  ],
});
const T3 = table({
  widths: [640, 560, 800, 800, 680, 640, 480],
  size: 8,
  header: ["변인", "집단", "사전", "사후", "*t* (*p*)", "*F* (*p*)", "η~p~^2^"],
  rows: [
    ["VO~2~max\n(mL/kg/min)", "실험군", "【 】", "【 】", "【 】", "【 】", "【 】"],
    ["", "통제군", "【 】", "【 】", "【 】", "", ""],
    ["FMS 총점\n(점)", "실험군", "【 】", "【 】", "【 】", "【 】", "【 】"],
    ["", "통제군", "【 】", "【 】", "【 】", "", ""],
  ],
})

// ---------------------------------------------------------------- 문서 구성
const absPara = (label, text, en) => new Paragraph({
  children: [...runs(label + (en ? ": " : "  "), { size: 9, bold: true, font: en ? MJ : GT }), ...runs(text, { size: 9 })],
  alignment: AlignmentType.JUSTIFIED,
  spacing: { line: LS(150), before: 0, after: 60 },
});
const footer = new Footer({
  children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ children: [PageNumber.CURRENT], font: MJ, size: pt(9) })] })],
});
const page = { size: { width: PAGE_W, height: 16838 }, margin: { top: MARGIN, bottom: MARGIN, left: MARGIN, right: MARGIN, header: MM(15), footer: MM(10) } };

const firstSection = {
  properties: { page },
  footers: { default: footer },
  children: [
    new Paragraph({ children: runs(TITLE_KO, { size: 18, bold: true }), alignment: AlignmentType.CENTER, spacing: { line: LS(150), after: 240 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER, spacing: { line: LS(150) },
      children: runs("【국문 성명】^1^ · 【국문 성명】^2*^ · 【국문 성명】^3^ · 【국문 성명】^4^ · 【국문 성명】^5^", { font: GT, size: 10 }),
    }),
    center("Han-Kyung Lee^1^ · Ki-Seung Ham^2*^ · Yong-Yeon Park^3^ · Sang-Kun Hwang^4^ · Jun-Seok Lee^5^", { font: GT, size: 10, after: 80 }),
    center("^1^용인대학교 【학과, 직위】 · Yongin University, 【Department, Position】", { size: 8.5, ls: 130 }),
    center("^2^【소속, 직위】 · 【Affiliation, Position】", { size: 8.5, ls: 130 }),
    center("^3^【소속】 박사과정 · 【Affiliation】, Doctoral Student", { size: 8.5, ls: 130 }),
    center("^4^【소속】 박사과정 · 【Affiliation】, Doctoral Student", { size: 8.5, ls: 130 }),
    center("^5^【소속】 석사과정 · 【Affiliation】, Master's Student", { size: 8.5, ls: 130, after: 240 }),
    new Paragraph({ children: runs("국문요약", { font: GT, size: 10, bold: true }), spacing: { after: 80, line: LS(150) } }),
    ...ABS_KO.map(([l, t]) => absPara(l, t, false)),
    new Paragraph({
      children: [...runs("주제어: ", { font: GT, size: 9, bold: true }), ...runs(KW_KO, { size: 9 })],
      spacing: { before: 80, after: 200, line: LS(150) },
      border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: "000000", space: 6 } },
    }),
    new Paragraph({ children: runs("* 교신저자: 【국문 성명】(Ki-Seung Ham), 【소속】, 【이메일】", { size: 8 }), spacing: { line: LS(130) } }),
    new Paragraph({ children: runs("이 논문은 【연도】년 8월 국기원 국제학술대회(【공식 행사명】)에서 포스터로 발표한 연구의 자료를 재분석하여 작성한 것임. 【연구비 지원 기관·과제번호(해당 시)】 이해상충: 【없음】", { size: 8 }), spacing: { line: LS(130), after: 200 } }),
  ],
};

const bodyChildren = [
  h1("Ⅰ. 서론"), ...INTRO.map(body),
  h1("Ⅱ. 연구방법"),
  h2("1. 연구대상"), ...M1.map(body),
  tcap("표 1. 대상자의 일반적 특성과 사전 종속변인"), T1, tnote("값은 평균±표준편차(성별은 명). *p*, 독립표본 *t*검정(성별은 Fisher의 정확검정)."),
  h2("2. 운동 프로그램"), ...M2.map(body),
  tcap("표 2. 고강도 기능성 인터벌 트레이닝(H-IFIT) 구성"), T2,
  tnote("A–C는 동작 유형에 따른 묶음이며 실시 순서가 아님. HRmax, 최대심박수; RPE, 자각운동강도(【척도: Borg 6–20】)."),
  h2("3. 측정 항목 및 방법"),
  h3("1) 최대산소섭취량"), body(M3a),
  h3("2) 기능적 움직임 검사"), body(M3b),
  h2("4. 자료처리"), body(M4),
  h1("Ⅲ. 연구결과"),
  h2("1. 사전 동질성"), body(R0),
  h2("2. 최대산소섭취량"), body(R1),
  h2("3. FMS 총점"), body(R2),
  tcap("표 3. 집단별 최대산소섭취량과 FMS 총점의 변화"), T3,
  tnote("값은 평균±표준편차. *t*, 집단별 대응표본 *t*검정(사전−사후); *F*, 2(집단)×2(시기) 반복측정 분산분석의 집단×시기 상호작용(자유도 1, 【df】); η~p~^2^, 부분 에타제곱."),
  h1("Ⅳ. 논의"), ...DISC.map(body),
  h1("Ⅴ. 결론"), ...CONC.map(body),
  blank(80),
  new Paragraph({
    children: runs("【AI 활용 고지: 사용한 도구명·버전과 사용 범위(예: 원고 구조화·문장 다듬기·문헌 검색·통계 검산), 저자의 최종 검증 사실을 기재(국기원 연구윤리지침)】", { size: 9 }),
    spacing: { before: 120, after: 120, line: LS(150) },
  }),
  new Paragraph({ children: runs("참고문헌", { font: GT, size: 13, bold: true }), alignment: AlignmentType.CENTER, keepNext: true, spacing: { before: 300, after: 160, line: LS(150) } }),
  ...REFS.map(ref),
];

const bodySection = {
  properties: { type: SectionType.CONTINUOUS, page, column: { count: 2, space: COL_GAP, equalWidth: true } },
  footers: { default: footer },
  children: bodyChildren,
};

const enSection = {
  properties: { type: SectionType.NEXT_PAGE, page },
  footers: { default: footer },
  children: [
    new Paragraph({ children: runs(TITLE_EN, { size: 13, bold: true }), alignment: AlignmentType.CENTER, spacing: { line: LS(150), after: 200 } }),
    center("Han-Kyung Lee^1^ · Ki-Seung Ham^2*^ · Yong-Yeon Park^3^ · Sang-Kun Hwang^4^ · Jun-Seok Lee^5^", { font: GT, size: 10, after: 60 }),
    center("^1^Yongin University · ^2–5^【Affiliations】", { size: 8.5, ls: 130, after: 240 }),
    new Paragraph({ children: runs("ABSTRACT", { font: GT, size: 10, bold: true }), alignment: AlignmentType.CENTER, spacing: { after: 120, line: LS(150) } }),
    ...ABS_EN.map(([l, t]) => absPara(l, t, true)),
    new Paragraph({
      children: [...runs("Keywords: ", { size: 9, bold: true }), ...runs(KW_EN, { size: 9 })],
      spacing: { before: 80, line: LS(150) },
    }),
  ],
};

const doc = new Document({
  creator: "", title: TITLE_KO,
  styles: { default: { document: { run: { font: MJ, size: pt(10) } } } },
  sections: [firstSection, bodySection, enSection],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log("saved", OUT); });
