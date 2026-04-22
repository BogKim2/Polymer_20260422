/**
 * 견적서 DOCX 생성 - 깔끔한-견적서-양식-엑셀.xlsx 구조 기반
 * 동일한 레이아웃을 Word 표 구조로 재현
 */
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, BorderStyle, WidthType, ShadingType, VerticalAlign,
  Footer, PageNumber,
} = require("docx");
const fs = require("fs");

// ── 상수 ─────────────────────────────────────────────────────────────────────
// A4, 좌우 여백 0.5인치 → 사용 가능 폭 = 11906 - 1440 = 10466 DXA
const TW = 10466;

const THIN  = { style: BorderStyle.SINGLE, size: 4,  color: "888888" };
const MED   = { style: BorderStyle.SINGLE, size: 12, color: "333333" };
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: "FFFFFF" };

// 컬럼 비율 (48컬럼 → 6개 논리 컬럼)
// 품목명:단위:수량:단가:공급가액:세액 = 13:4:3:8:11:9
const COLS = (() => {
  const total = 48;
  const parts = [13, 4, 3, 8, 11, 9];
  return parts.map(p => Math.floor(TW * p / total));
})();
// 마지막 컬럼 보정
COLS[5] = TW - COLS.slice(0, 5).reduce((a, b) => a + b, 0);

function font(size = 18, bold = false, color = "000000") {
  return { font: "맑은 고딕", size, bold, color };
}

function bdr(l, r, t, b) {
  return {
    left: l, right: r, top: t, bottom: b,
    insideHorizontal: NONE, insideVertical: NONE,
  };
}

function tc(text, opts = {}) {
  const {
    sz = 18, bold = false, color = "000000",
    halign = "center", valign = VerticalAlign.CENTER,
    colspan = 1, width = null,
    borders = bdr(THIN, THIN, THIN, THIN),
    margins = { top: 60, bottom: 60, left: 80, right: 80 },
    wrap = false,
  } = opts;
  return new TableCell({
    columnSpan: colspan,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    borders,
    verticalAlign: valign,
    margins,
    children: [new Paragraph({
      alignment: halign === "center" ? AlignmentType.CENTER
               : halign === "right"  ? AlignmentType.RIGHT
               : AlignmentType.LEFT,
      children: [new TextRun({ text: String(text ?? ""), ...font(sz, bold, color) })],
    })],
  });
}

function row(height, cells) {
  return new TableRow({
    height: { value: height, rule: "exact" },
    children: cells,
  });
}

// ── 타이틀 행 ──────────────────────────────────────────────────────────────────
const titleRow = row(720, [
  tc("견  적  서", {
    sz: 36, bold: true, colspan: 6, width: TW,
    borders: bdr(MED, MED, MED, MED),
    margins: { top: 120, bottom: 120, left: 200, right: 200 },
  }),
]);

// ── 견적번호 행 ────────────────────────────────────────────────────────────────
const [W0, W1, W2, W3, W4, W5] = COLS;
const noRow = row(280, [
  tc("NO.", { sz: 18, width: W0, borders: bdr(MED, THIN, THIN, THIN) }),
  tc("", { sz: 18, colspan: 5, width: TW - W0, borders: bdr(THIN, MED, THIN, MED) }),
]);

// ── 날짜 행 ───────────────────────────────────────────────────────────────────
const dateRow = row(260, [
  tc("    년      월      일", {
    sz: 18, halign: "left", width: TW,
    colspan: 6,
    borders: bdr(MED, MED, THIN, THIN),
  }),
]);

// ── 수신처 행 ─────────────────────────────────────────────────────────────────
const recipientRow = row(340, [
  new TableCell({
    columnSpan: 6,
    width: { size: TW, type: WidthType.DXA },
    borders: bdr(MED, MED, THIN, MED),
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [
        new TextRun({ text: "                                    ", ...font(20, true) }),
        new TextRun({ text: "  귀  중", ...font(20, true) }),
      ],
    })],
  }),
]);

// 안내문
const guideRow = row(260, [
  tc("아래와 같이 견적 드립니다.", {
    sz: 18, halign: "left", colspan: 6, width: TW,
    borders: bdr(MED, MED, THIN, THIN),
  }),
]);

// ── 공급자 정보 행 ─────────────────────────────────────────────────────────────
// W공급자 라벨 + 5개 필드
const SUP_LABEL_W = Math.floor(TW * 0.06);
const SUP_FL_W    = Math.floor(TW * 0.16);  // field label
const SUP_FV1_W   = Math.floor(TW * 0.25);  // value 1
const SUP_FL2_W   = Math.floor(TW * 0.14);  // field label 2
const SUP_FV2_W   = TW - SUP_LABEL_W - SUP_FL_W - SUP_FV1_W - SUP_FL2_W;

function supRow(label1, val1, label2, val2, rowspan, isTop, isBot, isSeal) {
  const t = isTop ? MED : THIN;
  const b = isBot ? MED : THIN;
  const cells = [];
  if (rowspan) {
    cells.push(new TableCell({
      rowSpan: rowspan,
      width: { size: SUP_LABEL_W, type: WidthType.DXA },
      borders: bdr(MED, THIN, MED, MED),
      verticalAlign: VerticalAlign.CENTER,
      margins: { top: 60, bottom: 60, left: 80, right: 80 },
      children: [new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: "공  급  자", ...font(17, true) })],
      })],
    }));
  }
  cells.push(
    tc(label1, { sz: 17, width: SUP_FL_W, borders: bdr(THIN, THIN, t, b) }),
    tc(val1,   { sz: 17, halign: "left", width: SUP_FV1_W, borders: bdr(THIN, THIN, t, b) }),
  );
  if (label2 !== undefined) {
    cells.push(
      tc(label2, { sz: 17, width: SUP_FL2_W, borders: bdr(THIN, THIN, t, b) }),
      tc(val2,   { sz: 17, halign: "left",
                   width: TW - SUP_LABEL_W - SUP_FL_W - SUP_FV1_W - SUP_FL2_W,
                   borders: bdr(THIN, MED, t, b) }),
    );
  } else {
    cells.push(
      tc(val1 + " ", {
        sz: 17, halign: "left",
        colspan: 2,
        width: SUP_FL2_W + (TW - SUP_LABEL_W - SUP_FL_W - SUP_FV1_W - SUP_FL2_W),
        borders: bdr(THIN, MED, t, b),
      }),
    );
    // fix: label2/val2 undefined → just extend val1 col
    cells.pop(); cells.pop(); // remove last 2
    cells.push(
      tc(val1, {
        sz: 17, halign: "left", wrap: true,
        width: SUP_FL2_W + (TW - SUP_LABEL_W - SUP_FL_W - SUP_FV1_W - SUP_FL2_W),
        borders: bdr(THIN, MED, t, b),
        colspan: 2,
      }),
    );
  }
  return row(300, cells);
}

// 공급자 5행
const sup1 = new TableRow({ height: { value: 300, rule: "exact" }, children: [
  new TableCell({
    rowSpan: 5,
    width: { size: SUP_LABEL_W, type: WidthType.DXA },
    borders: bdr(MED, THIN, MED, MED),
    verticalAlign: VerticalAlign.CENTER,
    margins: { top: 60, bottom: 60, left: 60, right: 60 },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text: "공\n\n급\n\n자", ...font(17, true) })] })],
  }),
  tc("상  호  명",   { sz: 17, width: SUP_FL_W,  borders: bdr(THIN,THIN,MED,THIN) }),
  tc("(주) 케이에스모듈테크", { sz: 17, halign:"left", width: SUP_FV1_W, borders: bdr(THIN,THIN,MED,THIN) }),
  tc("사업자번호",   { sz: 17, width: SUP_FL2_W, borders: bdr(THIN,THIN,MED,THIN) }),
  tc("808-81-03318", { sz: 17, halign:"left", width: TW-SUP_LABEL_W-SUP_FL_W-SUP_FV1_W-SUP_FL2_W, borders: bdr(THIN,MED,MED,THIN) }),
]});

const sup2 = new TableRow({ height: { value: 300, rule: "exact" }, children: [
  tc("상호(대리점)", { sz: 17, width: SUP_FL_W,  borders: bdr(THIN,THIN,THIN,THIN) }),
  tc("",             { sz: 17, halign:"left", width: SUP_FV1_W, borders: bdr(THIN,THIN,THIN,THIN) }),
  tc("대    표",     { sz: 17, width: SUP_FL2_W, borders: bdr(THIN,THIN,THIN,THIN) }),
  tc("김 복 기  (인)", { sz: 17, halign:"left", width: TW-SUP_LABEL_W-SUP_FL_W-SUP_FV1_W-SUP_FL2_W, borders: bdr(THIN,MED,THIN,THIN) }),
]});

const sup3 = new TableRow({ height: { value: 300, rule: "exact" }, children: [
  tc("주       소", { sz: 17, width: SUP_FL_W, borders: bdr(THIN,THIN,THIN,THIN) }),
  new TableCell({
    columnSpan: 3,
    width: { size: TW - SUP_LABEL_W - SUP_FL_W, type: WidthType.DXA },
    borders: bdr(THIN, MED, THIN, THIN),
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [
      new TextRun({ text: "부산시 금정구 부산대학로 63번길 2, 부산대학교 물리학과(장전동) 307호", ...font(16) }),
    ]})],
  }),
]});

const sup4 = new TableRow({ height: { value: 300, rule: "exact" }, children: [
  tc("담       당", { sz: 17, width: SUP_FL_W,  borders: bdr(THIN,THIN,THIN,THIN) }),
  tc("",             { sz: 17, halign:"left", width: SUP_FV1_W, borders: bdr(THIN,THIN,THIN,THIN) }),
  tc("전    화",     { sz: 17, width: SUP_FL2_W, borders: bdr(THIN,THIN,THIN,THIN) }),
  tc("",             { sz: 17, halign:"left", width: TW-SUP_LABEL_W-SUP_FL_W-SUP_FV1_W-SUP_FL2_W, borders: bdr(THIN,MED,THIN,THIN) }),
]});

const sup5 = new TableRow({ height: { value: 300, rule: "exact" }, children: [
  tc("업       태", { sz: 17, width: SUP_FL_W, borders: bdr(THIN,THIN,THIN,MED) }),
  new TableCell({
    columnSpan: 3,
    width: { size: TW - SUP_LABEL_W - SUP_FL_W, type: WidthType.DXA },
    borders: bdr(THIN, MED, THIN, MED),
    margins: { top: 60, bottom: 60, left: 80, right: 80 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({ alignment: AlignmentType.LEFT, children: [
      new TextRun({ text: "제조업, 판매업, 설계업, 통신판매업", ...font(16) }),
    ]})],
  }),
]});

// ── 합계금액 행 ────────────────────────────────────────────────────────────────
const totalRow = row(480, [
  tc("(공급가액 + 세액)", { sz: 18, bold: true, width: COLS[0] + COLS[1], colspan: 2, borders: bdr(MED,THIN,MED,MED) }),
  new TableCell({
    columnSpan: 4,
    width: { size: TW - COLS[0] - COLS[1], type: WidthType.DXA },
    borders: bdr(THIN, MED, MED, MED),
    margins: { top: 80, bottom: 80, left: 120, right: 120 },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      alignment: AlignmentType.LEFT,
      children: [new TextRun({ text: "일금                                            원정 (부가세 포함)", ...font(20, true, "CC0000") })],
    })],
  }),
]);

// ── 품목 테이블 헤더 ────────────────────────────────────────────────────────────
const [C0, C1, C2, C3, C4, C5] = COLS;
const hdrRow = row(380, [
  tc("품  목  명",    { sz: 18, bold:true, width:C0, borders:bdr(MED,THIN,MED,MED) }),
  tc("단 위",         { sz: 18, bold:true, width:C1, borders:bdr(THIN,THIN,MED,MED) }),
  tc("수 량",         { sz: 18, bold:true, width:C2, borders:bdr(THIN,THIN,MED,MED) }),
  tc("단 가",         { sz: 18, bold:true, width:C3, borders:bdr(THIN,THIN,MED,MED) }),
  tc("공  급  가  액",{ sz: 18, bold:true, width:C4, borders:bdr(MED,MED,MED,MED) }),
  tc("세 액",         { sz: 18, bold:true, width:C5, borders:bdr(MED,MED,MED,MED) }),
]);

// ── 품목 행 (15행) ─────────────────────────────────────────────────────────────
function itemRow(n) {
  return row(300, [
    tc("", { sz: 17, halign:"left", width:C0, borders:bdr(MED,THIN,THIN,THIN) }),
    tc("", { sz: 17, width:C1, borders:bdr(THIN,THIN,THIN,THIN) }),
    tc("", { sz: 17, width:C2, borders:bdr(THIN,THIN,THIN,THIN) }),
    tc("", { sz: 17, halign:"right", width:C3, borders:bdr(THIN,THIN,THIN,THIN) }),
    tc("", { sz: 17, halign:"right", width:C4, borders:bdr(MED,MED,THIN,THIN) }),
    tc("", { sz: 17, halign:"right", width:C5, borders:bdr(MED,MED,THIN,THIN) }),
  ]);
}
const itemRows = Array.from({ length: 15 }, (_, i) => itemRow(i + 1));

// ── 합계 행 ───────────────────────────────────────────────────────────────────
const sumRow = row(420, [
  tc("합       계", { sz: 18, bold:true, colspan:4,
    width: C0+C1+C2+C3, borders: bdr(MED,THIN,MED,MED) }),
  tc("", { sz: 18, bold:true, halign:"right", width:C4, borders:bdr(MED,MED,MED,MED) }),
  tc("", { sz: 18, bold:true, halign:"right", width:C5, borders:bdr(MED,MED,MED,MED) }),
]);

// ── 주의사항 행 ────────────────────────────────────────────────────────────────
const noteRow1 = row(260, [
  tc("※ 상기 금액은 부가세(VAT 10%) 포함 금액이며, 납품 조건 및 결제 방법은 협의 후 결정합니다.",
     { sz: 16, halign:"left", colspan:6, width:TW, borders:bdr(NONE,NONE,THIN,NONE) }),
]);
const noteRow2 = row(260, [
  tc("※ 제품 사양 및 가격은 사전 예고 없이 변경될 수 있습니다. 문의사항은 담당자에게 연락 주시기 바랍니다.",
     { sz: 16, halign:"left", colspan:6, width:TW, borders:bdr(NONE,NONE,NONE,NONE) }),
]);

// ── 메인 테이블 ────────────────────────────────────────────────────────────────
const mainTable = new Table({
  width: { size: TW, type: WidthType.DXA },
  columnWidths: COLS,
  rows: [
    titleRow,
    noRow,
    dateRow,
    recipientRow,
    guideRow,
    sup1, sup2, sup3, sup4, sup5,
    totalRow,
    hdrRow,
    ...itemRows,
    sumRow,
    noteRow1,
    noteRow2,
  ],
});

// ── 문서 생성 ──────────────────────────────────────────────────────────────────
const doc = new Document({
  styles: {
    default: { document: { run: { font: "맑은 고딕", size: 18 } } },
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 },
        margin: { top: 720, right: 720, bottom: 720, left: 720 },
      },
    },
    children: [mainTable],
  }],
});

Packer.toBuffer(doc).then(buf => {
  const out = "E:/venture/proposal/JBLab/2nd/1quotation/견적서_케이에스모듈테크.docx";
  fs.writeFileSync(out, buf);
  console.log("저장 완료:", out);
}).catch(e => { console.error(e); process.exit(1); });
