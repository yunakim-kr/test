#!/usr/bin/env node
// 사용법: node build_agenda.js 입력.json 출력.docx
// 협의체 안건 샘플 서식(제목 박스, □ 소제목, ○ 개조식, 인용 박스, ⇨ 결론 박스, 참고 대비표)을 재현한다.
const fs = require('fs');
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, WidthType, ShadingType,
  BorderStyle, AlignmentType, VerticalAlign, Footer, PageNumber, PageBreak,
} = require('docx');

const F = { title: 'HY헤드라인M', head: 'HY견고딕', body: '휴먼명조', box: '굴림' };
let BODY = 30, SUB = 28, LINE = 380; // 입력 JSON의 compact:true이면 본문 14pt로 축소
const BLUE = '0000CC';
const PURPLE = 'E9DDF0';
const CONTENT_W = 9026; // A4 세로, 좌우 여백 1440

// ---- 인라인 강조: **텍스트** → 굵게(+ 본문에서는 파란색) ----
function runs(text, o) {
  const parts = String(text).split(/(\*\*[^*]+\*\*)/g).filter(s => s !== '');
  return parts.map(s => {
    const em = s.startsWith('**') && s.endsWith('**');
    const t = em ? s.slice(2, -2) : s;
    return new TextRun({
      text: t, font: o.font, size: o.size,
      bold: em ? true : o.bold,
      color: em ? (o.emColor === undefined ? BLUE : o.emColor) : o.color,
    });
  });
}
const para = (text, o = {}) => new Paragraph({
  alignment: o.align ?? AlignmentType.JUSTIFIED,
  keepNext: o.keepNext,
  spacing: { before: o.before ?? 0, after: o.after ?? 80, line: o.line ?? LINE },
  indent: o.indent,
  children: runs(text, { font: o.font ?? F.body, size: o.size ?? BODY, bold: o.bold, color: o.color, emColor: o.emColor }),
});

const none = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const thin = { style: BorderStyle.SINGLE, size: 6, color: '000000' };
const hair = { style: BorderStyle.SINGLE, size: 4, color: '000000' };
const B = (b) => ({ top: b, bottom: b, left: b, right: b });

function boxTable(children, { border, fill, width = CONTENT_W, margins } = {}) {
  return new Table({
    width: { size: width, type: WidthType.DXA }, columnWidths: [width], alignment: AlignmentType.CENTER,
    rows: [new TableRow({ cantSplit: true, children: [new TableCell({
      width: { size: width, type: WidthType.DXA }, borders: B(border || none),
      shading: fill ? { fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
      margins: margins || { top: 80, bottom: 80, left: 140, right: 140 },
      children,
    })] })],
  });
}
const gap = (n = 120) => new Paragraph({ spacing: { before: 0, after: n, line: 240 }, children: [] });

// ---- 블록 ----
function titleBox(text) {
  return [
    boxTable([para(text, { font: F.title, size: 40, bold: true, align: AlignmentType.CENTER, line: 420, after: 0 })],
      { border: { style: BorderStyle.SINGLE, size: 12, color: '000000' }, width: 8400, margins: { top: 100, bottom: 100, left: 140, right: 140 } }),
    gap(200),
  ];
}
const heading = (t) => para('□ ' + t, { font: F.head, size: 32, bold: true, align: AlignmentType.LEFT, before: 160, after: 100, keepNext: true, line: 380 });
const circle = (t, kn) => para('○ ' + t, { indent: { left: 480, hanging: 480 }, keepNext: kn });
const sub = (t, kn) => para('- ' + t, { size: SUB, indent: { left: 880, hanging: 400 }, keepNext: kn });
const note = (t) => para('※ ' + t, { size: 24, indent: { left: 480, hanging: 300 }, color: '333333', emColor: '000000' });

function quote(lines) {
  const ps = lines.map(l => {
    const [text, lv] = Array.isArray(l) ? l : [l, 0];
    return para(text, { font: F.box, size: 21, line: 300, after: 20, indent: { left: 120 + lv * 200, hanging: lv ? 0 : 0 } });
  });
  return [boxTable(ps, { border: hair, margins: { top: 70, bottom: 70, left: 120, right: 120 } }), gap(140)];
}
function summaryBox(text) {
  const lines = Array.isArray(text) ? text : [text];
  const ps = lines.map(t => para(t, { font: F.box, size: 27, bold: true, line: 360, after: 40, emColor: '000000' }));
  return [boxTable(ps, { border: { style: BorderStyle.SINGLE, size: 10, color: '1F3864' }, fill: 'EAF0F8', margins: { top: 110, bottom: 110, left: 170, right: 170 } }), gap(160)];
}
function conclusion(lines) {
  const ps = lines.map((t, i) => para((i === 0 ? '⇨ ' : '') + t, { font: F.box, size: 26, line: 360, after: 40, bold: false, indent: i === 0 ? { left: 340, hanging: 340 } : { left: 340 } }));
  return [gap(60), boxTable(ps, { fill: PURPLE, margins: { top: 100, bottom: 100, left: 160, right: 160 } }), gap(100)];
}

function table(spec) {
  const cols = spec.headers.length;
  const widths = spec.widths || Array.from({ length: cols }, (_, i) => Math.floor(CONTENT_W / cols) + (i === cols - 1 ? CONTENT_W - Math.floor(CONTENT_W / cols) * cols : 0));
  const cellParas = (v, o) => (Array.isArray(v) ? v : String(v).split('\n')).map(t =>
    para(t, { font: F.body, size: 21, line: 320, after: 0, align: o.align, emColor: '000000', bold: o.bold }));
  const mk = (v, w, o = {}) => new TableCell({
    width: { size: w, type: WidthType.DXA }, borders: B(hair), verticalAlign: o.v || VerticalAlign.TOP,
    shading: o.fill ? { fill: o.fill, type: ShadingType.CLEAR, color: 'auto' } : undefined,
    margins: { top: 80, bottom: 80, left: 100, right: 100 }, children: cellParas(v, o),
  });
  return new Table({
    width: { size: CONTENT_W, type: WidthType.DXA }, columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: spec.headers.map((h, i) => mk(h, widths[i], { fill: 'F2F2F2', align: AlignmentType.CENTER, bold: true, v: VerticalAlign.CENTER })) }),
      ...spec.rows.map(r => new TableRow({ cantSplit: !!spec.cantSplit, children: r.map((c, i) => mk(c, widths[i], { align: (spec.align && spec.align[i]) || AlignmentType.LEFT })) })),
    ],
  });
}

function renderBlock(b, kn) {
  switch (b.type) {
    case 'heading': return [heading(b.text)];
    case 'circle': return [circle(b.text, kn)];
    case 'sub': return [sub(b.text, kn)];
    case 'note': return [note(b.text)];
    case 'quote': return quote(b.lines);
    case 'conclusion': return conclusion(b.lines);
    case 'table': return [table(b), gap(120)];
    default: throw new Error('알 수 없는 블록 유형: ' + b.type);
  }
}

function agendaChildren(a, first) {
  const out = [];
  if (!first) out.push(new Paragraph({ children: [new PageBreak()] }));
  out.push(...titleBox(a.title));
  if (a.subtitle) out.push(para(a.subtitle, { font: F.box, size: 22, align: AlignmentType.CENTER, after: 140, line: 300, color: '444444' }));
  if (a.summary) out.push(...summaryBox(a.summary));
  (a.sections || []).forEach(s => {
    out.push(heading(s.heading));
    (s.blocks || []).forEach((b, i, arr) => out.push(...renderBlock(b, arr[i + 1] && arr[i + 1].type === 'conclusion')));
  });
  (a.appendix || []).forEach(ap => {
    out.push(new Paragraph({ children: [new PageBreak()] }));
    out.push(para(ap.title, { font: F.head, size: 32, bold: true, align: AlignmentType.LEFT, after: 160, line: 380 }));
    (ap.blocks || []).forEach(b => out.push(...renderBlock(b)));
  });
  return out;
}

const input = JSON.parse(fs.readFileSync(process.argv[2], 'utf8'));
if (input.compact) { BODY = 28; SUB = 26; LINE = 340; }
const children = [];
input.agendas.forEach((a, i) => children.push(...agendaChildren(a, i === 0)));

const doc = new Document({
  styles: { default: { document: { run: { font: F.body, size: 30 } } } },
  sections: [{
    properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: input.compact ? 1000 : 1300, bottom: input.compact ? 1000 : 1300, left: 1440, right: 1440 } } },
    footers: { default: new Footer({ children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [
      new TextRun({ text: '- ', font: F.box, size: 20 }),
      new TextRun({ children: [PageNumber.CURRENT], font: F.box, size: 20 }),
      new TextRun({ text: ' -', font: F.box, size: 20 }),
    ] })] }) },
    children,
  }],
});
Packer.toBuffer(doc).then(buf => { fs.writeFileSync(process.argv[3], buf); console.log('생성 완료: ' + process.argv[3]); });
