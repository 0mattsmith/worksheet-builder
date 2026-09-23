/*
 * Worksheet Builder – layout engine.
 * Turns worksheet data (questions, header, answer key…) into Word Open XML so that
 * every worksheet uses exactly the same styles, spacing and tick boxes.
 * Works in the browser (window.WSOoxml) and in Node (module.exports) for testing.
 */
(function (global) {
  'use strict';

  const DEFAULT_THEME = {
    bodyFont: 'Arial',
    headingFont: 'Arial',
    baseSize: 12,            // pt
    accent: '1F4E79',        // hex without #
    questionGap: 14,         // pt of space before each question
    lineHeight: 'normal',    // narrow | normal | wide
    lineStyle: 'single',     // single | dotted | dashed
    checkboxMode: 'clickable', // clickable | print
    numberFormat: '1.',      // 1. | 1) | Q1 | Q1.
    marksFormat: '[n]',      // [n] | (n marks) | [n marks]
    showMarks: true,
    showHints: true,         // "Tick one box." etc.
    headerStyle: 'classic',  // classic | banner | boxed
    sectionStyle: 'underline', // underline | band | plain
    answerArea: 'lines',     // lines | box | grid  (space for written answers)
    questionStyle: 'plain',  // plain | tinted
    numberBadge: false,      // question numbers in a coloured badge
  };

  const LINE_HEIGHTS = { narrow: 400, normal: 500, wide: 640 }; // twips
  const INDENT = 567;   // 1 cm – question text indent
  const OPT_TAB = 1134; // 2 cm – option letter
  const OPT_TEXT = 1701; // 3 cm – option text

  const NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" ' +
    'xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" ' +
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="w14 w15"';

  // ---------- small helpers ----------
  const esc = (s) => String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const letter = (i) => String.fromCharCode(65 + i);
  const hp = (pt) => Math.round(pt * 2); // half-points

  function tint(hex, amount) {
    const n = parseInt(hex, 16);
    const mix = (c) => Math.round(c + (255 - c) * amount);
    const r = mix((n >> 16) & 255), g = mix((n >> 8) & 255), b = mix(n & 255);
    return ((1 << 24) + (r << 16) + (g << 8) + b).toString(16).slice(1).toUpperCase();
  }

  function numberLabel(n, fmt) {
    switch (fmt) {
      case '1)': return n + ')';
      case 'Q1': return 'Q' + n;
      case 'Q1.': return 'Q' + n + '.';
      default: return n + '.';
    }
  }

  // Label as it appears in the document (badges get a little padding).
  function questionLabel(n, theme) {
    const l = numberLabel(n, theme.numberFormat);
    return theme.numberBadge ? '\u00A0' + l.replace(/\.$/, '') + '\u00A0' : l;
  }

  function marksLabel(m, fmt) {
    const s = m === 1 ? '' : 's';
    if (fmt === '(n marks)') return '(' + m + ' mark' + s + ')';
    if (fmt === '[n marks]') return '[' + m + ' mark' + s + ']';
    return '[' + m + ']';
  }

  // Run properties in schema order.
  function rPr(o) {
    if (!o) return '';
    let x = '';
    if (o.style) x += '<w:rStyle w:val="' + o.style + '"/>';
    if (o.font) x += '<w:rFonts w:ascii="' + esc(o.font) + '" w:eastAsia="' + esc(o.font) + '" w:hAnsi="' + esc(o.font) + '" w:cs="' + esc(o.font) + '"' + (o.hint ? ' w:hint="' + o.hint + '"' : '') + '/>';
    if (o.b) x += '<w:b/><w:bCs/>';
    if (o.i) x += '<w:i/><w:iCs/>';
    if (o.color) x += '<w:color w:val="' + o.color + '"/>';
    if (o.sz) x += '<w:sz w:val="' + o.sz + '"/><w:szCs w:val="' + o.sz + '"/>';
    if (o.u) x += '<w:u w:val="single"/>';
    if (o.shd) x += '<w:shd w:val="clear" w:color="auto" w:fill="' + o.shd + '"/>';
    return x ? '<w:rPr>' + x + '</w:rPr>' : '';
  }

  // Paragraph properties in schema order.
  function pPr(o) {
    let x = '';
    if (o.style) x += '<w:pStyle w:val="' + o.style + '"/>';
    if (o.keepNext) x += '<w:keepNext/>';
    if (o.keepLines) x += '<w:keepLines/>';
    if (o.pageBreakBefore) x += '<w:pageBreakBefore/>';
    if (o.bdr) x += '<w:pBdr>' + o.bdr + '</w:pBdr>';
    if (o.shd) x += '<w:shd w:val="clear" w:color="auto" w:fill="' + o.shd + '"/>';
    if (o.tabs) x += '<w:tabs>' + o.tabs.map((t) => '<w:tab w:val="' + (t.val || 'left') + '"' + (t.leader ? ' w:leader="' + t.leader + '"' : '') + ' w:pos="' + t.pos + '"/>').join('') + '</w:tabs>';
    if (o.spacing) x += '<w:spacing' + ['before', 'after', 'line', 'lineRule'].filter((k) => o.spacing[k] != null).map((k) => ' w:' + k + '="' + o.spacing[k] + '"').join('') + '/>';
    if (o.ind) x += '<w:ind' + ['left', 'right', 'hanging', 'firstLine'].filter((k) => o.ind[k] != null).map((k) => ' w:' + k + '="' + o.ind[k] + '"').join('') + '/>';
    if (o.jc) x += '<w:jc w:val="' + o.jc + '"/>';
    if (o.rPr) x += o.rPr;
    return x ? '<w:pPr>' + x + '</w:pPr>' : '';
  }

  const para = (props, content) => '<w:p>' + pPr(props || {}) + (content || '') + '</w:p>';
  const TAB = '<w:r><w:tab/></w:r>';
  const RIGHT_TAB = '<w:r><w:ptab w:relativeTo="margin" w:alignment="right" w:leader="none"/></w:r>';

  function run(text, props) {
    const parts = String(text == null ? '' : text).split('\n');
    const inner = parts.map((p, i) => (i ? '<w:br/>' : '') + (p ? '<w:t xml:space="preserve">' + esc(p) + '</w:t>' : '')).join('');
    return '<w:r>' + rPr(props) + inner + '</w:r>';
  }

  // **bold** and *italic* inline formatting
  function rich(text, base) {
    base = base || {};
    const out = [];
    const re = /(\*\*[^*]+\*\*|\*[^*\s][^*]*\*)/g;
    let last = 0, m;
    text = String(text || '');
    while ((m = re.exec(text))) {
      if (m.index > last) out.push(run(text.slice(last, m.index), base));
      const tok = m[0];
      if (tok.startsWith('**')) out.push(run(tok.slice(2, -2), Object.assign({}, base, { b: true })));
      else out.push(run(tok.slice(1, -1), Object.assign({}, base, { i: true })));
      last = re.lastIndex;
    }
    if (last < text.length) out.push(run(text.slice(last), base));
    return out.join('');
  }

  function sdtInline(tag, alias, content, hidden) {
    return '<w:sdt><w:sdtPr><w:alias w:val="' + esc(alias) + '"/><w:tag w:val="' + esc(tag) + '"/>' +
      (hidden ? '<w15:appearance w15:val="hidden"/>' : '') +
      '</w:sdtPr><w:sdtContent>' + content + '</w:sdtContent></w:sdt>';
  }
  const sdtBlock = sdtInline; // same markup, used around paragraphs

  // ---------- tick boxes ----------
  const GLYPHS = { box: ['2610', '2612'], bubble: ['25CB', '25CF'] };

  function checkbox(checked, kind, theme) {
    const g = GLYPHS[kind] || GLYPHS.box;
    const code = checked ? g[1] : g[0];
    const glyphRun = run(String.fromCharCode(parseInt(code, 16)), { font: 'MS Gothic', hint: 'eastAsia', sz: hp(theme.baseSize + 1) });
    if (theme.checkboxMode !== 'clickable') return glyphRun;
    return '<w:sdt><w:sdtPr>' + rPr({ font: 'MS Gothic', hint: 'eastAsia', sz: hp(theme.baseSize + 1) }) +
      '<w:alias w:val="Tick box"/><w:tag w:val="ws-box"/>' +
      '<w14:checkbox><w14:checked w14:val="' + (checked ? 1 : 0) + '"/>' +
      '<w14:checkedState w14:val="' + g[1] + '" w14:font="MS Gothic"/>' +
      '<w14:uncheckedState w14:val="' + g[0] + '" w14:font="MS Gothic"/></w14:checkbox>' +
      '</w:sdtPr><w:sdtContent>' + glyphRun + '</w:sdtContent></w:sdt>';
  }

  // ---------- styles ----------
  function styleDefs(theme) {
    const b = theme.baseSize, a = theme.accent;
    return [
      { id: 'WSTitle', name: 'WS Title', type: 'paragraph', font: 'heading', size: b + 10, bold: true, color: a, spacing: { before: 0, after: 60 }, keepNext: true },
      { id: 'WSSubtitle', name: 'WS Subtitle', type: 'paragraph', font: 'body', size: b, color: '595959', spacing: { before: 0, after: 160 } },
      { id: 'WSSection', name: 'WS Section', type: 'paragraph', font: 'heading', size: b + 3, bold: true, color: a, spacing: { before: 320, after: 120 }, keepNext: true,
        bdr: '<w:bottom w:val="single" w:sz="8" w:space="2" w:color="' + a + '"/>' },
      { id: 'WSInstructions', name: 'WS Instructions', type: 'paragraph', font: 'body', size: b, color: '262626', spacing: { before: 120, after: 200 },
        bdr: ['top', 'left', 'bottom', 'right'].map((s) => '<w:' + s + ' w:val="single" w:sz="4" w:space="4" w:color="' + a + '"/>').join(''), shd: tint(a, 0.9), ind: { left: 85, right: 85 } },
      { id: 'WSQuestion', name: 'WS Question', type: 'paragraph', font: 'body', size: b, color: '000000', spacing: { before: theme.questionGap * 20, after: 100 }, keepNext: true,
        ind: { left: INDENT, hanging: INDENT }, tabs: [{ pos: INDENT }] },
      { id: 'WSOption', name: 'WS Option', type: 'paragraph', font: 'body', size: b, color: '000000', spacing: { before: 0, after: 80 },
        ind: { left: OPT_TEXT, hanging: OPT_TEXT - INDENT }, tabs: [{ pos: OPT_TAB }, { pos: OPT_TEXT }] },
      { id: 'WSLine', name: 'WS Answer Line', type: 'paragraph', font: 'body', size: b, color: '000000', spacing: { before: 0, after: 0 }, ind: { left: INDENT } },
      { id: 'WSCell', name: 'WS Table Text', type: 'paragraph', font: 'body', size: b, color: '000000', spacing: { before: 60, after: 60 } },
      { id: 'WSWordBank', name: 'WS Word Bank', type: 'paragraph', font: 'body', size: b, color: '000000', spacing: { before: 60, after: 120 }, jc: 'center',
        bdr: ['top', 'left', 'bottom', 'right'].map((s) => '<w:' + s + ' w:val="dashed" w:sz="4" w:space="4" w:color="7F7F7F"/>').join(''), ind: { left: INDENT + 85, right: 85 } },
      { id: 'WSPassage', name: 'WS Passage', type: 'paragraph', font: 'body', size: b, color: '000000', spacing: { before: 0, after: 140, line: 300 } },
      { id: 'WSKey', name: 'WS Answer Key', type: 'paragraph', font: 'body', size: b, color: '000000', spacing: { before: 0, after: 100 }, ind: { left: INDENT, hanging: INDENT }, tabs: [{ pos: INDENT }] },
      { id: 'WSFooter', name: 'WS Footer', type: 'paragraph', font: 'body', size: b - 3, color: '7F7F7F', spacing: { before: 0, after: 0 } },
      { id: 'WSNumber', name: 'WS Question Number', type: 'character', font: 'body', bold: true, color: a },
      { id: 'WSMarks', name: 'WS Marks', type: 'character', font: 'body', size: b - 1, color: '7F7F7F' },
      { id: 'WSHint', name: 'WS Hint', type: 'character', font: 'body', size: b - 1, italic: true, color: '595959' },
      { id: 'WSTeacher', name: 'WS Teacher Answer', type: 'character', font: 'body', bold: true, color: 'C00000' },
    ];
  }

  const fontOf = (d, theme) => (d.font === 'heading' ? theme.headingFont : theme.bodyFont);

  function stylesXml(theme) {
    const s = styleDefs(theme).map((d) => {
      const r = rPr({ font: fontOf(d, theme), b: d.bold, i: d.italic, color: d.color, sz: d.size ? hp(d.size) : null });
      const p = d.type === 'paragraph' ? pPr({ keepNext: d.keepNext, bdr: d.bdr, shd: d.shd, tabs: d.tabs, spacing: Object.assign({ line: 264, lineRule: 'auto' }, d.spacing), ind: d.ind, jc: d.jc }) : '';
      return '<w:style w:type="' + d.type + '" w:customStyle="1" w:styleId="' + d.id + '"><w:name w:val="' + d.name + '"/>' +
        (d.type === 'paragraph' ? '<w:next w:val="' + d.id + '"/>' : '') + '<w:qFormat/>' + p + r + '</w:style>';
    }).join('');
    return '<w:styles ' + NS + '>' + s + '</w:styles>';
  }

  // ---------- question rendering ----------
  function lineProps(theme) {
    const style = { single: 'single', dotted: 'dotted', dashed: 'dashed' }[theme.lineStyle] || 'single';
    const border = (side) => '<w:' + side + ' w:val="' + style + '" w:sz="4" w:space="1" w:color="8C8C8C"/>';
    return { style: 'WSLine', bdr: border('bottom') + border('between'), spacing: { before: 0, after: 0, line: LINE_HEIGHTS[theme.lineHeight] || 500, lineRule: 'exact' }, ind: { left: INDENT } };
  }

  function answerBox(count, theme, answer) {
    const bd = ['top', 'left', 'bottom', 'right'].map((s) => '<w:' + s + ' w:val="single" w:sz="6" w:space="4" w:color="8C8C8C"/>').join('');
    const out = [];
    const n = Math.max(1, count | 0);
    for (let i = 0; i < n; i++) {
      out.push(para({ style: 'WSLine', keepLines: true, keepNext: i < n - 1, bdr: bd, spacing: { before: 0, after: 0, line: LINE_HEIGHTS[theme.lineHeight] || 500, lineRule: 'exact' }, ind: { left: INDENT + 85, right: 85 } },
        i === 0 && answer ? run(answer, { style: 'WSTeacher' }) : ''));
    }
    return out.join('') + para({ style: 'WSCell', spacing: { before: 0, after: 0, line: 160, lineRule: 'exact' } }, '');
  }

  function squaredGrid(count, theme, answer) {
    const cols = 28, size = 300; // 5.3 mm squares
    const rows = Math.max(4, Math.round((count || 3) * 1.7));
    const b = (s) => '<w:' + s + ' w:val="single" w:sz="2" w:space="0" w:color="BFBFBF"/>';
    const cellP = para({ style: 'WSCell', spacing: { before: 0, after: 0, line: 240, lineRule: 'auto' }, rPr: '<w:rPr><w:sz w:val="4"/><w:szCs w:val="4"/></w:rPr>' }, '');
    const cell = '<w:tc><w:tcPr><w:tcW w:w="' + size + '" w:type="dxa"/></w:tcPr>' + cellP + '</w:tc>';
    let rowsXml = '';
    for (let r = 0; r < rows; r++) rowsXml += '<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="' + size + '" w:hRule="exact"/></w:trPr>' + cell.repeat(cols) + '</w:tr>';
    const pre = answer ? para({ style: 'WSLine', ind: { left: INDENT }, spacing: { before: 0, after: 60 } }, run(answer, { style: 'WSTeacher' })) : '';
    return pre + '<w:tbl><w:tblPr><w:tblW w:w="' + cols * size + '" w:type="dxa"/><w:tblInd w:w="' + INDENT + '" w:type="dxa"/>' +
      '<w:tblBorders>' + ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map(b).join('') + '</w:tblBorders>' +
      '<w:tblLayout w:type="fixed"/><w:tblCellMar><w:left w:w="0" w:type="dxa"/><w:right w:w="0" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr>' +
      '<w:tblGrid>' + ('<w:gridCol w:w="' + size + '"/>').repeat(cols) + '</w:tblGrid>' + rowsXml + '</w:tbl>' +
      para({ style: 'WSCell', spacing: { before: 0, after: 0, line: 160, lineRule: 'exact' } }, '');
  }

  function answerSpace(area, count, theme, answer) {
    if (area === 'box') return answerBox(count, theme, answer);
    if (area === 'grid') return squaredGrid(count, theme, answer);
    return writingLines(count, theme, answer);
  }

  function writingLines(count, theme, answer) {
    const out = [];
    const n = Math.max(1, count | 0);
    for (let i = 0; i < n; i++) {
      out.push(para(lineProps(theme), i === 0 && answer ? run(answer, { style: 'WSTeacher' }) : ''));
    }
    return out.join('');
  }

  function blanksRuns(text, show) {
    const out = [];
    const re = /\[([^\]]+)\]/g;
    let last = 0, m;
    while ((m = re.exec(text))) {
      if (m.index > last) out.push(rich(text.slice(last, m.index)));
      const ans = m[1].trim();
      if (show) out.push(run(ans, { style: 'WSTeacher', u: true }));
      else out.push(run('_'.repeat(Math.max(10, Math.ceil(ans.length * 1.6)))));
      last = re.lastIndex;
    }
    if (last < text.length) out.push(rich(text.slice(last)));
    return out.join('');
  }

  function blanksAnswers(text) {
    const res = [];
    const re = /\[([^\]]+)\]/g; let m;
    while ((m = re.exec(String(text || '')))) res.push(m[1].trim());
    return res;
  }

  const HINTS = { mc: 'Tick one box.', multi: 'Tick all that apply.', tf: 'Tick true or false.', match: 'Match each item on the left to one on the right.' };

  function questionPara(q, num, theme, show, keepNext) {
    const numProps = theme.numberBadge ? { style: 'WSNumber', color: 'FFFFFF', shd: theme.accent } : { style: 'WSNumber' };
    const numRun = sdtInline('ws-num', 'Question number', run(questionLabel(num, theme), numProps), true);
    const stem = q.type === 'blanks' ? blanksRuns(q.text, show) : rich(q.text);
    const hint = theme.showHints && HINTS[q.type] ? run(' ' + HINTS[q.type], { style: 'WSHint' }) : '';
    const marks = theme.showMarks && q.marks > 0 ? RIGHT_TAB + run(marksLabel(q.marks, theme.marksFormat), { style: 'WSMarks' }) : '';
    const qp = { style: 'WSQuestion', keepNext: keepNext, keepLines: true };
    if (theme.questionStyle === 'tinted') {
      qp.shd = tint(theme.accent, 0.9);
      qp.bdr = '<w:left w:val="single" w:sz="18" w:space="4" w:color="' + theme.accent + '"/>';
      qp.spacing = { before: theme.questionGap * 20, after: 120 };
    }
    return para(qp, numRun + TAB + stem + hint + marks);
  }

  function orderOf(q, len) {
    const o = Array.isArray(q.order) && q.order.length === len ? q.order : null;
    return o || Array.from({ length: len }, (_, i) => i);
  }

  function matchTable(q, show, theme) {
    const pairs = (q.pairs || []).filter((p) => p.left || p.right);
    const order = orderOf(q, pairs.length); // order[j] = index of pair shown in right-hand row j
    const widths = [3700, 1300, 3400];
    let keep = true;
    const cell = (w, content, jc) => '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/><w:vAlign w:val="center"/></w:tcPr>' +
      para({ style: 'WSCell', jc: jc, keepNext: keep }, content) + '</w:tc>';
    const rows = pairs.map((p, i) => {
      keep = i < pairs.length - 1;
      const j = order.indexOf(i);
      const right = pairs[order[i]];
      return '<w:tr><w:trPr><w:cantSplit/></w:trPr>' +
        cell(widths[0], run(i + 1 + '.  ', { b: true }) + rich(p.left)) +
        cell(widths[1], show ? run(letter(j), { style: 'WSTeacher' }) : run('_____'), 'center') +
        cell(widths[2], run(letter(i) + '.  ', { b: true }) + rich(right.right)) + '</w:tr>';
    }).join('');
    const none = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => '<w:' + s + ' w:val="nil"/>').join('');
    return '<w:tbl><w:tblPr><w:tblW w:w="' + widths.reduce((a, b) => a + b) + '" w:type="dxa"/><w:tblInd w:w="' + INDENT + '" w:type="dxa"/>' +
      '<w:tblBorders>' + none + '</w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="0000"/></w:tblPr>' +
      '<w:tblGrid>' + widths.map((w) => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid>' + rows + '</w:tbl>' +
      para({ style: 'WSCell', spacing: { before: 0, after: 0, line: 120, lineRule: 'exact' } }, '');
  }

  /**
   * Render one question.
   * q: { id, type: mc|multi|tf|written|blanks|match, text, marks, options[{text,correct}], answer, lines, pairs[{left,right}], order[], wordBank }
   * opts: { show: teacher answers visible, wrap: wrap in a tagged content control }
   */
  function renderQuestion(q, num, theme, opts) {
    theme = Object.assign({}, DEFAULT_THEME, theme);
    opts = opts || {};
    const show = !!opts.show;
    let body = '';
    const t = q.type;
    if (t === 'mc' || t === 'multi') {
      const opts2 = q.options || [];
      const order = orderOf(q, opts2.length);
      body += questionPara(q, num, theme, show, true);
      order.forEach((oi, j) => {
        const o = opts2[oi];
        const last = j === order.length - 1;
        body += para({ style: 'WSOption', keepNext: !last, keepLines: true },
          checkbox(show && o.correct, t === 'mc' ? 'bubble' : 'box', theme) + TAB + run(letter(j), { b: true }) + TAB + rich(o.text));
      });
    } else if (t === 'tf') {
      body += questionPara(q, num, theme, show, true);
      const isTrue = String(q.answer).toLowerCase() === 'true';
      body += para({ style: 'WSOption', keepLines: true, tabs: [{ pos: OPT_TAB }, { pos: 2835 }, { pos: 3402 }] },
        checkbox(show && isTrue, 'bubble', theme) + TAB + run('True') + TAB + checkbox(show && !isTrue, 'bubble', theme) + TAB + run('False'));
    } else if (t === 'blanks') {
      const bank = q.wordBank ? blanksAnswers(q.text) : [];
      body += questionPara(q, num, theme, show, bank.length > 0);
      if (bank.length) {
        const order = orderOf(q, bank.length);
        body += para({ style: 'WSWordBank', keepLines: true }, run('Word bank:   ', { b: true }) + run(order.map((i) => bank[i]).join('      ')));
      }
    } else if (t === 'match') {
      body += questionPara(q, num, theme, show, true);
      body += matchTable(q, show, theme);
    } else { // written
      body += questionPara(q, num, theme, show, true);
      body += answerSpace(q.area || theme.answerArea, q.lines || 3, theme, show ? q.answer : '');
    }
    if (opts.wrap === false) return body;
    const alias = 'Question: ' + String(q.text || '').replace(/[\[\]*]/g, '').slice(0, 40);
    return sdtBlock('ws-q:' + q.id, alias, body);
  }

  function answerText(q) {
    switch (q.type) {
      case 'mc': case 'multi': {
        const order = orderOf(q, (q.options || []).length);
        const right = order.map((oi, j) => ({ o: q.options[oi], j })).filter((x) => x.o.correct);
        return right.map((x) => letter(x.j) + ' – ' + x.o.text.replace(/\*/g, '')).join(';  ') || '—';
      }
      case 'tf': return String(q.answer).toLowerCase() === 'true' ? 'True' : 'False';
      case 'blanks': return blanksAnswers(q.text).join(';  ');
      case 'match': {
        const pairs = (q.pairs || []).filter((p) => p.left || p.right);
        const order = orderOf(q, pairs.length);
        return pairs.map((_, i) => (i + 1) + ' – ' + letter(order.indexOf(i))).join(',   ');
      }
      default: return q.answer || '(Teacher to mark)';
    }
  }

  // ---------- other blocks ----------
  function sectionPara(text, theme, extra) {
    theme = Object.assign({}, DEFAULT_THEME, theme);
    const p = Object.assign({ style: 'WSSection' }, extra || {});
    if (theme.sectionStyle === 'band') {
      p.bdr = '<w:left w:val="single" w:sz="36" w:space="6" w:color="' + theme.accent + '"/><w:bottom w:val="nil"/>';
      p.shd = tint(theme.accent, 0.86);
    } else if (theme.sectionStyle === 'plain') {
      p.bdr = '<w:bottom w:val="nil"/>';
    } else if (theme.sectionStyle === 'solid') {
      p.bdr = ['top', 'left', 'bottom', 'right'].map((s) => '<w:' + s + ' w:val="single" w:sz="4" w:space="3" w:color="' + theme.accent + '"/>').join('');
      p.shd = theme.accent;
      p.ind = { left: 57, right: 57 };
      return para(p, rich(text, { color: 'FFFFFF' }));
    }
    return para(p, rich(text));
  }
  function renderSection(text, theme, opts) {
    const x = sectionPara(text, theme);
    return opts && opts.wrap === false ? x : sdtBlock('ws-sec', 'Section heading', x, true);
  }

  function renderChecklist(title, items, theme) {
    theme = Object.assign({}, DEFAULT_THEME, theme);
    let x = title ? sectionPara(title, theme) : '';
    items.filter((s) => s.trim()).forEach((it, i, arr) => {
      x += para({ style: 'WSOption', keepNext: i < arr.length - 1, ind: { left: OPT_TAB, hanging: OPT_TAB - INDENT }, tabs: [{ pos: OPT_TAB }] },
        checkbox(false, 'box', theme) + TAB + rich(it.trim()));
    });
    return x;
  }

  /** l: { id, title, objectives[], vocabulary[] }  opts: { pageBreak, objectives, vocab, fields[] } */
  function renderLesson(l, theme, opts) {
    theme = Object.assign({}, DEFAULT_THEME, theme);
    opts = opts || {};
    const a = theme.accent;
    let x = para({ style: 'WSTitle', pageBreakBefore: !!opts.pageBreak, keepNext: true, spacing: { before: 0, after: 80 },
      bdr: '<w:bottom w:val="single" w:sz="12" w:space="4" w:color="' + a + '"/>' }, rich(l.title || 'Lesson', { sz: hp(theme.baseSize + 6) }));
    if (opts.fields && opts.fields.length) x += studentInfoTable(opts.fields, theme);
    if (opts.objectives !== false && l.objectives && l.objectives.length) {
      x += para({ style: 'WSPassage', keepNext: true, spacing: { before: 120, after: 40 } }, run('Learning objectives', { b: true, color: a }));
      l.objectives.forEach((o) => {
        x += para({ style: 'WSPassage', keepNext: true, spacing: { before: 0, after: 40 }, ind: { left: 284, hanging: 284 }, tabs: [{ pos: 284 }] }, run('•', { color: a }) + TAB + rich(o));
      });
    }
    if (opts.vocab !== false && l.vocabulary && l.vocabulary.length) {
      x += para({ style: 'WSWordBank', ind: { left: 85, right: 85 }, spacing: { before: 120, after: 60 } }, run('Key vocabulary:   ', { b: true }) + run(l.vocabulary.join('   ·   ')));
    }
    return opts.wrap === false ? x : sdtBlock('ws-lesson:' + (l.id || ''), 'Lesson', x);
  }

  function renderText(text, theme) {
    return String(text || '').split(/\n/).map((t) => para({ style: 'WSPassage' }, rich(t))).join('');
  }

  function renderLines(count, theme) {
    return writingLines(count, Object.assign({}, DEFAULT_THEME, theme), '');
  }

  function bandBdr(c) {
    return ['top', 'left', 'bottom', 'right'].map((s) => '<w:' + s + ' w:val="single" w:sz="4" w:space="6" w:color="' + c + '"/>').join('');
  }

  function studentInfoTable(fields, theme) {
    if (!fields.length) return '';
    const total = 9000;
    const k = ((theme && theme.baseSize) || 12) / 12 * (/comic|century|verdana/i.test((theme && theme.bodyFont) || '') ? 1.2 : 1);
    const labelW = fields.map((f) => Math.round((360 + f.length * 170) * k));
    const free = total - labelW.reduce((a, b) => a + b, 0);
    const weights = fields.map((f, i) => (i === 0 ? 2.2 : 1));
    const wsum = weights.reduce((a, b) => a + b, 0);
    const lineW = weights.map((w) => Math.floor((free * w) / wsum));
    const widths = [];
    fields.forEach((f, i) => { widths.push(labelW[i], lineW[i]); });
    const cells = fields.map((f, i) =>
      '<w:tc><w:tcPr><w:tcW w:w="' + labelW[i] + '" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr>' + para({ style: 'WSCell', jc: i ? 'right' : 'left' }, run(f + ':', { b: true })) + '</w:tc>' +
      '<w:tc><w:tcPr><w:tcW w:w="' + lineW[i] + '" w:type="dxa"/><w:tcBorders><w:bottom w:val="single" w:sz="6" w:space="0" w:color="404040"/></w:tcBorders><w:vAlign w:val="bottom"/></w:tcPr>' + para({ style: 'WSCell' }, '') + '</w:tc>'
    ).join('');
    const none = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => '<w:' + s + ' w:val="nil"/>').join('');
    return '<w:tbl><w:tblPr><w:tblW w:w="' + total + '" w:type="dxa"/><w:tblBorders>' + none + '</w:tblBorders><w:tblLayout w:type="fixed"/>' +
      '<w:tblCellMar><w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tblCellMar><w:tblLook w:val="0000"/></w:tblPr>' +
      '<w:tblGrid>' + widths.map((w) => '<w:gridCol w:w="' + w + '"/>').join('') + '</w:tblGrid><w:tr>' + cells + '</w:tr></w:tbl>';
  }

  /** h: { school, title, subject, group, topic, fields[], instructions, showTotal } */
  function renderHeader(h, theme, totalMarks, opts) {
    theme = Object.assign({}, DEFAULT_THEME, theme);
    let x = '';
    if (h.school) x += para({ style: 'WSSubtitle', spacing: { before: 0, after: 40 } }, run(h.school, { b: true, color: theme.accent }));
    const sub = [h.subject, h.group, h.topic].filter(Boolean).join('   ·   ');
    if (theme.headerStyle === 'banner') {
      const band = { shd: theme.accent, ind: { left: 113, right: 113 } };
      x += para(Object.assign({ style: 'WSTitle', spacing: { before: 0, after: 0 }, bdr: bandBdr(theme.accent) }, band), rich(h.title || 'Worksheet', { color: 'FFFFFF' }));
      x += para(Object.assign({ style: 'WSSubtitle', spacing: { before: 0, after: 200 }, bdr: bandBdr(theme.accent) }, band), run(sub || ' ', { color: 'FFFFFF' }));
    } else if (theme.headerStyle === 'boxed') {
      const bd = ['top', 'left', 'bottom', 'right'].map((s) => '<w:' + s + ' w:val="single" w:sz="12" w:space="6" w:color="' + theme.accent + '"/>').join('');
      const box = { bdr: bd, ind: { left: 142, right: 142 }, jc: 'center' };
      x += para(Object.assign({ style: 'WSTitle', spacing: { before: 120, after: 0 } }, box), rich(h.title || 'Worksheet'));
      x += para(Object.assign({ style: 'WSSubtitle', spacing: { before: 0, after: 240 } }, box), run(sub || ' '));
    } else if (theme.headerStyle === 'stripe') {
      const st = { bdr: '<w:left w:val="single" w:sz="48" w:space="8" w:color="' + theme.accent + '"/>', shd: tint(theme.accent, 0.88), ind: { left: 170, right: 57 } };
      x += para(Object.assign({ style: 'WSTitle', spacing: { before: 0, after: 0 } }, st), rich(h.title || 'Worksheet'));
      x += para(Object.assign({ style: 'WSSubtitle', spacing: { before: 0, after: 200 } }, st), run(sub || ' '));
    } else if (theme.headerStyle === 'minimal') {
      x += para({ style: 'WSTitle', spacing: { before: 0, after: 0 }, rPr: '' }, rich(h.title || 'Worksheet', { sz: hp(theme.baseSize + 6) }));
      x += para({ style: 'WSSubtitle', spacing: { before: 0, after: 200 }, bdr: '<w:bottom w:val="single" w:sz="4" w:space="4" w:color="' + theme.accent + '"/>' }, run(sub || ' '));
    } else {
      x += para({ style: 'WSTitle' }, rich(h.title || 'Worksheet'));
      if (sub) x += para({ style: 'WSSubtitle' }, run(sub));
    }
    x += studentInfoTable(h.fields || [], theme);
    if (h.instructions) x += para({ style: 'WSInstructions' }, run('Instructions:  ', { b: true }) + rich(h.instructions));
    if (h.showTotal) {
      x += para({ style: 'WSSubtitle', jc: 'right', spacing: { before: 120, after: 120 } },
        run('Total:  ', { b: true, color: '000000' }) + run('______ / ', { color: '000000' }) +
        sdtInline('ws-total', 'Total marks', run(String(totalMarks || 0), { b: true, color: '000000' }), true) + run(' marks', { color: '000000' }));
    } else if (!h.instructions) {
      x += para({ style: 'WSSubtitle', spacing: { before: 0, after: 120 } }, '');
    }
    return opts && opts.wrap === false ? x : sdtBlock('ws-header', 'Worksheet header', x);
  }

  function renderAnswerKey(list, theme, h, opts) {
    theme = Object.assign({}, DEFAULT_THEME, theme);
    let total = 0;
    let x = sectionPara('Answer key' + (h && h.title ? ' – ' + h.title : ''), theme, { pageBreakBefore: true });
    let group = null;
    list.forEach((item) => {
      const q = item.q;
      total += q.marks || 0;
      if (item.group && item.group !== group) {
        group = item.group;
        x += para({ style: 'WSKey', keepNext: true, spacing: { before: 200, after: 80 } }, run(group, { b: true, color: theme.accent }));
      }
      x += para({ style: 'WSKey', keepLines: true },
        run(numberLabel(item.num, theme.numberFormat), { style: 'WSNumber' }) + TAB + rich(answerText(q)) +
        (q.marks ? RIGHT_TAB + run(marksLabel(q.marks, theme.marksFormat), { style: 'WSMarks' }) : ''));
    });
    x += para({ style: 'WSKey', jc: 'right', spacing: { before: 200 } }, run('Total: ' + total + ' marks', { b: true }));
    return opts && opts.wrap === false ? x : sdtBlock('ws-key', 'Answer key', x);
  }

  function renderFooter(text) {
    return para({ style: 'WSFooter', tabs: [] }, run(text || '') + RIGHT_TAB + run('Page ') +
      '<w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple>');
  }

  // ---------- packaging ----------
  function wrapPackage(bodyXml, theme) {
    theme = Object.assign({}, DEFAULT_THEME, theme);
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">' +
      '<pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>' +
      '<w:document ' + NS + '><w:body>' + bodyXml + '</w:body></w:document>' +
      '</pkg:xmlData></pkg:part>' +
      '<pkg:part pkg:name="/word/styles.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"><pkg:xmlData>' +
      stylesXml(theme) +
      '</pkg:xmlData></pkg:part></pkg:package>';
  }

  // Suggested marks for a question (used when the teacher hasn't typed one).
  function autoMarks(q) {
    switch (q.type) {
      case 'multi': return Math.max(1, (q.options || []).filter((o) => o.correct).length);
      case 'blanks': return Math.max(1, blanksAnswers(q.text).length);
      case 'match': return Math.max(1, (q.pairs || []).filter((p) => p.left && p.right).length);
      case 'written': return Math.max(1, Math.round((q.lines || 3) / 2));
      default: return 1;
    }
  }

  function shuffle(n) {
    const a = Array.from({ length: n }, (_, i) => i);
    for (let i = n - 1; i > 0; i--) { const j = Math.floor(Math.random() * (i + 1)); [a[i], a[j]] = [a[j], a[i]]; }
    if (n > 1 && a.every((v, i) => v === i)) [a[0], a[1]] = [a[1], a[0]];
    return a;
  }

  /**
   * Quick build: turn plain text into questions.
   * Returns [{kind:'section', text}|{kind:'question', q}], errors[]
   */
  function parseQuick(src) {
    const items = [], errors = [];
    let cur = null;
    const marksRe = /\s*\((\d+)\s*marks?\)\s*$/i;
    const finish = () => {
      if (!cur) return;
      const q = cur;
      if (q.options.length) {
        const nCorrect = q.options.filter((o) => o.correct).length;
        q.type = q.forceType || (nCorrect > 1 ? 'multi' : 'mc');
        if (nCorrect === 0) errors.push('“' + q.text.slice(0, 40) + '” has no correct option (mark it with *).');
      } else if (q.pairs.length) q.type = 'match';
      else if (/\[[^\]]+\]/.test(q.text)) q.type = 'blanks';
      else if (/^(true|false|t|f)$/i.test(q.answer || '')) { q.type = 'tf'; q.answer = /^t/i.test(q.answer) ? 'true' : 'false'; }
      else q.type = 'written';
      if (q.type === 'written' && !q.lines) q.lines = q.marks ? Math.min(14, Math.max(2, q.marks * 2)) : 3;
      if (!q.marksGiven) q.marks = autoMarks(q);
      delete q.marksGiven;
      if (q.type === 'match') q.order = shuffle(q.pairs.length);
      if (q.type === 'blanks' && q.wordBank) q.order = shuffle(blanksAnswers(q.text).length);
      if ((q.type === 'mc' || q.type === 'multi') && q.shuffle) q.order = shuffle(q.options.length);
      delete q.forceType; delete q.shuffle;
      if (!q.text) errors.push('A question is missing its text.');
      items.push({ kind: 'question', q });
      cur = null;
    };
    String(src || '').split(/\r?\n/).forEach((raw) => {
      const line = raw.trim();
      if (!line) return;
      let m;
      if ((m = line.match(/^#+\s*(.+)$/)) || (m = line.match(/^section:\s*(.+)$/i))) { finish(); items.push({ kind: 'section', text: m[1].trim() }); return; }
      if ((m = line.match(/^={3,}\s*(.+)$/))) { finish(); items.push({ kind: 'lesson', lesson: { title: m[1].replace(/=+\s*$/, '').trim(), objectives: [], vocabulary: [] } }); return; }
      const lastItem = items[items.length - 1];
      if (!cur && lastItem && lastItem.kind === 'lesson') {
        if ((m = line.match(/^(?:LO|objective|learning objective)\s*:\s*(.+)$/i))) { lastItem.lesson.objectives.push(m[1].trim()); return; }
        if ((m = line.match(/^(?:vocab|vocabulary|key words?)\s*:\s*(.+)$/i))) { lastItem.lesson.vocabulary.push.apply(lastItem.lesson.vocabulary, m[1].split(/\s*[,;]\s*/).filter(Boolean)); return; }
      }
      if ((m = line.match(/^>\s?(.*)$/))) {
        finish();
        const last = items[items.length - 1];
        if (last && last.kind === 'text') last.text += '\n' + m[1]; else items.push({ kind: 'text', text: m[1] });
        return;
      }
      if ((m = line.match(/^\[\s?\]\s*(.+)$/))) {
        finish();
        const last = items[items.length - 1];
        if (last && last.kind === 'check') last.items.push(m[1].trim()); else items.push({ kind: 'check', items: [m[1].trim()] });
        return;
      }
      if ((m = line.match(/^(?:Q\d*[:.)]|\d+[.)])\s*(.+)$/i))) {
        finish();
        let text = m[1];
        let marks = 0;
        const mm = text.match(marksRe);
        if (mm) { marks = +mm[1]; text = text.replace(marksRe, ''); }
        cur = { text: text.trim(), marks, marksGiven: !!mm, options: [], pairs: [], answer: '', lines: 0 };
        return;
      }
      if (!cur) { errors.push('Ignored (no question above it): “' + line.slice(0, 50) + '”'); return; }
      if ((m = line.match(/^(?:A|answer)\s*:\s*(.*)$/i))) { cur.answer = m[1].trim(); return; }
      if ((m = line.match(/^lines\s*:\s*(\d+)/i))) { cur.lines = +m[1]; return; }
      if (/^(word ?bank|bank)\s*:\s*(yes|y|true|on)/i.test(line)) { cur.wordBank = true; return; }
      if (/^shuffle\s*:\s*(yes|y|true|on)/i.test(line)) { cur.shuffle = true; return; }
      if ((m = line.match(/^(?:space|area)\s*:\s*(lines|box|grid)/i))) { cur.area = m[1].toLowerCase(); return; }
      if ((m = line.match(/^type\s*:\s*(multi|mc)/i))) { cur.forceType = m[1].toLowerCase(); return; }
      if ((m = line.match(/^([*\-•+]|[a-h][).])\s+(.+)$/i))) {
        let text = m[2].trim();
        let correct = m[1] === '*' || m[1] === '+';
        if (/\s*(\*|✓|\(correct\))$/i.test(text)) { correct = true; text = text.replace(/\s*(\*|✓|\(correct\))$/i, ''); }
        cur.options.push({ text, correct });
        return;
      }
      if ((m = line.match(/^(.+?)\s*(?:=|->|→)\s*(.+)$/))) { cur.pairs.push({ left: m[1].trim(), right: m[2].trim() }); return; }
      cur.text += '\n' + line; // continuation of the question text
    });
    finish();
    return { items, errors };
  }

  const api = { DEFAULT_THEME, questionLabel, tint, styleDefs, stylesXml, renderQuestion, renderSection, renderChecklist, renderLines, renderHeader,
    renderAnswerKey, renderFooter, renderText, renderLesson, answerText, wrapPackage, numberLabel, marksLabel, autoMarks, shuffle, parseQuick, blanksAnswers, checkbox, para, run, NS };
  if (typeof module !== 'undefined' && module.exports) module.exports = api;
  else global.WSOoxml = api;
})(typeof window !== 'undefined' ? window : globalThis);
