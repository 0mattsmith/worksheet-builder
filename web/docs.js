/*
 * Worksheet Builder – reading scheme-of-work files in the browser.
 * Word, PowerPoint, Excel and text files are turned into plain text here.
 * PDFs and photos are passed to the AI service as they are (Claude and Gemini read them directly).
 */
(function (global) {
  'use strict';

  const MAX_BYTES = 20 * 1024 * 1024;
  const IMAGE_TYPES = { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', webp: 'image/webp', gif: 'image/gif' };

  const ext = (name) => (String(name).toLowerCase().match(/\.([a-z0-9]+)$/) || [])[1] || '';
  const decoder = () => new TextDecoder('utf-8');

  function readBuffer(file) {
    if (file.arrayBuffer) return file.arrayBuffer();
    return new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = () => rej(r.error); r.readAsArrayBuffer(file); });
  }

  function toBase64(buf) {
    const bytes = new Uint8Array(buf);
    let bin = '';
    for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
    return btoa(bin);
  }

  function unzip(buf) {
    if (!global.fflate) throw new Error('The unzip library did not load.');
    return global.fflate.unzipSync(new Uint8Array(buf));
  }
  const xml = (bytes) => new DOMParser().parseFromString(decoder().decode(bytes), 'application/xml');
  const byLocal = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));

  // ----- Word -----
  function paraText(p) {
    let t = '';
    const walk = (n) => {
      n.childNodes.forEach((c) => {
        if (c.nodeType !== 1) return;
        const ln = c.localName;
        if (ln === 't') t += c.textContent;
        else if (ln === 'tab') t += '\t';
        else if (ln === 'br' || ln === 'cr') t += '\n';
        else if (ln !== 'pPr' && ln !== 'rPr' && ln !== 'del' && ln !== 'instrText') walk(c);
      });
    };
    walk(p);
    return t;
  }
  function paraPrefix(p) {
    const style = byLocal(p, 'pStyle')[0];
    const s = style ? style.getAttribute('w:val') || '' : '';
    const m = s.match(/heading\s?(\d)|Titl/i);
    if (m) return m[1] ? '#'.repeat(Math.min(3, +m[1])) + ' ' : '# ';
    return byLocal(p, 'numPr').length ? '• ' : '';
  }
  function docxText(files) {
    const main = files['word/document.xml'];
    if (!main) throw new Error('This does not look like a Word document.');
    const body = byLocal(xml(main), 'body')[0];
    const out = [];
    const block = (node) => {
      node.childNodes.forEach((c) => {
        if (c.nodeType !== 1) return;
        if (c.localName === 'p') { const t = paraText(c).trim(); if (t) out.push(paraPrefix(c) + t); }
        else if (c.localName === 'tbl') {
          byLocal(c, 'tr').forEach((tr) => {
            if (tr.parentNode !== c) return; // nested tables are read inside their cell
            const cells = Array.from(tr.childNodes).filter((x) => x.localName === 'tc')
              .map((tc) => byLocal(tc, 'p').map(paraText).map((s) => s.trim()).filter(Boolean).join(' / '));
            if (cells.some(Boolean)) out.push('| ' + cells.join(' | ') + ' |');
          });
          out.push('');
        } else if (c.localName === 'sdt' || c.localName === 'sdtContent' || c.localName === 'customXml') block(c);
      });
    };
    block(body);
    return out.join('\n');
  }

  // ----- PowerPoint -----
  function pptxText(files) {
    const slides = Object.keys(files).filter((k) => /^ppt\/slides\/slide\d+\.xml$/.test(k))
      .sort((a, b) => +a.match(/(\d+)\.xml/)[1] - +b.match(/(\d+)\.xml/)[1]);
    if (!slides.length) throw new Error('No slides found in this presentation.');
    return slides.map((k, i) => {
      const d = xml(files[k]);
      const paras = byLocal(d, 'p').map((p) => byLocal(p, 't').map((t) => t.textContent).join('')).filter((s) => s.trim());
      return '## Slide ' + (i + 1) + '\n' + paras.join('\n');
    }).join('\n\n');
  }

  // ----- Excel -----
  function colIndex(ref) {
    const letters = (ref.match(/^[A-Z]+/) || ['A'])[0];
    let n = 0;
    for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
    return n - 1;
  }
  function xlsxText(files) {
    const shared = files['xl/sharedStrings.xml'] ? byLocal(xml(files['xl/sharedStrings.xml']), 'si').map((si) => byLocal(si, 't').map((t) => t.textContent).join('')) : [];
    const names = files['xl/workbook.xml'] ? byLocal(xml(files['xl/workbook.xml']), 'sheet').map((s) => s.getAttribute('name')) : [];
    const sheets = Object.keys(files).filter((k) => /^xl\/worksheets\/sheet\d+\.xml$/.test(k))
      .sort((a, b) => +a.match(/(\d+)\.xml/)[1] - +b.match(/(\d+)\.xml/)[1]);
    return sheets.map((k, i) => {
      const rows = byLocal(xml(files[k]), 'row').map((row) => {
        const cells = [];
        byLocal(row, 'c').forEach((c) => {
          const v = byLocal(c, 'v')[0];
          const is = byLocal(c, 't');
          let val = '';
          if (c.getAttribute('t') === 's' && v) val = shared[+v.textContent] || '';
          else if (c.getAttribute('t') === 'inlineStr') val = is.map((t) => t.textContent).join('');
          else if (v) val = v.textContent;
          cells[colIndex(c.getAttribute('r') || 'A1')] = String(val).replace(/\s+/g, ' ').trim();
        });
        return cells;
      }).filter((r) => r.some(Boolean));
      return '## Sheet: ' + (names[i] || 'Sheet ' + (i + 1)) + '\n' + rows.map((r) => '| ' + Array.from(r, (x) => x || '').join(' | ') + ' |').join('\n');
    }).join('\n\n');
  }

  /**
   * Read a dropped / chosen file.
   * Returns { name, size, kind: 'text' | 'pdf' | 'image', text?, base64?, mime?, summary }
   */
  async function readFile(file) {
    const e = ext(file.name);
    if (file.size > MAX_BYTES) throw new Error('That file is over 20 MB. Please use a smaller file or save just the relevant pages.');
    if (e === 'doc' || e === 'ppt' || e === 'xls') throw new Error('Old Office format (.' + e + '). Open it and use File → Save As to save it as .' + e + 'x, then try again.');
    if (e === 'pages' || e === 'key' || e === 'numbers') throw new Error('Please export this file as PDF or Word first.');
    const buf = await readBuffer(file);
    const base = { name: file.name, size: file.size };
    if (e === 'pdf') return Object.assign(base, { kind: 'pdf', mime: 'application/pdf', base64: toBase64(buf), summary: 'PDF – sent to the AI as it is' });
    if (IMAGE_TYPES[e]) return Object.assign(base, { kind: 'image', mime: IMAGE_TYPES[e], base64: toBase64(buf), summary: 'Image – the AI will read it' });
    let text;
    if (e === 'docx' || e === 'docm' || e === 'dotx') text = docxText(unzip(buf));
    else if (e === 'pptx') text = pptxText(unzip(buf));
    else if (e === 'xlsx' || e === 'xlsm') text = xlsxText(unzip(buf));
    else if (['txt', 'md', 'csv', 'tsv', 'rtf', 'html', 'htm', 'json'].includes(e) || !e) {
      text = decoder().decode(buf);
      if (e === 'html' || e === 'htm') text = new DOMParser().parseFromString(text, 'text/html').body.innerText || '';
      if (e === 'rtf') text = text.replace(/\\par[d]?/g, '\n').replace(/\{\\\*[^}]*\}|\\[a-z]+-?\d* ?|[{}]/gi, '');
    } else throw new Error('Unsupported file type (.' + e + '). Use Word, PDF, PowerPoint, Excel, text or an image.');
    text = String(text || '').replace(/\n{3,}/g, '\n\n').trim();
    if (!text) throw new Error('No text could be found in that file.');
    const words = text.split(/\s+/).length;
    return Object.assign(base, { kind: 'text', text, summary: words.toLocaleString('en-GB') + ' words found' });
  }

  global.WSDocs = { readFile, docxText, pptxText, xlsxText };
})(typeof window !== 'undefined' ? window : globalThis);
