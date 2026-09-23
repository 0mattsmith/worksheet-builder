/*
 * Worksheet Builder – document operations.
 * Everything that reads or changes the Word document lives here, so it can be called
 * either directly (task pane) or on behalf of the pop-out window (which can't use Word APIs).
 * Every operation takes plain data and returns plain data (JSON-safe).
 */
(function (global) {
  'use strict';
  const WS = global.WSOoxml;
  const settings = () => Office.context.document.settings;
  const pkg = (xml, theme) => WS.wrapPackage(xml, theme);
  const isSet = (v) => Office.context.requirements.isSetSupported('WordApi', v);
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const clone = (o) => JSON.parse(JSON.stringify(o));

  async function findAncestor(ctx, range, match) {
    let cc = range.parentContentControlOrNullObject;
    cc.load('tag,isNullObject');
    await ctx.sync();
    for (let depth = 0; depth < 6 && !cc.isNullObject; depth++) {
      if (match(cc.tag || '')) return cc;
      const parent = cc.parentContentControlOrNullObject;
      parent.load('tag,isNullObject');
      await ctx.sync();
      cc = parent;
    }
    return null;
  }
  const isBlockTag = (t) => t.startsWith('ws-q:') || t === 'ws-header' || t === 'ws-key';

  function saveSettings(partial) {
    Object.keys(partial).forEach((k) => settings().set('ws.' + k, partial[k]));
    return new Promise((res) => settings().saveAsync(() => res(true)));
  }

  const ops = {
    // Everything the interface needs to start: saved worksheet data + what this Word supports.
    async state() {
      const s = settings();
      return {
        questions: s.get('ws.questions') || {},
        lessons: s.get('ws.lessons') || {},
        theme: s.get('ws.theme') || null,
        mode: s.get('ws.mode') || 'student',
        header: s.get('ws.header') || null,
        sets: { '1.3': isSet('1.3'), '1.5': isSet('1.5'), '1.9': isSet('1.9') },
      };
    },

    async save(partial) { return saveSettings(partial || {}); },

    // Insert a chunk of worksheet at the cursor (or end of document), never inside another question.
    async insertXml({ xml, where, theme }) {
      await Word.run(async (ctx) => {
        let range;
        if (where === 'end') {
          range = ctx.document.body.insertOoxml(pkg(xml, theme), 'End');
        } else {
          let sel = ctx.document.getSelection();
          sel.load('isEmpty');
          await ctx.sync();
          const anc = await findAncestor(ctx, sel, isBlockTag);
          if (anc) {
            const p = anc.insertParagraph('', 'After');
            range = p.insertOoxml(pkg(xml, theme), 'Replace');
          } else {
            if (!sel.isEmpty) sel = sel.getRange('End');
            range = sel.insertOoxml(pkg(xml, theme), 'Replace');
          }
        }
        await ctx.sync();
        try { range.getRange('End').select(); await ctx.sync(); } catch (e) { /* cursor placement is optional */ }
      });
      return true;
    },

    // Renumber questions (restarting at each lesson), update total marks, refresh the answer key.
    async refresh({ questions, lessons, theme, header }) {
      questions = questions || {};
      lessons = lessons || {};
      let result = { count: 0, total: 0, lessons: 0, header: false };
      await Word.run(async (ctx) => {
        const ccs = ctx.document.body.contentControls;
        ccs.load('items/tag');
        await ctx.sync();
        const seen = new Set();
        const entries = [];
        let keyCc = null, hasHeader = false, lessonCount = 0;
        const totals = [];
        ccs.items.forEach((cc) => {
          const tag = cc.tag || '';
          if (tag.startsWith('ws-q:')) {
            let id = tag.slice(5);
            if (seen.has(id) && questions[id]) { // copied question → give it its own identity
              const nid = newId();
              questions[nid] = Object.assign(clone(questions[id]), { id: nid });
              cc.tag = 'ws-q:' + nid;
              id = nid;
            }
            seen.add(id);
            const nums = cc.contentControls.getByTag('ws-num');
            nums.load('items');
            entries.push({ id, nums });
          } else if (tag.startsWith('ws-lesson:')) { entries.push({ lesson: tag.slice(10) }); lessonCount++; }
          else if (tag === 'ws-total') totals.push(cc);
          else if (tag === 'ws-key') keyCc = cc;
          else if (tag === 'ws-header') hasHeader = true;
        });
        await ctx.sync();
        let total = 0, n = 0, count = 0, group = null;
        const keyList = [];
        entries.forEach((e) => {
          if (e.lesson !== undefined) { n = 0; group = lessons[e.lesson] ? lessons[e.lesson].title : 'Lesson'; return; }
          n++; count++;
          if (e.nums.items[0]) e.nums.items[0].insertText(WS.questionLabel(n, theme), 'Replace');
          const q = questions[e.id];
          if (q) { total += q.marks || 0; keyList.push({ num: n, q, group }); }
        });
        totals.forEach((t) => t.insertText(String(total), 'Replace'));
        if (keyCc) keyCc.insertOoxml(pkg(WS.renderAnswerKey(keyList, theme, header, { wrap: false }), theme), 'Replace');
        await ctx.sync();
        result = { count, total, lessons: lessonCount, header: hasHeader };
      });
      await saveSettings({ questions, lessons });
      result.questions = questions;
      return result;
    },

    // Re-draw every question, lesson page, section and the header (teacher/student copy, style changes).
    async rerender({ questions, lessons, theme, mode, header }) {
      questions = questions || {};
      lessons = lessons || {};
      await Word.run(async (ctx) => {
        const ccs = ctx.document.body.contentControls;
        ccs.load('items/tag,items/text');
        await ctx.sync();
        let n = 0;
        ccs.items.forEach((cc) => {
          const tag = cc.tag || '';
          if (tag.startsWith('ws-q:')) {
            n++;
            const q = questions[tag.slice(5)];
            if (q) cc.insertOoxml(pkg(WS.renderQuestion(q, n, theme, { show: mode === 'teacher', wrap: false }), theme), 'Replace');
          } else if (tag.startsWith('ws-lesson:')) {
            const l = lessons[tag.slice(10)];
            if (l) cc.insertOoxml(pkg(WS.renderLesson(l, theme, Object.assign({}, l.opts, { wrap: false })), theme), 'Replace');
          } else if (tag === 'ws-sec') {
            const text = (cc.text || '').trim();
            if (text) cc.insertOoxml(pkg(WS.renderSection(text, theme, { wrap: false }), theme), 'Replace');
          } else if (tag === 'ws-header') {
            if (header) cc.insertOoxml(pkg(WS.renderHeader(header, theme, 0, { wrap: false }), theme), 'Replace');
          }
        });
        await ctx.sync();
      });
      return ops.refresh({ questions, lessons, theme, header });
    },

    async replaceQuestion({ q, theme, mode }) {
      await Word.run(async (ctx) => {
        const coll = ctx.document.body.contentControls.getByTag('ws-q:' + q.id);
        coll.load('items');
        await ctx.sync();
        if (!coll.items.length) throw new Error('Could not find that question in the document any more.');
        coll.items.forEach((cc) => cc.insertOoxml(pkg(WS.renderQuestion(q, 1, theme, { show: mode === 'teacher', wrap: false }), theme), 'Replace'));
        await ctx.sync();
      });
      return true;
    },

    async selectedQuestionId() {
      let id = null;
      await Word.run(async (ctx) => {
        const cc = await findAncestor(ctx, ctx.document.getSelection(), (t) => t.startsWith('ws-q:'));
        if (cc) id = cc.tag.slice(5);
      });
      return id;
    },

    async deleteSelectedQuestion() {
      let found = false;
      await Word.run(async (ctx) => {
        const cc = await findAncestor(ctx, ctx.document.getSelection(), (t) => t.startsWith('ws-q:'));
        if (cc) { cc.delete(false); found = true; await ctx.sync(); }
      });
      return found;
    },

    // Update the WS styles already in the document (Word 2021+/365/web).
    async applyStyles({ theme }) {
      if (!isSet('1.5')) return false;
      await Word.run(async (ctx) => {
        const defs = WS.styleDefs(theme);
        const styles = ctx.document.getStyles();
        const found = defs.map((d) => { const s = styles.getByNameOrNullObject(d.name); s.load('isNullObject'); return s; });
        await ctx.sync();
        defs.forEach((d, i) => {
          let s = found[i];
          if (s.isNullObject) s = ctx.document.addStyle(d.name, d.type === 'paragraph' ? 'Paragraph' : 'Character');
          s.font.name = d.font === 'heading' ? theme.headingFont : theme.bodyFont;
          if (d.size) s.font.size = d.size;
          s.font.bold = !!d.bold;
          s.font.italic = !!d.italic;
          if (d.color) s.font.color = '#' + d.color;
          if (d.type === 'paragraph' && d.spacing) {
            s.paragraphFormat.spaceBefore = d.spacing.before / 20;
            s.paragraphFormat.spaceAfter = d.spacing.after / 20;
          }
        });
        await ctx.sync();
      });
      return true;
    },

    async insertHeader({ h, theme }) {
      await saveSettings({ header: h });
      await Word.run(async (ctx) => {
        const existing = ctx.document.body.contentControls.getByTag('ws-header');
        existing.load('items');
        await ctx.sync();
        if (existing.items.length) existing.items[0].insertOoxml(pkg(WS.renderHeader(h, theme, 0, { wrap: false }), theme), 'Replace');
        else ctx.document.body.insertOoxml(pkg(WS.renderHeader(h, theme, 0), theme), 'Start');
        if (h.footer) {
          const footer = ctx.document.sections.getFirst().getFooter('Primary');
          footer.insertOoxml(pkg(WS.renderFooter([h.subject, h.title].filter(Boolean).join(' – ')), theme), 'Replace');
        }
        await ctx.sync();
      });
      return true;
    },

    // Add an (empty) answer key at the end if there isn't one; refresh() fills it in.
    async ensureKey({ theme, header }) {
      await Word.run(async (ctx) => {
        const existing = ctx.document.body.contentControls.getByTag('ws-key');
        existing.load('items');
        await ctx.sync();
        if (!existing.items.length) ctx.document.body.insertOoxml(pkg(WS.renderAnswerKey([], theme, header), theme), 'End');
        await ctx.sync();
      });
      return true;
    },

    async insertCheckbox({ theme }) {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        if (theme.checkboxMode === 'clickable' && isSet('1.9')) {
          sel.getRange('Start').insertContentControl('CheckBox');
        } else if (theme.checkboxMode === 'clickable') {
          sel.getRange('Start').insertOoxml(pkg(WS.para({}, WS.checkbox(false, 'box', theme) + WS.run(' ')), theme), 'Replace');
        } else {
          sel.insertText('☐ ', 'Start');
        }
        await ctx.sync();
      });
      return true;
    },

    async selectionToChecklist({ theme }) {
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        const paras = sel.paragraphs;
        paras.load('items/text');
        await ctx.sync();
        const lines = paras.items.map((p) => p.text.replace(/^[☐☒○●•\-*]\s*/, '').trim()).filter(Boolean);
        if (!lines.length) throw new Error('Select the lines you want to turn into a checklist first.');
        sel.insertOoxml(pkg(WS.renderChecklist('', lines, theme), theme), 'Replace');
        await ctx.sync();
      });
      return true;
    },

    async selectionText() {
      let text = '';
      await Word.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load('text');
        await ctx.sync();
        text = sel.text || '';
      });
      return text;
    },

    async questionIds() {
      const ids = [];
      await Word.run(async (ctx) => {
        const ccs = ctx.document.body.contentControls;
        ccs.load('items/tag');
        await ctx.sync();
        ccs.items.forEach((cc) => { if ((cc.tag || '').startsWith('ws-q:')) ids.push(cc.tag.slice(5)); });
      });
      return ids;
    },
  };

  global.WSDocOps = ops;
})(typeof window !== 'undefined' ? window : globalThis);
