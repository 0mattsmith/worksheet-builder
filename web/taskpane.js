/* Worksheet Builder – task pane logic */
(function () {
  'use strict';
  const WS = window.WSOoxml;
  const $ = (id) => document.getElementById(id);
  const esc = (s) => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

  let inWord = false;
  const POPOUT = /[?&]window=1(&|$)/.test(location.search); // running in the pop-out window
  let bridge = null;       // pop-out window: link to the task pane
  let hostWin = null;      // task pane: the open pop-out window
  let apiSets = {};        // Word API versions supported
  let docHeader = null;    // header saved in the document
  let theme = null;        // current house style
  let questions = {};      // id -> question data (saved inside the document)
  let lessons = {};        // id -> lesson page data (lesson packs)
  let mode = 'student';    // student | teacher

  // ---------------- storage ----------------
  const local = {
    get(k, d) { try { const v = localStorage.getItem('ws.' + k); return v ? JSON.parse(v) : d; } catch (e) { return d; } },
    set(k, v) { try { localStorage.setItem('ws.' + k, JSON.stringify(v)); } catch (e) { /* ignore */ } },
  };
  // All document work goes through WSDocOps – directly in the task pane, via the task pane in the pop-out window.
  const canEdit = () => inWord || (POPOUT && !!bridge);
  const Doc = {
    call(op, args) {
      if (POPOUT) return bridge ? bridge.call(op, args || {}) : Promise.reject(new Error('This window has lost its link to Word. Close it and open it again from the Worksheet Builder panel.'));
      if (!inWord) return Promise.reject(new Error('Open this pane inside Word to insert content.'));
      return window.WSDocOps[op](args || {});
    },
  };
  const persist = (partial) => (canEdit() ? Doc.call('save', partial) : Promise.resolve(true));
  const newId = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
  const clone = (o) => JSON.parse(JSON.stringify(o));

  // ---------------- UI helpers ----------------
  let toastTimer;
  function toast(msg, isErr) {
    const t = $('toast');
    t.textContent = msg;
    t.className = 'toast' + (isErr ? ' err' : '');
    clearTimeout(toastTimer);
    toastTimer = setTimeout(() => t.classList.add('hidden'), isErr ? 6000 : 2800);
  }
  function guard(btn, fn) {
    btn.addEventListener('click', async (e) => {
      if (btn.disabled) return;
      btn.disabled = true;
      try { await fn(e); } catch (err) {
        console.error(err);
        toast(friendlyError(err), true);
      } finally { btn.disabled = false; }
    });
  }
  function friendlyError(err) {
    const m = (err && (err.message || err.toString())) || 'Something went wrong';
    if (/Word is not defined|Office is not defined/.test(m)) return 'Open this pane inside Word to insert content.';
    if (/InvalidArgument|GeneralException/.test(m)) return 'Word could not insert here. Click on an empty line and try again.';
    return m;
  }
  function needWord() {
    if (canEdit()) return true;
    toast(POPOUT ? 'This window has lost its link to Word – close it and open it again from the panel.' : 'Preview only – open this add-in inside Word to insert.', true);
    return false;
  }
  const isSet = (v) => !!apiSets[v];

  // ---------------- tabs ----------------
  document.querySelectorAll('.tabs button').forEach((b) => b.addEventListener('click', () => showTab(b.dataset.tab)));
  function showTab(name) {
    document.querySelectorAll('.tabs button').forEach((x) => x.classList.toggle('active', x.dataset.tab === name));
    document.querySelectorAll('.tab').forEach((x) => x.classList.toggle('active', x.id === 'tab-' + name));
  }

  // =====================================================================
  //  Word operations
  // =====================================================================
  async function insertXml(xml, whereOverride) {
    await Doc.call('insertXml', { xml, where: whereOverride || local.get('insertAt', 'cursor'), theme });
  }

  // Renumber questions, update totals and the answer key (also saves the question data in the document).
  async function refreshDoc() {
    if (!canEdit()) return { count: 0, total: 0, lessons: 0, header: false };
    const r = await Doc.call('refresh', { questions, lessons, theme, header: docHeader });
    if (r && r.questions) questions = r.questions;
    return r;
  }

  // Re-draw everything (teacher/student copy, style changes).
  async function rerenderAll() {
    const r = await Doc.call('rerender', { questions, lessons, theme, mode, header: docHeader });
    if (r && r.questions) questions = r.questions;
    return r;
  }

  async function saveQuestion(q) { questions[q.id] = q; }
  const replaceQuestion = (q) => Doc.call('replaceQuestion', { q, theme, mode });
  const selectedQuestionId = () => Doc.call('selectedQuestionId');
  const applyStyles = () => Doc.call('applyStyles', { theme });

  // =====================================================================
  //  Wizard
  // =====================================================================
  const TYPES = {
    mc: { ico: '◉', name: 'Multiple choice', desc: 'One correct answer' },
    multi: { ico: '☑', name: 'Tick all that apply', desc: 'Several correct answers' },
    tf: { ico: '✓✗', name: 'True or false', desc: 'Pupils tick one' },
    short: { ico: '✎', name: 'Short answer', desc: 'A few writing lines', base: 'written', lines: 3 },
    long: { ico: '¶', name: 'Extended answer', desc: 'A longer written answer', base: 'written', lines: 10 },
    blanks: { ico: '▁▁', name: 'Fill in the blanks', desc: 'Missing words, optional word bank' },
    match: { ico: '⇄', name: 'Matching', desc: 'Join each item to its pair' },
  };
  const PLACEHOLDER = {
    mc: 'e.g. What is the capital of France?',
    multi: 'e.g. Which of these are prime numbers?',
    tf: 'e.g. Water boils at 100 °C at sea level.',
    short: 'e.g. Name two renewable energy sources.',
    long: 'e.g. Explain how plants make their own food.',
    blanks: 'e.g. The [Sun] is at the centre of our [solar system].',
    match: 'e.g. Match each animal to its group.',
  };
  const wiz = { type: null, options: [], pairs: [], order: null, marksTouched: false, editingId: null, draft: null };

  function buildTypeGrid() {
    $('typeGrid').innerHTML = Object.keys(TYPES).map((k) =>
      '<button class="type-card" data-type="' + k + '"><div class="ico">' + TYPES[k].ico + '</div><b>' + TYPES[k].name + '</b><span>' + TYPES[k].desc + '</span></button>').join('');
    $('typeGrid').querySelectorAll('.type-card').forEach((b) => b.addEventListener('click', () => startWizard(b.dataset.type, null, wiz.draft)));
  }

  function startWizard(type, existing, draftCtx) {
    wiz.draft = draftCtx || null;
    wiz.type = type;
    wiz.order = null;
    wiz.marksTouched = false;
    wiz.editingId = existing ? existing.id : null;
    const t = TYPES[type];
    $('wiz-step1').classList.add('hidden');
    $('wiz-step2').classList.remove('hidden');
    $('wizTitle').textContent = existing ? 'Edit question' : '2 · ' + t.name;
    $('qText').placeholder = PLACEHOLDER[type];
    $('qTextHelp').textContent = type === 'blanks' ? 'Put each missing word in [square brackets].' : 'Tip: **bold** and *italic* work.';
    const base = t.base || type;
    $('optBlock').classList.toggle('hidden', base !== 'mc' && base !== 'multi');
    $('tfBlock').classList.toggle('hidden', base !== 'tf');
    $('writtenBlock').classList.toggle('hidden', base !== 'written');
    $('blanksBlock').classList.toggle('hidden', base !== 'blanks');
    $('matchBlock').classList.toggle('hidden', base !== 'match');
    $('optHelp').textContent = base === 'multi' ? 'Tick every correct answer' : 'Select the correct answer';
    const inDraft = !!wiz.draft;
    if (inDraft) $('wizTitle').textContent = (existing ? 'Edit question' : t.name) + ' – draft worksheet';
    $('wizInsertNext').classList.toggle('hidden', !!existing || inDraft);
    $('wizInsert').textContent = inDraft ? 'Save to worksheet' : existing ? 'Update question' : 'Insert';
    $('wizInsert').className = existing || inDraft ? 'primary' : 'secondary';
    $('wizCancelEdit').classList.toggle('hidden', !existing && !inDraft);
    $('wizBack').classList.toggle('hidden', !!existing);
    $('wizDraftNote').classList.add('hidden');
    $('wizError').classList.add('hidden');
    fillWizard(existing);
    $('qText').focus();
  }

  function fillWizard(q) {
    const t = TYPES[wiz.type];
    $('qText').value = q ? q.text : '';
    $('qMarks').value = q ? q.marks : 1;
    wiz.marksTouched = !!q;
    wiz.options = q && q.options ? clone(q.options) : [{ text: '', correct: true }, { text: '', correct: false }, { text: '', correct: false }, { text: '', correct: false }];
    wiz.pairs = q && q.pairs ? clone(q.pairs) : [{ left: '', right: '' }, { left: '', right: '' }, { left: '', right: '' }];
    wiz.order = q && q.order ? q.order.slice() : null;
    $('shuffleOpts').checked = !!(q && q.order && (q.type === 'mc' || q.type === 'multi'));
    $('qLines').value = q && q.lines ? q.lines : (t.lines || 3);
    $('qArea').value = q && q.area ? q.area : '';
    $('qAnswer').value = q && q.type === 'written' ? (q.answer || '') : '';
    $('wordBank').checked = !!(q && q.wordBank);
    const tfVal = q && q.type === 'tf' ? String(q.answer) : 'true';
    document.querySelectorAll('input[name=tf]').forEach((r) => { r.checked = r.value === tfVal; });
    renderOptionRows();
    renderPairRows();
    updateWizard();
  }

  function renderOptionRows() {
    const multi = (TYPES[wiz.type].base || wiz.type) === 'multi';
    $('optList').innerHTML = wiz.options.map((o, i) =>
      '<div class="opt-row"><span class="letter">' + String.fromCharCode(65 + i) + '</span>' +
      '<input type="' + (multi ? 'checkbox' : 'radio') + '" name="optCorrect" data-i="' + i + '" ' + (o.correct ? 'checked' : '') + ' title="Correct answer" />' +
      '<input type="text" data-i="' + i + '" value="' + esc(o.text) + '" placeholder="Option ' + String.fromCharCode(65 + i) + '" />' +
      '<button class="icon-btn" data-del="' + i + '" title="Remove">✕</button></div>').join('');
    $('optList').querySelectorAll('input[type=text]').forEach((inp) => inp.addEventListener('input', () => { wiz.options[+inp.dataset.i].text = inp.value; updateWizard(); }));
    $('optList').querySelectorAll('input[name=optCorrect]').forEach((inp) => inp.addEventListener('change', () => {
      if (!multi) wiz.options.forEach((o) => { o.correct = false; });
      wiz.options[+inp.dataset.i].correct = inp.checked;
      updateWizard();
    }));
    $('optList').querySelectorAll('[data-del]').forEach((b) => b.addEventListener('click', () => {
      if (wiz.options.length <= 2) return toast('A question needs at least two options.', true);
      wiz.options.splice(+b.dataset.del, 1);
      wiz.order = null;
      renderOptionRows(); updateWizard();
    }));
    // Enter in the last option adds a new one – fast typing.
    const inputs = $('optList').querySelectorAll('input[type=text]');
    inputs.forEach((inp, i) => inp.addEventListener('keydown', (e) => {
      if (e.key !== 'Enter') return;
      e.preventDefault();
      if (i === inputs.length - 1) { $('addOpt').click(); } else inputs[i + 1].focus();
    }));
  }

  function renderPairRows() {
    $('pairList').innerHTML = wiz.pairs.map((p, i) =>
      '<div class="pair-row"><input data-i="' + i + '" data-side="left" value="' + esc(p.left) + '" placeholder="Item ' + (i + 1) + '" />' +
      '<span class="arrow">→</span><input data-i="' + i + '" data-side="right" value="' + esc(p.right) + '" placeholder="Its match" />' +
      '<button class="icon-btn" data-del="' + i + '" title="Remove">✕</button></div>').join('');
    $('pairList').querySelectorAll('input').forEach((inp) => inp.addEventListener('input', () => { wiz.pairs[+inp.dataset.i][inp.dataset.side] = inp.value; updateWizard(); }));
    $('pairList').querySelectorAll('[data-del]').forEach((b) => b.addEventListener('click', () => {
      if (wiz.pairs.length <= 2) return toast('Matching needs at least two pairs.', true);
      wiz.pairs.splice(+b.dataset.del, 1); wiz.order = null; renderPairRows(); updateWizard();
    }));
  }

  function ensureOrder(n) {
    if (!wiz.order || wiz.order.length !== n) wiz.order = WS.shuffle(n);
    return wiz.order;
  }

  // Build the question object from the form.
  function wizQuestion() {
    const t = TYPES[wiz.type];
    const type = t.base || wiz.type;
    const q = { id: wiz.editingId || newId(), type, text: $('qText').value.trim(), marks: Math.max(0, parseInt($('qMarks').value, 10) || 0) };
    if (type === 'mc' || type === 'multi') {
      q.options = wiz.options.filter((o) => o.text.trim()).map((o) => ({ text: o.text.trim(), correct: !!o.correct }));
      if ($('shuffleOpts').checked) q.order = ensureOrder(q.options.length); else wiz.order = null;
    } else if (type === 'tf') {
      q.answer = (document.querySelector('input[name=tf]:checked') || {}).value || 'true';
    } else if (type === 'written') {
      q.lines = Math.max(1, Math.min(40, parseInt($('qLines').value, 10) || 3));
      q.answer = $('qAnswer').value.trim();
      if ($('qArea').value) q.area = $('qArea').value;
    } else if (type === 'blanks') {
      q.wordBank = $('wordBank').checked;
      const n = WS.blanksAnswers(q.text).length;
      if (q.wordBank && n) q.order = ensureOrder(n);
    } else if (type === 'match') {
      q.pairs = wiz.pairs.filter((p) => p.left.trim() || p.right.trim()).map((p) => ({ left: p.left.trim(), right: p.right.trim() }));
      q.order = ensureOrder(q.pairs.length);
    }
    return q;
  }

  function validate(q) {
    if (!q.text) return 'Please type the question.';
    if (q.type === 'mc' || q.type === 'multi') {
      if (q.options.length < 2) return 'Add at least two answer options.';
      const n = q.options.filter((o) => o.correct).length;
      if (n === 0) return 'Choose which option is correct.';
      if (q.type === 'mc' && n > 1) return 'Multiple choice has only one correct answer.';
    }
    if (q.type === 'blanks' && !WS.blanksAnswers(q.text).length) return 'Put at least one answer in [square brackets].';
    if (q.type === 'match') {
      if (q.pairs.length < 2) return 'Add at least two pairs.';
      if (q.pairs.some((p) => !p.left || !p.right)) return 'Each pair needs both sides filled in.';
    }
    return null;
  }

  function updateWizard() {
    if (!wiz.type) return;
    const q = wizQuestion();
    if (!wiz.marksTouched) { q.marks = WS.autoMarks(q); $('qMarks').value = q.marks; }
    $('marksAuto').classList.toggle('hidden', wiz.marksTouched);
    if (q.type === 'blanks') {
      const a = WS.blanksAnswers(q.text);
      $('blankChips').innerHTML = a.length ? a.map((x) => '<span class="chip">' + esc(x) + '</span>').join('') : '<small>Put each missing word in [square brackets].</small>';
    }
    $('preview').innerHTML = previewHtml(q);
  }

  // ---------- HTML preview (mirrors the Word layout) ----------
  function richHtml(s) {
    return esc(s).replace(/\*\*([^*]+)\*\*/g, '<b>$1</b>').replace(/\*([^*\s][^*]*)\*/g, '<i>$1</i>').replace(/\n/g, '<br>');
  }
  const HINTS = { mc: 'Tick one box.', multi: 'Tick all that apply.', tf: 'Tick true or false.', match: 'Match each item on the left to one on the right.' };

  function previewHtml(q, num, opts) {
    const th = (opts && opts.theme) || theme;
    const show = opts && opts.show != null ? opts.show : mode === 'teacher';
    if (!q.text) return '<div class="pv-empty">Your question will appear here.</div>';
    let stem = q.type === 'blanks'
      ? richHtml(q.text).replace(/\[([^\]]+)\]/g, (_, a) => show ? '<u style="color:#c00000"><b>' + a + '</b></u>' : '_'.repeat(Math.max(10, Math.ceil(a.length * 1.6))))
      : richHtml(q.text);
    if (th.showHints && HINTS[q.type]) stem += ' <span class="pv-hint">' + HINTS[q.type] + '</span>';
    const a = '#' + th.accent;
    const numHtml = th.numberBadge ? '<span style="background:' + a + ';color:#fff;border-radius:2px;padding:0 3px">' + WS.numberLabel(num || 1, th.numberFormat).replace(/\.$/, '') + '</span>'
      : '<span style="color:' + a + '">' + WS.numberLabel(num || 1, th.numberFormat) + '</span>';
    const qStyle = th.questionStyle === 'tinted' ? ' style="background:' + WS.tint(th.accent, 0.9).replace(/^/, '#') + ';border-left:3px solid ' + a + ';padding:2px 0"' : '';
    let h = '<div class="pv-q"' + qStyle + '><span class="pv-num">' + numHtml + '</span><span>' + stem + '</span>' +
      (th.showMarks && q.marks ? '<span class="pv-marks">' + WS.marksLabel(q.marks, th.marksFormat) + '</span>' : '<span></span>') + '</div>';
    if (q.type === 'mc' || q.type === 'multi') {
      const order = q.order && q.order.length === q.options.length ? q.order : q.options.map((_, i) => i);
      const g = q.type === 'mc' ? ['○', '●'] : ['☐', '☒'];
      order.forEach((oi, j) => {
        const o = q.options[oi];
        h += '<div class="pv-opt"><span></span><span>' + (show && o.correct ? g[1] : g[0]) + '</span><span class="l">' + String.fromCharCode(65 + j) + '</span><span>' + richHtml(o.text) + '</span></div>';
      });
    } else if (q.type === 'tf') {
      const t = q.answer === 'true';
      h += '<div class="pv-opt"><span></span><span>' + (show && t ? '●' : '○') + '</span><span></span><span>True &nbsp;&nbsp;&nbsp;&nbsp; ' + (show && !t ? '●' : '○') + ' &nbsp;False</span></div>';
    } else if (q.type === 'written' && (q.area || th.answerArea) !== 'lines') {
      const area = q.area || th.answerArea;
      const hgt = Math.min(q.lines, 12) * (area === 'grid' ? 17 : 22);
      h += '<div class="pv-' + area + '" style="height:' + hgt + 'px">' + (show && q.answer ? '<span style="color:#c00000;font-weight:700;background:#fff">' + esc(q.answer) + '</span>' : '') + '</div>';
    } else if (q.type === 'written') {
      for (let i = 0; i < Math.min(q.lines, 12); i++) {
        h += '<div class="pv-line ' + th.lineStyle + '">' + (i === 0 && show && q.answer ? '<span style="color:#c00000;font-weight:700">' + esc(q.answer) + '</span>' : '') + '</div>';
      }
      if (q.lines > 12) h += '<small>…and ' + (q.lines - 12) + ' more lines</small>';
    } else if (q.type === 'blanks' && q.wordBank) {
      const a = WS.blanksAnswers(q.text);
      const order = q.order && q.order.length === a.length ? q.order : a.map((_, i) => i);
      h += '<div class="pv-bank"><b>Word bank:</b> ' + order.map((i) => esc(a[i])).join(' &nbsp;&nbsp; ') + '</div>';
    } else if (q.type === 'match') {
      const order = q.order && q.order.length === q.pairs.length ? q.order : q.pairs.map((_, i) => i);
      h += '<div class="pv-match">';
      q.pairs.forEach((p, i) => {
        const r = q.pairs[order[i]];
        h += '<span><b>' + (i + 1) + '.</b> ' + richHtml(p.left) + '</span><span>' + (show ? '<b style="color:#c00000">' + String.fromCharCode(65 + order.indexOf(i)) + '</b>' : '____') + '</span><span><b>' + String.fromCharCode(65 + i) + '.</b> ' + richHtml(r ? r.right : '') + '</span>';
      });
      h += '</div>';
    }
    return h;
  }

  async function wizardSubmit(again) {
    const q = wizQuestion();
    const err = validate(q);
    if (err) { $('wizError').textContent = err; $('wizError').classList.remove('hidden'); return; }
    $('wizError').classList.add('hidden');
    if (wiz.draft) { saveToDraft(q, wiz.draft); return; }
    if (!needWord()) return;
    await saveQuestion(q);
    if (wiz.editingId) {
      await replaceQuestion(q);
      await refreshDoc();
      toast('Question updated.');
      resetWizard();
      return;
    }
    await insertXml(WS.renderQuestion(q, 1, theme, { show: mode === 'teacher' }));
    const r = await refreshDoc();
    toast('Question ' + r.count + ' inserted.');
    if (again) { startWizard(wiz.type); } else resetWizard();
  }

  function resetWizard() {
    const wasDraft = !!wiz.draft;
    wiz.type = null; wiz.editingId = null; wiz.draft = null;
    $('wiz-step2').classList.add('hidden');
    $('wiz-step1').classList.remove('hidden');
    $('wizDraftNote').classList.add('hidden');
    if (wasDraft) { showTab('templates'); showDraftView(); }
  }

  // Wizard opened from the draft editor: choose a type for a new question in the draft.
  function wizardPickForDraft(ctx) {
    wiz.type = null; wiz.editingId = null; wiz.draft = ctx;
    $('wiz-step2').classList.add('hidden');
    $('wiz-step1').classList.remove('hidden');
    $('wizDraftNote').classList.remove('hidden');
    showTab('wizard');
  }
  $('wizDraftCancel').addEventListener('click', () => resetWizard());

  $('wizBack').addEventListener('click', () => {
    if (wiz.draft) { const d = wiz.draft; wizardPickForDraft(d); return; }
    resetWizard();
  });
  $('wizCancelEdit').addEventListener('click', resetWizard);
  $('qText').addEventListener('input', updateWizard);
  ['qLines', 'qAnswer'].forEach((id) => $(id).addEventListener('input', updateWizard));
  $('qArea').addEventListener('change', updateWizard);
  ['shuffleOpts', 'wordBank'].forEach((id) => $(id).addEventListener('change', updateWizard));
  document.querySelectorAll('input[name=tf]').forEach((r) => r.addEventListener('change', updateWizard));
  $('qMarks').addEventListener('input', () => { wiz.marksTouched = true; updateWizard(); });
  $('addOpt').addEventListener('click', () => {
    if (wiz.options.length >= 8) return toast('Eight options is the maximum.', true);
    wiz.options.push({ text: '', correct: false }); wiz.order = null; renderOptionRows(); updateWizard();
    const inputs = $('optList').querySelectorAll('input[type=text]'); inputs[inputs.length - 1].focus();
  });
  $('addPair').addEventListener('click', () => { wiz.pairs.push({ left: '', right: '' }); wiz.order = null; renderPairRows(); updateWizard(); });
  guard($('wizInsert'), () => wizardSubmit(false));
  guard($('wizInsertNext'), () => wizardSubmit(true));

  // =====================================================================
  //  Quick build
  // =====================================================================
  const EXAMPLE = [
    '# Part A – Multiple choice',
    'Q: What is the capital of France?',
    '* Paris', '- London', '- Berlin', '- Madrid',
    '',
    'Q: Which of these are **prime** numbers? (2 marks)',
    '- 4', '* 7', '* 11', '- 15',
    '',
    '# Part B – Short answers',
    'Q: Water boils at 100 °C at sea level.',
    'A: true',
    '',
    'Q: The [Sun] is at the centre of our solar system and [Jupiter] is the largest planet.',
    'wordbank: yes',
    '',
    'Q: Match each animal to its group.',
    'Dog = Mammal', 'Salmon = Fish', 'Eagle = Bird',
    '',
    'Q: Explain how plants make their own food. (4 marks)',
    'A: Photosynthesis – light energy, carbon dioxide and water make glucose and oxygen.',
  ].join('\n');
  const TYPE_NAMES = { mc: 'Multiple choice', multi: 'Tick all that apply', tf: 'True/false', written: 'Written', blanks: 'Fill in the blanks', match: 'Matching' };

  function updateQuick() {
    const src = $('quickText').value;
    local.set('quickDraft', src);
    const s = summaryHtml(src);
    $('quickSummary').innerHTML = s.html;
    $('quickInsert').disabled = !s.count;
    $('quickInsert').textContent = s.count ? 'Insert ' + s.count + ' question' + (s.count === 1 ? '' : 's') : 'Insert questions';
  }

  $('quickText').addEventListener('input', updateQuick);
  $('quickExample').addEventListener('click', () => { $('quickText').value = EXAMPLE; updateQuick(); });
  $('quickClear').addEventListener('click', () => { $('quickText').value = ''; updateQuick(); });
  async function insertQuickText(src, where) {
    if (!needWord()) return;
    const { items } = WS.parseQuick(src);
    if (!items.length) throw new Error('Nothing found to insert.');
    let hasContent = false;
    if (items.some((i) => i.kind === 'lesson')) { const r0 = await refreshDoc(); hasContent = r0.count + r0.lessons > 0; }
    let xml = '';
    let lessonNo = 0;
    items.forEach((it, i) => {
      if (it.kind === 'lesson') {
        const pageBreak = $('sowPageBreak').checked && (lessonNo > 0 || hasContent);
        const l = Object.assign({ id: newId() }, it.lesson, {
          opts: { pageBreak, objectives: $('sowShowLO').checked, vocab: $('sowShowVocab').checked,
            fields: $('sowFields').checked && pageBreak ? (readHeader().fields.length ? readHeader().fields : ['Name', 'Date']) : [] },
        });
        lessons[l.id] = l;
        lessonNo++;
        xml += WS.renderLesson(l, theme, l.opts);
        return;
      }
      const next = items[i + 1];
      // An empty line after a heading with nothing under it yet – somewhere to click and add questions.
      if (it.kind === 'section') xml += WS.renderSection(it.text, theme) + (!next || next.kind === 'section' ? WS.renderText('', theme) : '');
      else if (it.kind === 'text') xml += WS.renderText(it.text, theme);
      else if (it.kind === 'check') xml += WS.renderChecklist('', it.items, theme);
      else {
        it.q.id = newId();
        questions[it.q.id] = it.q;
        xml += WS.renderQuestion(it.q, 1, theme, { show: mode === 'teacher' });
      }
    });
    await insertXml(xml, where || (lessonNo ? 'end' : undefined));
    const r = await refreshDoc();
    toast('Inserted – ' + (r.lessons ? r.lessons + ' lessons, ' : '') + r.count + ' questions (' + r.total + ' marks).');
    return r;
  }
  guard($('quickInsert'), () => insertQuickText($('quickText').value));

  function summaryHtml(src) {
    const { items, errors } = WS.parseQuick(src);
    const qs = items.filter((i) => i.kind === 'question');
    const marks = qs.reduce((a, i) => a + (i.q.marks || 0), 0);
    let h = '';
    if (qs.length) {
      h += '<div class="ok">' + qs.length + ' question' + (qs.length === 1 ? '' : 's') + ' · ' + marks + ' marks</div><ul>';
      items.forEach((i) => {
        h += i.kind === 'lesson' ? '<li><b style="color:var(--accent)">Lesson: ' + esc(i.lesson.title) + '</b></li>'
          : i.kind === 'section' ? '<li><b>' + esc(i.text) + '</b></li>'
          : i.kind === 'text' ? '<li><i>Text: ' + esc(i.text.slice(0, 40)) + '</i></li>'
          : i.kind === 'check' ? '<li><i>Checklist (' + i.items.length + ' items)</i></li>'
          : '<li>' + esc(TYPE_NAMES[i.q.type]) + ' – ' + esc(i.q.text.replace(/\*/g, '').slice(0, 42)) + (i.q.text.length > 42 ? '…' : '') + ' <small>[' + i.q.marks + ']</small></li>';
      });
      h += '</ul>';
    }
    if (errors.length) h += '<div class="warn">' + errors.map(esc).join('<br>') + '</div>';
    return { html: h, count: qs.length };
  }

  // =====================================================================
  //  AI
  // =====================================================================
  const AI = window.WSAI;
  const aiCfg = (p) => Object.assign({ key: '', model: AI.PROVIDERS[p].defaultModel || '', models: null }, local.get('ai.' + p, {}));
  const aiProvider = () => $('aiProvider').value;
  const aiIsApi = (p) => AI.PROVIDERS[p].kind === 'api';
  function aiReadyProvider() {
    return [aiProvider(), 'claude', 'gemini'].find((p) => aiIsApi(p) && aiCfg(p).key) || null;
  }

  function aiProviderChanged() {
    const p = aiProvider();
    const P = AI.PROVIDERS[p];
    local.set('aiProvider', p);
    const api = aiIsApi(p);
    $('aiApiSettings').classList.toggle('hidden', !api);
    $('aiBridgeHelp').classList.toggle('hidden', api);
    if (api) {
      const c = aiCfg(p);
      $('aiKey').value = c.key;
      $('aiModel').value = c.model || P.defaultModel;
      $('aiKeyHint').textContent = P.keyHint;
      $('aiModelList').innerHTML = (c.models || P.models).map((m) => '<option>' + esc(m) + '</option>').join('');
      $('aiConnSummary').textContent = 'Connection: ' + (c.key ? P.name + ' key saved ✓' : 'add your ' + P.name + ' API key');
      if (!c.key) $('aiConn').open = true;
      $('aiGenerate').textContent = 'Generate questions';
      $('aiResultHelp').textContent = 'This uses the Quick build format, so you can edit it before inserting.';
      $('aiResultWrap').classList.toggle('hidden', !$('aiResult').value.trim());
      $('aiResult').placeholder = '';
    } else {
      $('aiBridgeText').textContent = P.help;
      $('aiConnSummary').textContent = 'How ' + P.name + ' works here';
      $('aiGenerate').textContent = 'Copy prompt & open ' + P.name;
      $('aiResultHelp').textContent = 'Paste ' + P.name + '’s answer here, check it, then insert.';
      $('aiResult').placeholder = 'Paste the answer from ' + P.name + ' here…';
      $('aiResultWrap').classList.remove('hidden');
    }
    $('aiTidy').classList.toggle('hidden', !aiReadyProvider());
    updateAiMode();
  }

  function aiSaveCfg(extra) {
    const p = aiProvider();
    if (!aiIsApi(p)) return;
    local.set('ai.' + p, Object.assign(aiCfg(p), { key: $('aiKey').value.trim(), model: $('aiModel').value.trim() }, extra || {}));
    aiProviderChanged();
  }

  function aiOptions() {
    return {
      topic: $('aiTopic').value.trim(),
      level: $('aiLevel').value.trim(),
      count: Math.max(1, Math.min(40, parseInt($('aiCount').value, 10) || 8)),
      types: Array.from($('aiTypes').querySelectorAll('input:checked')).map((i) => i.value),
      difficulty: $('aiDifficulty').value,
      source: $('aiSource').value.trim(),
      extra: $('aiExtra').value.trim(),
    };
  }

  function updateAiSummary() {
    const s = summaryHtml($('aiResult').value);
    $('aiSummary').innerHTML = s.html;
    $('aiInsert').disabled = !s.count;
    $('aiInsert').textContent = s.count ? 'Insert ' + s.count + ' question' + (s.count === 1 ? '' : 's') : 'Insert questions';
  }

  async function copyText(text) {
    try { await navigator.clipboard.writeText(text); return true; } catch (e) { /* fall back */ }
    const ta = document.createElement('textarea');
    ta.value = text; ta.style.position = 'fixed'; ta.style.opacity = '0';
    document.body.appendChild(ta); ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (e) { ok = false; }
    ta.remove();
    return ok;
  }
  function openUrl(url) {
    try {
      if (POPOUT && bridge) return bridge.call('openUrl', { url }).catch(() => window.open(url, '_blank'));
      if (inWord && Office.context.ui && Office.context.ui.openBrowserWindow) return Office.context.ui.openBrowserWindow(url);
    } catch (e) { /* fall back */ }
    window.open(url, '_blank');
  }

  async function withBusy(btn, label, fn) {
    const old = btn.textContent;
    btn.textContent = label;
    try { return await fn(); } finally { btn.textContent = old; }
  }

  $('aiProvider').addEventListener('change', aiProviderChanged);
  $('aiKey').addEventListener('change', () => aiSaveCfg());
  $('aiModel').addEventListener('change', () => aiSaveCfg());
  $('aiResult').addEventListener('input', () => {
    updateAiSummary();
    if (sow.reviewing != null && sow.pack[sow.reviewing]) { sow.pack[sow.reviewing].text = $('aiResult').value; }
  });
  $('aiResult').addEventListener('change', () => { if (sow.reviewing != null) renderPack(); });
  guard($('aiSaveKey'), async () => { aiSaveCfg(); toast('Saved.'); });
  guard($('aiForget'), async () => { $('aiKey').value = ''; aiSaveCfg({ key: '' }); toast('Key removed from this computer.'); });
  guard($('aiLoadModels'), async () => {
    aiSaveCfg();
    const p = aiProvider();
    const models = await withBusy($('aiLoadModels'), 'Loading…', () => AI.listModels(p, aiCfg(p)));
    aiSaveCfg({ models });
    toast(models.length + ' models available – pick one from the Model box.');
  });
  guard($('aiTest'), async () => {
    aiSaveCfg();
    const p = aiProvider();
    const reply = await withBusy($('aiTest'), 'Testing…', () => AI.complete(p, aiCfg(p), 'Reply with just the word OK.', { maxTokens: 20 }));
    toast('Connected to ' + AI.PROVIDERS[p].name + ' ✓' + (reply ? ' (“' + reply.trim().slice(0, 20) + '”)' : ''));
  });
  guard($('aiUseSelection'), async () => {
    if (!needWord()) return;
    const text = (await Doc.call('selectionText')) || '';
    if (!text.trim()) return toast('Select some text in the document first.', true);
    $('aiSource').value = text.trim();
    toast('Selected text added as source material (' + text.trim().split(/\s+/).length + ' words).');
  });

  guard($('aiGenerate'), async () => {
    const p = aiProvider();
    const o = aiOptions();
    if (aiMode === 'sow' && aiIsApi(p)) return generatePack();
    if (aiMode === 'topic' && !o.topic && !o.source) throw new Error('Type a topic or add some source material first.');
    if (!aiIsApi(p)) {
      const prompt = aiMode === 'sow' ? AI.bridgeSowPrompt(Object.assign(o, { lessonFilter: $('sowBridgeFilter').value.trim() })) : AI.bridgePrompt(o);
      const ok = await copyText(prompt);
      $('aiResultWrap').classList.remove('hidden');
      if (!ok) {
        $('aiResult').value = '';
        $('aiPromptBox').value = prompt;
        $('aiPromptWrap').classList.remove('hidden');
        toast('Copy the prompt from the box below, then paste it into ' + AI.PROVIDERS[p].name + '.', true);
      } else {
        $('aiPromptWrap').classList.add('hidden');
        toast('Prompt copied – paste it into ' + AI.PROVIDERS[p].name + ', then paste the answer back here.');
      }
      openUrl(AI.PROVIDERS[p].site);
      return;
    }
    aiSaveCfg();
    const text = await withBusy($('aiGenerate'), 'Writing questions…', () => AI.complete(p, aiCfg(p), AI.buildPrompt(o)));
    $('aiResult').value = AI.cleanOutput(text);
    $('aiResultWrap').classList.remove('hidden');
    updateAiSummary();
    $('aiResultWrap').scrollIntoView({ behavior: 'smooth' });
  });

  guard($('aiTidy'), async () => {
    const src = $('aiResult').value.trim();
    if (!src) throw new Error('Paste some questions into the box first.');
    const p = aiReadyProvider();
    if (!p) throw new Error('Tidying needs a Claude or Gemini API key.');
    const text = await withBusy($('aiTidy'), 'Tidying…', () => AI.complete(p, aiCfg(p), AI.tidyPrompt(src)));
    $('aiResult').value = AI.cleanOutput(text);
    updateAiSummary();
    toast('Converted with ' + AI.PROVIDERS[p].name + ' – check it, then insert.');
  });
  guard($('aiInsert'), () => (sow.reviewing != null && sow.pack[sow.reviewing] ? insertPack([sow.pack[sow.reviewing]]) : insertQuickText($('aiResult').value)));
  $('aiToQuick').addEventListener('click', () => { $('quickText').value = $('aiResult').value; updateQuick(); showTab('quick'); });

  // ---------------------------------------------------------------------
  //  Scheme of work → lesson pack
  // ---------------------------------------------------------------------
  let aiMode = 'topic';
  const sow = { file: null, unit: null, selected: new Set(), pack: [], stop: false, running: false, reviewing: null };

  function setAiMode(m) {
    aiMode = m;
    local.set('aiMode', m);
    $('aiMode').querySelectorAll('button').forEach((b) => b.classList.toggle('active', b.dataset.mode === m));
    sow.reviewing = null;
    updateAiMode();
  }
  $('aiMode').querySelectorAll('button').forEach((b) => b.addEventListener('click', () => setAiMode(b.dataset.mode)));

  function updateAiMode() {
    const p = aiProvider();
    const api = aiIsApi(p);
    const isSow = aiMode === 'sow';
    if (sow.reviewing != null && !(isSow && api)) { sow.reviewing = null; $('aiResult').value = ''; updateAiSummary(); }
    $('aiTopicRow').classList.toggle('hidden', isSow);
    $('aiSourceBlock').classList.toggle('hidden', isSow);
    $('sowPanel').classList.toggle('hidden', !isSow);
    $('sowDrop').classList.toggle('hidden', !isSow || !api || !!sow.file);
    $('sowFileCard').classList.toggle('hidden', !isSow || !api || !sow.file);
    $('sowAnalyseRow').classList.toggle('hidden', !api);
    $('sowAnalyse').disabled = !sow.file;
    $('sowBridgeSteps').classList.toggle('hidden', api);
    document.querySelectorAll('.bridge-name').forEach((e) => { e.textContent = AI.PROVIDERS[p].name; });
    $('sowOutline').classList.toggle('hidden', !isSow || !api || !sow.unit);
    $('sowPack').classList.toggle('hidden', !isSow || !api || !sow.pack.length);
    $('aiCount').value = $('aiCount').value || 8;
    if (isSow && api) {
      const n = sow.unit ? sow.selected.size : 0;
      $('aiGenerate').textContent = n ? 'Write worksheets for ' + n + ' lesson' + (n === 1 ? '' : 's') : 'Write the worksheets';
      $('aiGenerate').disabled = !n || sow.running;
      $('aiResultWrap').classList.toggle('hidden', sow.reviewing == null);
    } else {
      $('aiGenerate').disabled = false;
      $('aiGenerate').textContent = api ? 'Generate questions' : 'Copy prompt & open ' + AI.PROVIDERS[p].name;
      if (api) $('aiResultWrap').classList.toggle('hidden', !$('aiResult').value.trim());
      else $('aiResultWrap').classList.remove('hidden');
      $('aiResultTitle').textContent = 'Review and edit';
    }
  }

  async function takeFile(file) {
    if (!file) return;
    try {
      const f = await window.WSDocs.readFile(file);
      sow.file = f; sow.unit = null; sow.pack = []; sow.reviewing = null;
      $('sowFileName').textContent = f.name;
      $('sowFileInfo').textContent = (f.size / 1024 < 1024 ? Math.max(1, Math.round(f.size / 1024)) + ' KB' : (f.size / 1048576).toFixed(1) + ' MB') + ' · ' + f.summary;
      updateAiMode();
      toast('Scheme of work loaded – now press “Find the lessons”.');
    } catch (e) { toast(e.message || String(e), true); }
  }
  const drop = $('sowDrop');
  ['dragenter', 'dragover'].forEach((ev) => drop.addEventListener(ev, (e) => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'dragend'].forEach((ev) => drop.addEventListener(ev, () => drop.classList.remove('over')));
  drop.addEventListener('drop', (e) => { e.preventDefault(); drop.classList.remove('over'); takeFile(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]); });
  window.addEventListener('dragover', (e) => e.preventDefault());
  window.addEventListener('drop', (e) => { e.preventDefault(); if (aiMode === 'sow' && e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]) takeFile(e.dataTransfer.files[0]); });
  $('sowBrowse').addEventListener('click', () => $('sowFileInput').click());
  drop.addEventListener('keydown', (e) => { if (e.key === 'Enter' || e.key === ' ') { e.preventDefault(); $('sowFileInput').click(); } });
  $('sowFileInput').addEventListener('change', (e) => { takeFile(e.target.files[0]); e.target.value = ''; });
  $('sowFileClear').addEventListener('click', () => { sow.file = null; sow.unit = null; sow.pack = []; sow.reviewing = null; updateAiMode(); });

  guard($('sowAnalyse'), async () => {
    const p = aiProvider();
    if (!sow.file) throw new Error('Drop a scheme of work file first.');
    aiSaveCfg();
    const f = sow.file;
    const att = f.kind === 'text' ? null : { kind: f.kind, mime: f.mime, base64: f.base64 };
    const text = await withBusy($('sowAnalyse'), 'Reading the scheme of work…', () => AI.complete(p, aiCfg(p), AI.outlinePrompt(f.kind === 'text' ? f.text : null),
      { attachment: att, maxTokens: 12000, system: 'You are an expert UK curriculum planner. You extract accurate, structured information from school planning documents.' }));
    const unit = AI.parseJsonObject(text);
    const list = Array.isArray(unit.lessons) ? unit.lessons : [];
    if (!list.length) throw new Error('No lessons could be found in that file. Is it a scheme of work?');
    unit.lessons = list.map((l, i) => ({
      title: String(l.title || 'Lesson ' + (i + 1)).trim(),
      objectives: (Array.isArray(l.objectives) ? l.objectives : [l.objectives]).filter(Boolean).map(String),
      content: String(l.content || ''),
      vocabulary: (Array.isArray(l.vocabulary) ? l.vocabulary : String(l.vocabulary || '').split(/\s*,\s*/)).filter(Boolean).map(String),
    }));
    sow.unit = unit;
    sow.selected = new Set(unit.lessons.map((_, i) => i));
    sow.pack = [];
    if (unit.yearGroup && !$('aiLevel').value) $('aiLevel').value = unit.yearGroup;
    renderOutline();
    toast(unit.lessons.length + ' lessons found. Untick any you don’t need.');
  });

  function renderOutline() {
    const u = sow.unit;
    $('sowUnit').textContent = u.unit || 'Unit';
    $('sowUnitInfo').textContent = [u.subject, u.yearGroup, u.lessons.length + ' lessons'].filter(Boolean).join(' · ');
    $('sowLessons').innerHTML = u.lessons.map((l, i) => '<label class="lesson"><input type="checkbox" data-i="' + i + '"' + (sow.selected.has(i) ? ' checked' : '') + ' />' +
      '<b>' + esc(l.title) + '</b><span></span><small>' + esc(l.objectives.slice(0, 2).join(' · ') || l.content.slice(0, 120)) + '</small></label>').join('');
    $('sowLessons').querySelectorAll('input').forEach((c) => c.addEventListener('change', () => {
      if (c.checked) sow.selected.add(+c.dataset.i); else sow.selected.delete(+c.dataset.i);
      $('sowCount').textContent = sow.selected.size + ' selected';
      updateAiMode();
    }));
    $('sowCount').textContent = sow.selected.size + ' selected';
    updateAiMode();
  }
  $('sowAll').addEventListener('click', () => { sow.unit.lessons.forEach((_, i) => sow.selected.add(i)); renderOutline(); });
  $('sowNone').addEventListener('click', () => { sow.selected.clear(); renderOutline(); });

  function packStats(text) {
    const qs = WS.parseQuick(text).items.filter((i) => i.kind === 'question');
    return { n: qs.length, marks: qs.reduce((a, i) => a + (i.q.marks || 0), 0) };
  }

  function renderPack() {
    const ICON = { waiting: '⏳', working: '✍️', done: '✅', error: '⚠️' };
    $('sowPackList').innerHTML = sow.pack.map((it, i) => {
      const st = it.status === 'done' ? packStats(it.text) : null;
      const info = it.status === 'done' ? st.n + ' questions · ' + st.marks + ' marks' : it.status === 'error' ? esc(it.error) : it.status === 'working' ? 'Writing…' : 'Waiting';
      const btns = it.status === 'done' ? '<button class="secondary" data-rev="' + i + '">Review</button><button class="secondary" data-ins="' + i + '">Insert</button>'
        : it.status === 'error' ? '<button class="secondary" data-retry="' + i + '">Retry</button>' : '';
      return '<div class="lesson' + (it.status === 'error' ? ' err' : '') + (sow.reviewing === i ? ' active' : '') + '"><span class="st">' + ICON[it.status] + '</span><b>' + esc(it.lesson.title) + '</b><span class="btns">' + btns + '</span><small>' + info + '</small></div>';
    }).join('');
    $('sowPackList').querySelectorAll('[data-rev]').forEach((b) => b.addEventListener('click', () => reviewLesson(+b.dataset.rev)));
    $('sowPackList').querySelectorAll('[data-ins]').forEach((b) => guard(b, () => insertPack([sow.pack[+b.dataset.ins]])));
    $('sowPackList').querySelectorAll('[data-retry]').forEach((b) => guard(b, () => runPack([+b.dataset.retry])));
    const done = sow.pack.filter((x) => x.status === 'done').length;
    $('sowInsertAll').disabled = !done || sow.running;
    $('sowInsertAll').textContent = done ? 'Insert all ' + done + ' lesson' + (done === 1 ? '' : 's') : 'Insert all lessons';
    updateAiMode();
  }

  function reviewLesson(i) {
    sow.reviewing = i;
    const it = sow.pack[i];
    $('aiResult').value = it.text;
    $('aiResultTitle').textContent = 'Review: ' + it.lesson.title;
    $('aiResultHelp').textContent = 'Edit this lesson’s questions, then insert it on its own – or use “Insert all lessons”.';
    updateAiSummary();
    renderPack();
    $('aiResultWrap').classList.remove('hidden');
    $('aiResultWrap').scrollIntoView({ behavior: 'smooth' });
  }

  async function runPack(indices) {
    const p = aiProvider();
    const cfg = aiCfg(p);
    const o = aiOptions();
    o.fullText = sow.file && sow.file.kind === 'text' ? sow.file.text : '';
    sow.running = true; sow.stop = false;
    $('sowStop').classList.remove('hidden');
    $('sowProgress').classList.remove('hidden');
    indices.forEach((i) => { sow.pack[i].status = 'waiting'; sow.pack[i].error = ''; });
    renderPack();
    let done = 0, next = 0, fatal = null;
    const progress = () => {
      $('sowBar').style.width = Math.round((done / indices.length) * 100) + '%';
      $('sowProgressText').textContent = 'Written ' + done + ' of ' + indices.length + ' lessons…';
    };
    progress();
    const worker = async () => {
      while (!sow.stop && !fatal && next < indices.length) {
        const it = sow.pack[indices[next++]];
        it.status = 'working'; renderPack();
        try {
          const text = await AI.complete(p, cfg, AI.lessonPrompt(o, sow.unit, it.lesson));
          it.text = AI.cleanOutput(text).split(/\r?\n/).filter((l) => !/^\s*(={3,}|LO\s*:|vocab\s*:)/i.test(l)).join('\n');
          it.status = packStats(it.text).n ? 'done' : 'error';
          if (it.status === 'error') it.error = 'No questions came back – try again.';
        } catch (e) {
          it.status = 'error'; it.error = e.message || String(e);
          if (/API key|credit/i.test(it.error)) fatal = it.error;
        }
        done++; progress(); renderPack();
      }
    };
    try { await Promise.all([worker(), worker(), worker()]); } finally {
      sow.running = false;
      $('sowStop').classList.add('hidden');
      $('sowProgress').classList.add('hidden');
      sow.pack.forEach((x) => { if (x.status === 'working' || x.status === 'waiting') x.status = x.text ? 'done' : 'error', x.error = x.error || 'Stopped'; });
      renderPack();
    }
    if (fatal) throw new Error(fatal);
    const ok = sow.pack.filter((x) => x.status === 'done').length;
    toast(ok + ' of ' + sow.pack.length + ' lesson worksheets ready – review them or insert them all.');
  }

  async function generatePack() {
    if (!sow.unit) throw new Error('Press “Find the lessons” first.');
    aiSaveCfg();
    const chosen = sow.unit.lessons.filter((_, i) => sow.selected.has(i));
    if (!chosen.length) throw new Error('Tick at least one lesson.');
    sow.pack = chosen.map((lesson) => ({ lesson, text: '', status: 'waiting', error: '' }));
    sow.reviewing = null;
    $('sowPack').classList.remove('hidden');
    await runPack(sow.pack.map((_, i) => i));
    $('sowPack').scrollIntoView({ behavior: 'smooth' });
  }
  $('sowStop').addEventListener('click', () => { sow.stop = true; $('sowProgressText').textContent = 'Stopping after the current lessons…'; });

  function lessonSource(it) {
    const l = it.lesson;
    return ['=== ' + l.title].concat((l.objectives || []).map((x) => 'LO: ' + x),
      l.vocabulary && l.vocabulary.length ? ['vocab: ' + l.vocabulary.join(', ')] : [], [it.text]).join('\n');
  }

  async function insertPack(items) {
    if (!needWord()) return;
    items = items.filter((x) => x && x.status === 'done');
    if (!items.length) throw new Error('No finished lessons to insert yet.');
    const r0 = await refreshDoc();
    if ($('sowHeader').checked && !r0.header && sow.unit) {
      const base = readHeader();
      const h = Object.assign(base, {
        title: sow.unit.unit || base.title || 'Unit', subject: base.subject || sow.unit.subject || '',
        group: base.group || sow.unit.yearGroup || $('aiLevel').value.trim(), topic: '',
        showTotal: false, instructions: base.instructions,
      });
      fillHeader(h);
      await insertHeader(h);
    }
    await insertQuickText(items.map(lessonSource).join('\n'), 'end');
  }
  guard($('sowInsertAll'), () => insertPack(sow.pack));

  guard($('aiAnswers'), async () => {
    if (!needWord()) return;
    const p = aiReadyProvider();
    if (!p) throw new Error('Writing model answers needs a Claude or Gemini API key (AI tab → Connection settings).');
    const ids = await Doc.call('questionIds');
    const todo = ids.map((id) => questions[id]).filter((q) => q && q.type === 'written' && !q.answer);
    if (!todo.length) return toast('Every written question already has a model answer.');
    const h = docHeader;
    const level = $('aiLevel').value.trim() || (h && h.group) || '';
    const text = await withBusy($('aiAnswers'), 'Writing ' + todo.length + ' answers…', () => AI.complete(p, aiCfg(p), AI.answersPrompt(todo, level)));
    let n = 0;
    AI.parseJsonArray(text).forEach((a) => {
      if (a && questions[a.id] && a.answer && !questions[a.id].answer) { questions[a.id].answer = String(a.answer).trim(); n++; }
    });
    if (mode === 'teacher') await rerenderAll(); else await refreshDoc();
    toast(n + ' model answer' + (n === 1 ? '' : 's') + ' added – see the teacher copy or answer key.');
  });

  // =====================================================================
  //  Header
  // =====================================================================
  function readHeader() {
    return {
      school: $('hSchool').value.trim(), title: $('hTitle').value.trim(), subject: $('hSubject').value.trim(),
      group: $('hGroup').value.trim(), topic: $('hTopic').value.trim(),
      fields: Array.from($('hFields').querySelectorAll('input:checked')).map((i) => i.value),
      instructions: $('hInstructions').value.trim(), showTotal: $('hTotal').checked, footer: $('hFooter').checked,
    };
  }
  function fillHeader(h) {
    if (!h) return;
    $('hSchool').value = h.school || ''; $('hTitle').value = h.title || ''; $('hSubject').value = h.subject || '';
    $('hGroup').value = h.group || ''; $('hTopic').value = h.topic || ''; $('hInstructions').value = h.instructions || '';
    $('hTotal').checked = h.showTotal !== false; $('hFooter').checked = h.footer !== false;
    $('hFields').querySelectorAll('input').forEach((i) => { i.checked = (h.fields || ['Name', 'Class', 'Date']).includes(i.value); });
  }
  function remember(listKey, value) {
    if (!value) return;
    const list = local.get(listKey, []).filter((v) => v !== value);
    list.unshift(value);
    local.set(listKey, list.slice(0, 12));
  }
  function fillDatalists() {
    $('subjectList').innerHTML = local.get('subjects', []).map((s) => '<option>' + esc(s) + '</option>').join('');
    $('groupList').innerHTML = local.get('groups', []).map((s) => '<option>' + esc(s) + '</option>').join('');
  }
  $('hInstrPreset').addEventListener('change', (e) => { if (e.target.value) $('hInstructions').value = e.target.value; e.target.value = ''; });
  guard($('hSave'), async () => {
    const h = readHeader(); h.title = ''; h.topic = '';
    local.set('header', h);
    toast('Saved. New worksheets will start with these details.');
  });
  guard($('hInsert'), async () => {
    if (!needWord()) return;
    await insertHeader(readHeader());
    toast('Header ready.');
  });
  async function insertHeader(h) {
    docHeader = h;
    remember('subjects', h.subject); remember('groups', h.group); fillDatalists();
    await Doc.call('insertHeader', { h, theme });
    await refreshDoc();
  }

  // =====================================================================
  //  Tools
  // =====================================================================
  guard($('tRefresh'), async () => {
    if (!needWord()) return;
    const r = await refreshDoc();
    toast(r.count + ' questions numbered · ' + r.total + ' marks in total.');
  });
  guard($('tKey'), async () => {
    if (!needWord()) return;
    await Doc.call('ensureKey', { theme, header: docHeader });
    await refreshDoc(); // fills in the key
    toast('Answer key updated at the end of the document.');
  });
  async function setMode(m) {
    if (!needWord()) return;
    mode = m;
    await persist({ mode: m });
    updateModeBadge();
    await rerenderAll();
    toast(m === 'teacher' ? 'Teacher copy: answers are shown in red.' : 'Student copy: answers hidden.');
  }
  guard($('tStudent'), () => setMode('student'));
  guard($('tTeacher'), () => setMode('teacher'));

  guard($('tEdit'), async () => {
    if (!needWord()) return;
    const id = await selectedQuestionId();
    if (!id) return toast('Click inside a question in the document first.', true);
    const q = questions[id];
    if (!q) return toast('This question was not made with the wizard, so it can’t be edited here.', true);
    const type = q.type === 'written' ? ((q.lines || 3) > 5 ? 'long' : 'short') : q.type;
    showTab('wizard');
    startWizard(type, q);
  });
  guard($('tDelete'), async () => {
    if (!needWord()) return;
    const found = await Doc.call('deleteSelectedQuestion');
    if (!found) return toast('Click inside a question in the document first.', true);
    const r = await refreshDoc();
    toast('Deleted. ' + r.count + ' questions left.');
  });

  guard($('tBox'), async () => {
    if (!needWord()) return;
    await Doc.call('insertCheckbox', { theme });
  });
  guard($('tBoxLines'), async () => {
    if (!needWord()) return;
    await Doc.call('selectionToChecklist', { theme });
    toast('Checklist created.');
  });
  guard($('tChecklist'), async () => {
    if (!needWord()) return;
    const items = $('tChkItems').value.split(/\r?\n/).filter((s) => s.trim());
    if (!items.length) return toast('Type at least one item (one per line).', true);
    await insertXml(WS.renderChecklist($('tChkTitle').value.trim(), items, theme));
  });
  guard($('tSectionBtn'), async () => {
    if (!needWord()) return;
    const t = $('tSection').value.trim();
    if (!t) return toast('Type a heading first.', true);
    await insertXml(WS.renderSection(t, theme));
    $('tSection').value = '';
  });
  guard($('tLinesBtn'), async () => {
    if (!needWord()) return;
    await insertXml(WS.renderLines(parseInt($('tLines').value, 10) || 5, theme));
  });
  $('insertAt').addEventListener('change', (e) => local.set('insertAt', e.target.value));

  function updateModeBadge() {
    $('modeBadge').textContent = mode === 'teacher' ? 'Teacher copy' : 'Student copy';
    $('modeBadge').classList.toggle('teacher', mode === 'teacher');
    if (wiz.type) updateWizard();
  }

  // =====================================================================
  //  Style
  // =====================================================================
  const TPL = window.WSTemplates;
  const STYLE_FIELDS = ['bodyFont', 'headingFont', 'baseSize', 'accent', 'questionGap', 'lineHeight', 'lineStyle', 'numberFormat', 'marksFormat', 'checkboxMode', 'showMarks', 'showHints',
    'headerStyle', 'sectionStyle', 'questionStyle', 'answerArea', 'numberBadge'];
  const fid = (k) => 's' + k[0].toUpperCase() + k.slice(1);

  function fillStyleForm() {
    STYLE_FIELDS.forEach((k) => {
      const el = $(fid(k));
      if (el.type === 'checkbox') el.checked = !!theme[k];
      else if (el.type === 'color') el.value = '#' + theme[k];
      else {
        if (el.tagName === 'SELECT' && !Array.from(el.options).some((o) => o.value === String(theme[k]))) {
          el.add(new Option(k === 'questionGap' ? theme[k] + ' pt' : String(theme[k]), theme[k]));
        }
        el.value = theme[k];
      }
    });
    applyThemeToPane();
    renderStylePickers();
  }
  function readStyleForm() {
    STYLE_FIELDS.forEach((k) => {
      const el = $(fid(k));
      if (el.type === 'checkbox') theme[k] = el.checked;
      else if (el.type === 'color') theme[k] = el.value.replace('#', '').toUpperCase();
      else if (el.type === 'number' || k === 'questionGap') theme[k] = parseInt(el.value, 10) || WS.DEFAULT_THEME[k];
      else theme[k] = el.value || WS.DEFAULT_THEME[k];
    });
    applyThemeToPane();
    if (wiz.type) updateWizard();
  }
  function applyThemeToPane() {
    document.documentElement.style.setProperty('--pv-font', '"' + theme.bodyFont + '"');
  }
  STYLE_FIELDS.forEach((k) => $(fid(k)).addEventListener('change', readStyleForm));
  // Design gallery + colour swatches (shared by the Style and Templates tabs)
  function designPicker(el, activeId, accent, onPick) {
    el.innerHTML = TPL.DESIGNS.map((d) => '<button class="design-opt' + (d.id === activeId ? ' active' : '') + '" data-d="' + d.id + '" title="' + esc(d.desc) + '">' +
      TPL.thumb(Object.assign({}, d.theme, { accent: accent }), d.name) + d.name + '</button>').join('');
    el.querySelectorAll('[data-d]').forEach((b) => b.addEventListener('click', () => onPick(b.dataset.d)));
  }
  function schemePicker(el, activeAccent, onPick) {
    el.innerHTML = TPL.SCHEMES.map((c) => '<button class="swatch' + (c.accent === activeAccent ? ' active' : '') + '" data-c="' + c.id + '" title="' + c.name + '" style="background:#' + c.accent + '"></button>').join('');
    el.querySelectorAll('[data-c]').forEach((b) => b.addEventListener('click', () => onPick(b.dataset.c)));
  }
  function renderStylePickers() {
    designPicker($('styleDesigns'), theme.design, theme.accent, (id) => {
      Object.assign(theme, TPL.designById(id).theme, { design: id });
      fillStyleForm();
      if (wiz.type) updateWizard();
      toast(TPL.designById(id).name + ' design loaded – press “Apply to this document”.');
    });
    schemePicker($('styleSchemes'), theme.accent, (id) => {
      theme.accent = TPL.schemeById(id).accent;
      fillStyleForm();
      if (wiz.type) updateWizard();
    });
  }
  guard($('sSave'), async () => { readStyleForm(); local.set('theme', theme); toast('Saved as your default style for new worksheets.'); });
  guard($('sApply'), async () => {
    readStyleForm();
    if (!needWord()) return;
    await persist({ theme });
    const stylesUpdated = await applyStyles();
    await rerenderAll();
    toast(stylesUpdated ? 'Style applied to the whole worksheet.' : 'Questions updated. (Your Word version can’t update fonts on existing text – new content will use the new style.)');
  });

  // =====================================================================
  //  Templates
  // =====================================================================
  const tpl = { subject: 'maths', current: null, design: null, scheme: null };
  const myTemplates = () => local.get('myTemplates', []);
  const findTemplate = (id) => TPL.byId(id) || myTemplates().find((t) => t.id === id);

  function themeOfTemplate(t, designId, schemeId) {
    if (t.custom) {
      const base = designId ? Object.assign({}, TPL.designById(designId).theme, { answerArea: t.custom.answerArea }) : Object.assign({}, t.custom);
      return Object.assign(base, { accent: schemeId ? TPL.schemeById(schemeId).accent : t.custom.accent });
    }
    return TPL.themeFor(t, designId, schemeId);
  }

  function renderSubjects() {
    $('subjectChips').innerHTML = TPL.SUBJECTS.map((s) => '<button class="' + (s.id === tpl.subject ? 'active' : '') + '" data-s="' + s.id + '">' + s.ico + ' ' + s.name + '</button>').join('');
    $('subjectChips').querySelectorAll('[data-s]').forEach((b) => b.addEventListener('click', () => { tpl.subject = b.dataset.s; local.set('tplSubject', tpl.subject); renderTemplates(); }));
  }

  function cardHtml(t, deletable) {
    const th = themeOfTemplate(t);
    const d = t.custom ? 'Your style' : TPL.designById(t.design).name + ' · ' + TPL.schemeById(t.scheme).name;
    return '<button class="tpl-card" data-t="' + esc(t.id) + '">' + TPL.thumb(th, t.header.title) + '<b>' + esc(t.name) + '</b><span>' + esc(t.desc || d) + '</span>' +
      (deletable ? '<span class="icon-btn del" data-del="' + esc(t.id) + '" title="Delete">✕</span>' : '') + '</button>';
  }

  function renderTemplates() {
    renderSubjects();
    renderTopicChips();
    const list = TPL.forSubject(tpl.subject);
    $('tplList').innerHTML = list.map((t) => cardHtml(t)).join('');
    $('tplGeneral').innerHTML = TPL.general().map((t) => cardHtml(t)).join('');
    const mine = myTemplates().filter((t) => t.subjects.includes(tpl.subject));
    $('tplMine').innerHTML = mine.length ? mine.map((t) => cardHtml(t, true)).join('') : '<div class="tpl-empty">None for ' + esc(TPL.SUBJECTS.find((s) => s.id === tpl.subject).name) + ' yet.</div>';
    document.querySelectorAll('#tab-templates .tpl-card').forEach((c) => c.addEventListener('click', (e) => {
      const del = e.target.closest('[data-del]');
      if (del) { local.set('myTemplates', myTemplates().filter((t) => t.id !== del.dataset.del)); renderTemplates(); toast('Template deleted.'); return; }
      openTemplate(c.dataset.t);
    }));
  }

  function openTemplate(id) {
    const t = findTemplate(id);
    if (!t) return;
    tpl.current = t; tpl.design = null; tpl.scheme = null;
    $('tplDraft').classList.add('hidden');
    $('tplBrowse').classList.add('hidden');
    $('tplDetail').classList.remove('hidden');
    $('tplName').textContent = t.name;
    $('tplDesc').textContent = t.desc || '';
    $('tplTitle').value = t.header.title || '';
    $('tplStructureRow').classList.toggle('hidden', !t.structure);
    $('tplExamplesRow').classList.toggle('hidden', !t.examples);
    $('tplUseExamples').checked = false;
    $('tplAi').classList.toggle('hidden', !t.ai);
    const parsed = WS.parseQuick(t.structure || '').items;
    const secs = parsed.filter((i) => i.kind === 'section').map((i) => i.text);
    const nq = parsed.filter((i) => i.kind === 'question').length;
    $('tplContents').innerHTML = '<b>Includes</b><ul><li>Header: ' + esc(t.header.fields.join(', ')) + (t.header.showTotal ? ', total marks' : '') + '</li>' +
      (secs.length ? '<li>Sections: ' + esc(secs.join(', ')) + '</li>' : '') + (nq ? '<li>' + nq + ' ready-made prompts</li>' : '') +
      (t.examples ? '<li>Optional example questions</li>' : '') + '</ul>';
    renderTemplateDetail();
    window.scrollTo(0, 0);
  }

  function currentTemplateTheme() { return themeOfTemplate(tpl.current, tpl.design, tpl.scheme); }

  function renderTemplateDetail() {
    const th = currentTemplateTheme();
    $('tplPreview').innerHTML = TPL.thumb(th, $('tplTitle').value || tpl.current.header.title);
    designPicker($('tplDesigns'), tpl.design || tpl.current.design, th.accent, (id) => { tpl.design = id; renderTemplateDetail(); });
    schemePicker($('tplSchemes'), th.accent, (id) => { tpl.scheme = id; renderTemplateDetail(); });
  }

  // Put example questions under the matching headings of the structure.
  function mergeQuick(structure, examples) {
    const blocks = (src) => {
      const out = [{ head: null, body: [] }];
      String(src || '').split(/\r?\n/).forEach((l) => { if (/^#/.test(l.trim())) out.push({ head: l.trim(), body: [] }); else if (l.trim()) out[out.length - 1].body.push(l); });
      return out;
    };
    const s = blocks(structure), e = blocks(examples);
    const used = new Set();
    const lines = [].concat(e[0].body, s[0].body);
    s.slice(1).forEach((b) => {
      lines.push(b.head);
      const m = e.find((x, i) => i > 0 && x.head === b.head);
      if (m) used.add(m);
      lines.push.apply(lines, (m ? m.body : []).concat(b.body));
    });
    e.slice(1).forEach((b) => { if (!used.has(b)) lines.push.apply(lines, [b.head].concat(b.body)); });
    return lines.join('\n');
  }

  $('tplBack').addEventListener('click', () => { $('tplDetail').classList.add('hidden'); $('tplBrowse').classList.remove('hidden'); });
  $('tplTitle').addEventListener('input', () => { $('tplPreview').innerHTML = TPL.thumb(currentTemplateTheme(), $('tplTitle').value); });

  guard($('tplCreate'), async () => {
    if (!needWord()) return;
    const t = tpl.current;
    if ($('tplUseStyle').checked) {
      theme = Object.assign({}, WS.DEFAULT_THEME, currentTemplateTheme(), { checkboxMode: theme.checkboxMode, design: tpl.design || t.design });
      fillStyleForm();
      await persist({ theme });
      await applyStyles();
    }
    if ($('tplUseHeader').checked) {
      const subj = TPL.SUBJECTS.find((x) => x.id === tpl.subject);
      const h = Object.assign(readHeader(), {
        title: $('tplTitle').value.trim() || t.header.title, instructions: t.header.instructions,
        fields: t.header.fields.slice(), showTotal: t.header.showTotal,
      });
      if (!h.subject && subj && !t.subjects.includes('*')) h.subject = subj.name;
      fillHeader(h);
      await insertHeader(h);
    }
    const src = $('tplUseStructure').checked && $('tplUseExamples').checked ? mergeQuick(t.structure, t.examples)
      : $('tplUseExamples').checked ? t.examples : $('tplUseStructure').checked ? t.structure : '';
    if (src.trim()) await insertQuickText(src, 'end');
    if ($('tplUseStyle').checked) await rerenderAll(); else await refreshDoc();
    toast('“' + t.name + '” worksheet ready. Add questions with the Wizard, AI or Quick build.');
  });

  $('tplAi').addEventListener('click', () => {
    const ai = tpl.current.ai || {};
    $('aiTypes').querySelectorAll('input').forEach((i) => { i.checked = (ai.types || []).includes(i.value); });
    if (ai.difficulty) $('aiDifficulty').value = ai.difficulty;
    const h = readHeader();
    if (h.group && !$('aiLevel').value) $('aiLevel').value = h.group;
    if (!$('aiTopic').value && h.topic) $('aiTopic').value = h.topic;
    showTab('ai');
    $('aiTopic').focus();
    toast('Question types set for this template – add a topic and generate.');
  });

  guard($('tplSave'), async () => {
    readStyleForm();
    const name = $('tplSaveName').value.trim();
    if (!name) throw new Error('Give your template a name first.');
    const h = readHeader();
    const list = myTemplates();
    list.push({ id: 'my-' + newId(), name, desc: 'Saved ' + new Date().toLocaleDateString('en-GB'), subjects: [tpl.subject],
      custom: Object.assign({}, theme), header: { title: h.title || name, instructions: h.instructions, fields: h.fields, showTotal: h.showTotal }, structure: '', examples: '', ai: null });
    local.set('myTemplates', list);
    $('tplSaveName').value = '';
    renderTemplates();
    toast('Saved to My templates.');
  });

  // =====================================================================
  //  Ready-made worksheets for a topic (AI suggestions) + draft editor
  // =====================================================================
  const ideas = { topic: '', subject: '', o: {}, list: [], provider: null };
  const draft = { idx: null, template: null, design: null, scheme: null, header: {}, items: [], editing: null, busy: null, history: [], view: 'visual' };

  function renderTopicChips() {
    const list = (TPL.TOPICS && TPL.TOPICS[tpl.subject]) || [];
    $('topicChips').innerHTML = list.map((t) => '<button data-topic="' + esc(t) + '">' + esc(t) + '</button>').join('');
    $('topicChips').querySelectorAll('[data-topic]').forEach((b) => b.addEventListener('click', () => { $('ideaTopic').value = b.dataset.topic; $('ideaTopic').focus(); }));
  }

  function quickStats(text) {
    const items = WS.parseQuick(text).items;
    const qs = items.filter((i) => i.kind === 'question');
    return { n: qs.length, marks: qs.reduce((a, i) => a + (i.q.marks || 0), 0), items: items.length };
  }

  function saveIdeas() {
    local.set('ideas', { topic: ideas.topic, subject: ideas.subject, o: ideas.o, list: ideas.list });
  }

  function loadIdeas() {
    const saved = local.get('ideas', null);
    if (saved && saved.list) {
      ideas.topic = saved.topic || ''; ideas.subject = saved.subject || ''; ideas.o = saved.o || {};
      ideas.list = saved.list.map((it) => (it.status === 'ready' || it.status === 'error') ? it : Object.assign(it, { status: 'error', error: 'Interrupted – try again.' }));
      $('ideaTopic').value = ideas.topic;
    }
    $('ideaLevel').value = (saved && saved.o && saved.o.level) || local.get('ideaLevel', '');
    renderIdeas();
  }

  function ideaTemplate(it) { return TPL.byId(it.template) || TPL.general()[0]; }

  function renderIdeas() {
    $('ideaMore').classList.toggle('hidden', !ideas.list.length);
    $('ideaList').innerHTML = ideas.list.map((it, i) => {
      const t = ideaTemplate(it);
      let status, btns = '';
      if (it.status === 'ready') {
        const st = quickStats(it.text);
        status = '<small class="st ok">✓ ' + st.n + ' questions · ' + st.marks + ' marks' + (it.minutes ? ' · about ' + it.minutes + ' min' : '') + '</small>';
        btns = '<div class="btns"><button class="primary" data-open="' + i + '">Preview &amp; edit</button><button class="secondary" data-redo="' + i + '">↻ Rewrite</button></div>';
      } else if (it.status === 'error') {
        status = '<small class="st err">⚠ ' + esc(it.error) + '</small>';
        btns = '<div class="btns"><button class="secondary" data-redo="' + i + '">Try again</button></div>';
      } else {
        status = '<small class="st"><span class="spinner"></span>' + (it.status === 'working' ? 'Writing the worksheet…' : 'Waiting…') + '</small>';
      }
      return '<div class="idea' + (it.status === 'ready' ? ' ready' : '') + '" data-i="' + i + '">' + TPL.thumb(themeOfTemplate(t), it.title) +
        '<div class="idea-body"><span class="pill">' + esc(it.purpose || t.name) + '</span><b>' + esc(it.title) + '</b><small>' + esc(it.summary) + '</small>' +
        '<small class="muted">Layout: ' + esc(t.name) + '</small>' + status + btns + '</div></div>';
    }).join('');
    $('ideaList').querySelectorAll('.idea.ready').forEach((el) => el.addEventListener('click', (e) => { if (!e.target.closest('button')) openDraft(+el.dataset.i); }));
    $('ideaList').querySelectorAll('[data-open]').forEach((b) => b.addEventListener('click', () => openDraft(+b.dataset.open)));
    $('ideaList').querySelectorAll('[data-redo]').forEach((b) => guard(b, () => writeIdeas([+b.dataset.redo])));
  }

  async function writeIdeas(indices) {
    const p = aiReadyProvider();
    if (!p) throw new Error('Suggestions need a Claude or Gemini API key – add one in the AI ✦ tab.');
    const cfg = aiCfg(p);
    indices.forEach((i) => { ideas.list[i].status = 'waiting'; ideas.list[i].error = ''; });
    renderIdeas();
    let next = 0;
    const worker = async () => {
      while (next < indices.length) {
        const it = ideas.list[indices[next++]];
        it.status = 'working'; renderIdeas();
        try {
          const text = await AI.complete(p, cfg, AI.ideaPrompt(ideas.o, it, ideaTemplate(it)));
          it.text = AI.cleanOutput(text);
          if (!quickStats(it.text).n) throw new Error('No questions came back – try again.');
          it.status = 'ready';
        } catch (e) {
          it.status = 'error'; it.error = (e && e.message) || String(e);
        }
        renderIdeas(); saveIdeas();
      }
    };
    await Promise.all([worker(), worker(), worker()]);
  }

  async function generateIdeas(more) {
    const topic = $('ideaTopic').value.trim();
    if (!topic) throw new Error('Type a topic first, or tap one of the suggestions under the box.');
    const p = aiReadyProvider();
    if (!p) {
      $('ideaNote').innerHTML = 'Suggestions are written by Claude or Gemini. Add an API key in the <b>AI ✦</b> tab (Connection settings), then try again.';
      $('ideaNote').classList.remove('hidden');
      throw new Error('Add a Claude or Gemini API key in the AI ✦ tab first.');
    }
    $('ideaNote').classList.add('hidden');
    const subj = TPL.SUBJECTS.find((x) => x.id === tpl.subject);
    const o = { topic, subject: subj ? subj.name : '', level: $('ideaLevel').value.trim(), n: +$('ideaCount').value || 3,
      types: ['mc', 'tf', 'blanks', 'written'], difficulty: 'mixed' };
    local.set('ideaLevel', o.level);
    if (!more || ideas.topic !== topic || ideas.subject !== tpl.subject) ideas.list = [];
    ideas.topic = topic; ideas.subject = tpl.subject; ideas.o = o;
    const catalogue = TPL.forSubject(tpl.subject).concat(TPL.general());
    const btn = more ? $('ideaMore') : $('ideaGo');
    const text = await withBusy(btn, 'Thinking of ideas…', () => AI.complete(p, aiCfg(p),
      AI.ideasPrompt(o, catalogue.map((t) => ({ id: t.id, name: t.name, desc: t.desc })), more ? ideas.list.map((i) => i.title) : []), { maxTokens: 2500 }));
    const arr = AI.parseJsonArray(text).filter((x) => x && x.title).slice(0, o.n).map((x) => ({
      title: String(x.title).trim(), purpose: String(x.purpose || '').trim(), summary: String(x.summary || '').trim(),
      template: TPL.byId(x.template) ? x.template : catalogue[0].id,
      questions: Math.max(3, Math.min(20, parseInt(x.questions, 10) || 8)), minutes: parseInt(x.minutes, 10) || 0,
      status: 'waiting', text: '', error: '',
    }));
    if (!arr.length) throw new Error('No ideas came back – try again.');
    const start = ideas.list.length;
    ideas.list.push.apply(ideas.list, arr);
    saveIdeas();
    renderIdeas();
    $('ideaList').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
    await writeIdeas(arr.map((_, i) => start + i));
    const ok = ideas.list.filter((x) => x.status === 'ready').length;
    toast(ok + ' worksheet' + (ok === 1 ? '' : 's') + ' ready – open one to preview and edit it.');
  }
  guard($('ideaGo'), () => generateIdeas(false));
  guard($('ideaMore'), () => generateIdeas(true));

  // ---------- draft editor ----------
  function showDraftView() {
    if (draft.idx == null) return;
    $('tplBrowse').classList.add('hidden');
    $('tplDetail').classList.add('hidden');
    $('tplDraft').classList.remove('hidden');
    renderDraft();
  }

  function openDraft(i) {
    const it = ideas.list[i];
    if (!it || it.status !== 'ready') return;
    const t = ideaTemplate(it);
    draft.idx = i; draft.template = t; draft.design = it.design || null; draft.scheme = it.scheme || null;
    draft.header = { fields: t.header.fields.slice(), showTotal: t.header.showTotal };
    draft.items = WS.parseQuick(it.text).items;
    draft.editing = null; draft.busy = null; draft.history = [];
    $('draftPurpose').textContent = it.purpose || t.name;
    $('draftName').textContent = it.title;
    $('draftInfo').textContent = it.summary;
    $('draftTitle').value = it.draftTitle || it.title;
    $('draftInstr').value = it.draftInstr != null ? it.draftInstr : (t.header.instructions || '');
    $('draftAnswers').checked = false;
    setDraftView('visual');
    showDraftView();
    window.scrollTo(0, 0);
  }

  function draftTheme() { return Object.assign({}, WS.DEFAULT_THEME, themeOfTemplate(draft.template, draft.design, draft.scheme)); }

  function persistDraft() {
    const it = ideas.list[draft.idx];
    if (!it) return;
    it.text = WS.toQuick(draft.items);
    it.draftTitle = $('draftTitle').value; it.draftInstr = $('draftInstr').value;
    it.design = draft.design; it.scheme = draft.scheme;
    saveIdeas();
  }

  function pushHistory() {
    draft.history.push(JSON.stringify(draft.items));
    if (draft.history.length > 40) draft.history.shift();
  }

  function draftHeaderHtml(th) {
    const a = '#' + th.accent, tint = '#' + WS.tint(th.accent, 0.88);
    const title = esc($('draftTitle').value || 'Worksheet');
    const subj = TPL.SUBJECTS.find((x) => x.id === ideas.subject);
    const sub = esc([subj && subj.name, ideas.o.level, ideas.topic].filter(Boolean).join('  ·  ')) || '&nbsp;';
    let head;
    switch (th.headerStyle) {
      case 'banner': head = '<div style="background:' + a + ';color:#fff"><div class="d-title">' + title + '</div><div class="d-sub" style="color:#fff">' + sub + '</div></div>'; break;
      case 'boxed': head = '<div style="border:2px solid ' + a + ';text-align:center"><div class="d-title" style="color:' + a + '">' + title + '</div><div class="d-sub">' + sub + '</div></div>'; break;
      case 'stripe': head = '<div style="border-left:6px solid ' + a + ';background:' + tint + '"><div class="d-title" style="color:' + a + '">' + title + '</div><div class="d-sub">' + sub + '</div></div>'; break;
      case 'minimal': head = '<div style="border-bottom:1px solid ' + a + '"><div class="d-title" style="color:' + a + ';padding-left:0;font-size:15px">' + title + '</div><div class="d-sub" style="padding-left:0">' + sub + '</div></div>'; break;
      default: head = '<div class="d-title" style="color:' + a + ';padding-left:0">' + title + '</div><div class="d-sub" style="padding-left:0">' + sub + '</div>';
    }
    const fields = (draft.header.fields || []).map((f) => '<span>' + esc(f) + ':</span>').join('');
    const instr = $('draftInstr').value.trim()
      ? '<div class="d-instr" style="border-color:' + a + ';background:' + tint + '"><b>Instructions:</b> ' + richHtml($('draftInstr').value.trim()) + '</div>' : '';
    return '<div class="d-head" style="font-family:\'' + th.headingFont + '\',sans-serif">' + head + '<div class="d-fields">' + fields + '</div>' + instr + '</div>';
  }

  function draftSectionHtml(text, th) {
    const a = '#' + th.accent, tint = '#' + WS.tint(th.accent, 0.86);
    const st = th.sectionStyle === 'band' ? 'background:' + tint + ';border-left:3px solid ' + a + ';color:' + a
      : th.sectionStyle === 'solid' ? 'background:' + a + ';color:#fff'
        : th.sectionStyle === 'plain' ? 'color:' + a : 'color:' + a + ';border-bottom:1px solid ' + a;
    return '<div class="d-sec" style="' + st + ";font-family:'" + th.headingFont + "',sans-serif\">" + richHtml(text) + '</div>';
  }

  const tbtn = (act, i, label, title, cls) => '<button data-act="' + act + '" data-i="' + i + '" title="' + title + '"' + (cls ? ' class="' + cls + '"' : '') + '>' + label + '</button>';

  function renderDraft() {
    const th = draftTheme();
    document.documentElement.style.setProperty('--dr-font', "'" + th.bodyFont + "'");
    designPicker($('draftDesigns'), draft.design || draft.template.design, th.accent, (id) => { draft.design = id; renderDraft(); });
    schemePicker($('draftSchemes'), th.accent, (id) => { draft.scheme = id; renderDraft(); });
    const show = $('draftAnswers').checked;
    let html = draftHeaderHtml(th);
    let n = 0;
    draft.items.forEach((it, i) => {
      let body = '', tools = '';
      const mv = tbtn('up', i, '↑', 'Move up') + tbtn('down', i, '↓', 'Move down');
      const del = tbtn('del', i, '✕', 'Delete', 'del');
      if (draft.editing === i && (it.kind === 'section' || it.kind === 'text' || it.kind === 'check')) {
        const val = it.kind === 'section' ? it.text : it.kind === 'text' ? it.text : it.items.join('\n');
        const field = it.kind === 'section' ? '<input data-editfield value="' + esc(val) + '" />' : '<textarea data-editfield rows="' + (it.kind === 'text' ? 7 : 4) + '">' + esc(val) + '</textarea>';
        html += '<div class="d-item d-edit" data-i="' + i + '">' + (it.kind === 'check' ? '<small>One item per line</small>' : '') + field +
          '<div class="row"><button class="primary small" data-act="save" data-i="' + i + '">Save</button><button class="secondary small" data-act="cancel" data-i="' + i + '">Cancel</button></div></div>';
        return;
      }
      if (it.kind === 'section') {
        body = draftSectionHtml(it.text, th);
        tools = tbtn('more', i, '✨', 'Add 2 AI questions to this section') + tbtn('addq', i, '+', 'Add a question to this section') + tbtn('edit', i, '✎', 'Rename') + mv + del;
      } else if (it.kind === 'question') {
        n++;
        body = previewHtml(it.q, n, { theme: th, show });
        tools = tbtn('edit', i, '✎', 'Edit in the wizard') + tbtn('regen', i, '↻', 'Rewrite this question with AI') + mv + del;
      } else if (it.kind === 'text') {
        body = '<div class="d-text">' + richHtml(it.text) + '</div>';
        tools = tbtn('edit', i, '✎', 'Edit text') + mv + del;
      } else if (it.kind === 'check') {
        body = '<div class="d-check">' + it.items.map((x) => '<div>☐&nbsp; ' + richHtml(x) + '</div>').join('') + '</div>';
        tools = tbtn('edit', i, '✎', 'Edit items') + mv + del;
      } else if (it.kind === 'lesson') {
        body = draftSectionHtml(it.lesson.title, th);
        tools = mv + del;
      }
      html += '<div class="d-item' + (draft.busy === i ? ' d-busy' : '') + '" data-i="' + i + '">' + body + '<div class="d-tools">' + tools + '</div></div>';
    });
    if (!draft.items.length) html += '<div class="pv-empty">This worksheet is empty – add a question below.</div>';
    $('draftVisual').innerHTML = html;
    const st = quickStats(WS.toQuick(draft.items));
    $('draftSummary').innerHTML = '<div class="ok">' + st.n + ' questions · ' + st.marks + ' marks</div>';
    $('draftUndo').classList.toggle('hidden', !draft.history.length);
    persistDraft();
  }

  function endOfSection(i) {
    let j = i + 1;
    while (j < draft.items.length && draft.items[j].kind !== 'section' && draft.items[j].kind !== 'lesson') j++;
    return j;
  }
  const questionLines = (except) => draft.items.filter((x, k) => x.kind === 'question' && k !== except).map((x) => x.q.text.replace(/\s+/g, ' ').slice(0, 140));
  const writtenType = (q) => (q.type === 'written' ? ((q.lines || 3) > 5 ? 'long' : 'short') : q.type);

  function saveToDraft(q, ctx) {
    pushHistory();
    if (ctx.index != null && draft.items[ctx.index] && draft.items[ctx.index].kind === 'question') {
      draft.items[ctx.index] = { kind: 'question', q };
    } else {
      const at = ctx.after == null ? draft.items.length : endOfSection(ctx.after);
      draft.items.splice(at, 0, { kind: 'question', q });
    }
    wiz.draft = null;
    resetWizard();
    showTab('templates');
    showDraftView();
    toast('Saved to your worksheet.');
  }

  async function aiDraftQuestions(sectionIndex) {
    const p = aiReadyProvider();
    if (!p) throw new Error('This needs a Claude or Gemini API key (AI ✦ tab).');
    const section = sectionIndex == null ? 'the worksheet' : draft.items[sectionIndex].text;
    draft.busy = sectionIndex == null ? -1 : sectionIndex;
    renderDraft();
    try {
      const text = await AI.complete(p, aiCfg(p), AI.moreQuestionsPrompt(ideas.o, section, 2, questionLines(-1)));
      const qs = WS.parseQuick(AI.cleanOutput(text)).items.filter((x) => x.kind === 'question');
      if (!qs.length) throw new Error('No questions came back – try again.');
      pushHistory();
      const at = sectionIndex == null ? draft.items.length : endOfSection(sectionIndex);
      draft.items.splice.apply(draft.items, [at, 0].concat(qs));
      toast(qs.length + ' questions added.');
    } finally { draft.busy = null; renderDraft(); }
  }

  async function onDraftAction(act, i) {
    const it = draft.items[i];
    if (act === 'up' || act === 'down') {
      const j = act === 'up' ? i - 1 : i + 1;
      if (j < 0 || j >= draft.items.length) return;
      pushHistory();
      [draft.items[i], draft.items[j]] = [draft.items[j], draft.items[i]];
    } else if (act === 'del') {
      pushHistory(); draft.items.splice(i, 1);
      toast('Removed – use Undo to bring it back.');
    } else if (act === 'edit') {
      if (it.kind === 'question') { showTab('wizard'); startWizard(writtenType(it.q), it.q, { index: i }); return; }
      draft.editing = i;
    } else if (act === 'cancel') {
      draft.editing = null;
    } else if (act === 'save') {
      const el = $('draftVisual').querySelector('[data-editfield]');
      const v = el ? el.value : '';
      pushHistory();
      if (it.kind === 'section') it.text = v.trim() || it.text;
      else if (it.kind === 'text') it.text = v.replace(/\s+$/, '');
      else if (it.kind === 'check') it.items = v.split(/\r?\n/).map((x) => x.trim()).filter(Boolean);
      draft.editing = null;
    } else if (act === 'addq') {
      wizardPickForDraft({ index: null, after: i }); return;
    } else if (act === 'more') {
      await aiDraftQuestions(i); return;
    } else if (act === 'regen') {
      const p = aiReadyProvider();
      if (!p) throw new Error('This needs a Claude or Gemini API key (AI ✦ tab).');
      draft.busy = i; renderDraft();
      try {
        const text = await AI.complete(p, aiCfg(p), AI.replaceQuestionPrompt(ideas.o, WS.questionToQuick(it.q), AI.TYPE_TEXT[it.q.type] || it.q.type, questionLines(i)));
        const q = WS.parseQuick(AI.cleanOutput(text)).items.find((x) => x.kind === 'question');
        if (!q) throw new Error('No question came back – try again.');
        pushHistory();
        draft.items[i] = q;
        toast('Question rewritten.');
      } finally { draft.busy = null; }
    }
    renderDraft();
  }

  $('draftVisual').addEventListener('click', async (e) => {
    const b = e.target.closest('button[data-act]');
    if (!b || b.disabled) return;
    b.disabled = true;
    try { await onDraftAction(b.dataset.act, +b.dataset.i); } catch (err) { toast(friendlyError(err), true); } finally { b.disabled = false; }
  });

  function setDraftView(v) {
    if (v === draft.view && v === 'visual' && !$('draftTextWrap').classList.contains('hidden')) { /* fall through */ }
    if (v === 'visual' && draft.view === 'text') {
      const parsed = WS.parseQuick($('draftText').value);
      if (parsed.errors.length) toast(parsed.errors[0], true);
      pushHistory();
      draft.items = parsed.items;
    }
    draft.view = v;
    $('draftView').querySelectorAll('button').forEach((b) => b.classList.toggle('active', b.dataset.v === v));
    $('draftVisual').classList.toggle('hidden', v !== 'visual');
    $('draftHint').classList.toggle('hidden', v !== 'visual');
    $('draftTextWrap').classList.toggle('hidden', v !== 'text');
    if (v === 'text') $('draftText').value = WS.toQuick(draft.items);
    else if (draft.idx != null) renderDraft();
  }
  $('draftView').querySelectorAll('button').forEach((b) => b.addEventListener('click', () => setDraftView(b.dataset.v)));
  $('draftAnswers').addEventListener('change', renderDraft);
  $('draftTitle').addEventListener('input', renderDraft);
  $('draftInstr').addEventListener('input', renderDraft);
  $('draftBack').addEventListener('click', () => {
    if (draft.view === 'text') setDraftView('visual');
    $('tplDraft').classList.add('hidden'); $('tplBrowse').classList.remove('hidden');
    draft.idx = null; renderIdeas();
  });
  $('draftUndo').addEventListener('click', () => {
    if (!draft.history.length) return;
    draft.items = JSON.parse(draft.history.pop());
    draft.editing = null;
    renderDraft();
  });
  $('draftAdd').addEventListener('click', () => {
    if (draft.view === 'text') setDraftView('visual');
    wizardPickForDraft({ index: null, after: null });
  });
  guard($('draftAiAdd'), async () => {
    if (draft.view === 'text') setDraftView('visual');
    let last = null;
    draft.items.forEach((x, k) => { if (x.kind === 'section') last = k; });
    await aiDraftQuestions(last);
  });
  $('draftToQuick').addEventListener('click', () => {
    if (draft.view === 'text') setDraftView('visual');
    $('quickText').value = WS.toQuick(draft.items); updateQuick(); showTab('quick');
  });
  guard($('draftCreate'), async () => {
    if (draft.view === 'text') setDraftView('visual');
    if (!needWord()) return;
    if (!draft.items.length) throw new Error('This worksheet is empty.');
    const th = draftTheme();
    theme = Object.assign({}, WS.DEFAULT_THEME, th, { checkboxMode: theme.checkboxMode, design: draft.design || draft.template.design });
    fillStyleForm();
    await persist({ theme });
    await applyStyles();
    const subj = TPL.SUBJECTS.find((x) => x.id === ideas.subject);
    const base = readHeader();
    const h = Object.assign(base, {
      title: $('draftTitle').value.trim() || ideas.list[draft.idx].title,
      instructions: $('draftInstr').value.trim(), fields: (draft.header.fields || []).slice(), showTotal: !!draft.header.showTotal,
      topic: ideas.topic,
    });
    if (!h.subject && subj) h.subject = subj.name;
    if (!h.group && ideas.o.level) h.group = ideas.o.level;
    fillHeader(h);
    await insertHeader(h);
    await insertQuickText(WS.toQuick(draft.items), 'end');
    await rerenderAll();
    toast('“' + h.title + '” is now in your document.');
  });

  // =====================================================================
  //  Pop-out window
  // =====================================================================
  // Task pane: run document requests coming from the pop-out window.
  async function hostHandle(op, args) {
    if (op === 'openUrl') { openUrl(args.url); return true; }
    if (op === 'ping') return true;
    const fn = window.WSDocOps[op];
    if (typeof fn !== 'function') throw new Error('Unknown request: ' + op);
    return fn(args);
  }

  async function popOut() {
    if (POPOUT || hostWin) return;
    if (!inWord) throw new Error('Open Worksheet Builder inside Word first.');
    if (!Office.context.requirements.isSetSupported('DialogApi', '1.2')) {
      throw new Error('This version of Word can’t open add-in windows – Microsoft 365 or Word 2021 or later is needed.');
    }
    const url = location.href.split('#')[0].split('?')[0] + '?window=1';
    hostWin = await window.WSBridge.openWindow({ url, width: 60, height: 85, handle: hostHandle, onClose: onWindowClosed });
    $('popOverlay').classList.remove('hidden');
  }

  async function onWindowClosed() {
    hostWin = null;
    $('popOverlay').classList.add('hidden');
    try { await reloadFromDocument(); } catch (e) { /* keep what we have */ }
    toast('Worksheet Builder is back in the panel.');
  }

  // Read the worksheet data saved in the document (after the pop-out window has been working on it).
  async function loadState() {
    let st = null;
    try {
      if (POPOUT) st = bridge ? await bridge.call('state') : null;
      else if (inWord) st = await window.WSDocOps.state();
    } catch (e) {
      if (POPOUT) bridge = null;
    }
    st = st || {};
    questions = st.questions || {};
    lessons = st.lessons || {};
    mode = st.mode || 'student';
    docHeader = st.header || null;
    apiSets = st.sets || {};
    return st;
  }

  async function reloadFromDocument() {
    const st = await loadState();
    theme = Object.assign({}, WS.DEFAULT_THEME, local.get('theme', {}), st.theme || {});
    fillStyleForm();
    if (st.header) fillHeader(st.header);
    fillDatalists();
    updateModeBadge();
    renderTemplates();
  }

  guard($('popBtn'), async () => {
    if (POPOUT) { if (bridge) bridge.dock(); else window.close(); return; }
    await popOut();
  });
  $('popBack').addEventListener('click', () => { if (hostWin) hostWin.close(); });
  $('autoPopout').addEventListener('change', (e) => local.set('autoPopout', e.target.checked));

  // =====================================================================
  //  Start-up
  // =====================================================================
  async function init() {
    if (POPOUT) {
      document.body.classList.add('popout');
      try { bridge = window.WSBridge.connect(); await bridge.ready; } catch (e) { bridge = null; }
    }
    const st = await loadState();
    theme = Object.assign({}, WS.DEFAULT_THEME, local.get('theme', {}), st.theme || {});
    fillHeader(st.header || local.get('header', null));
    fillDatalists();
    fillStyleForm();
    updateModeBadge();
    buildTypeGrid();
    tpl.subject = local.get('tplSubject', 'maths');
    renderTemplates();
    loadIdeas();
    $('insertAt').value = local.get('insertAt', 'cursor');
    $('quickText').value = local.get('quickDraft', '');
    $('quickText').placeholder = 'Q: What is the capital of France?\n* Paris\n- London\n- Berlin';
    updateQuick();
    $('aiProvider').value = local.get('aiProvider', 'claude');
    aiMode = local.get('aiMode', 'topic');
    $('aiMode').querySelectorAll('button').forEach((b) => b.classList.toggle('active', b.dataset.mode === aiMode));
    aiProviderChanged();
    $('autoPopout').checked = !!local.get('autoPopout', false);
    $('popBtn').classList.toggle('hidden', !(POPOUT || inWord));
    if (POPOUT) { $('popBtn').textContent = '⇲ Dock'; $('popBtn').title = 'Put Worksheet Builder back in the Word panel'; }

    if (!canEdit()) {
      $('banner').textContent = POPOUT
        ? 'This window couldn’t reach Word. Close it and open it again with the ⧉ button in the Worksheet Builder panel.'
        : 'You are viewing the add-in outside Word. Everything can be previewed, but inserting only works inside Word.';
      $('banner').classList.remove('hidden');
      $('apiInfo').textContent = '';
    } else {
      const sets = ['1.3', '1.5', '1.9'].filter((v) => isSet(v));
      $('apiInfo').textContent = 'Word API support: ' + (sets.length ? sets.join(', ') : 'basic') + (isSet('1.5') ? '' : ' – style updates need Word 2021, Microsoft 365 or Word on the web.');
      if (!isSet('1.3')) {
        $('banner').textContent = 'This version of Word is too old for Worksheet Builder. Please use Microsoft 365, Word 2021 or later, or Word on the web.';
        $('banner').classList.remove('hidden');
      }
    }
    if (inWord && !POPOUT && local.get('autoPopout', false)) popOut().catch((e) => toast(e.message || String(e), true));
  }

  if (typeof Office !== 'undefined' && Office.onReady) {
    Office.onReady((info) => { inWord = !POPOUT && !!(info && info.host === Office.HostType.Word); init(); });
  } else {
    init();
  }
})();
