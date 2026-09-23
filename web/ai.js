/*
 * Worksheet Builder – AI services.
 * Claude and Gemini are called directly with the teacher's own API key.
 * NotebookLM and Chalkie have no public API, so they use a copy-prompt / paste-result bridge.
 */
(function (global) {
  'use strict';

  const PROVIDERS = {
    claude: {
      name: 'Claude', kind: 'api', site: 'https://claude.ai',
      keyUrl: 'https://console.anthropic.com/settings/keys',
      keyHint: 'Starts with sk-ant-… Get one at console.anthropic.com → API keys.',
      defaultModel: 'claude-sonnet-5',
      models: ['claude-sonnet-5', 'claude-haiku-4-5-20251001', 'claude-opus-5-5'],
    },
    gemini: {
      name: 'Gemini', kind: 'api', site: 'https://gemini.google.com',
      keyUrl: 'https://aistudio.google.com/apikey',
      keyHint: 'Get a free key at aistudio.google.com → Get API key.',
      defaultModel: 'gemini-3.8-flash',
      models: ['gemini-3.8-flash', 'gemini-3.5-flash-lite'],
    },
    notebooklm: {
      name: 'NotebookLM', kind: 'bridge', site: 'https://notebooklm.google.com',
      help: 'NotebookLM has no public API. Open your notebook (with your lesson sources added), paste the prompt into the chat, then copy its answer back here.',
    },
    chalkie: {
      name: 'Chalkie', kind: 'bridge', site: 'https://app.chalkie.ai',
      help: 'Chalkie has no public API. Paste the prompt into Chalkie, then copy the questions it writes back here. If they come back in a different layout, use “Tidy into worksheet format”.',
    },
  };

  const FORMAT_RULES = [
    'Write the worksheet ONLY in this plain-text format. No introduction, no commentary, no markdown code fences, no answer list at the end.',
    '',
    '# Section heading            (optional, a line starting with #)',
    'Q: question text (2 marks)   (every question starts with "Q:"; the marks in brackets are optional)',
    '',
    'Multiple choice (one correct):   options on their own lines; "* " before the correct option, "- " before wrong ones.',
    'Tick all that apply:             the same, with two or more "* " options.',
    'True/false:                      a statement, then a line "A: true" or "A: false".',
    'Fill in the blanks:              put each missing word in [square brackets] inside the question; add a line "wordbank: yes" for a word bank.',
    'Matching:                        pairs on their own lines as "left = right" (3–6 pairs, write them in the correct pairing; they are shuffled later).',
    'Written answer:                  a line "A: model answer" and a line "lines: N" for the number of writing lines.',
    '',
    'Example:',
    '# Part A',
    'Q: What is the capital of France?',
    '* Paris',
    '- London',
    '- Berlin',
    '- Madrid',
    'Q: Water boils at 100 °C at sea level.',
    'A: true',
    'Q: The [Sun] is at the centre of our solar system.',
    'Q: Match each animal to its group. (3 marks)',
    'Dog = Mammal',
    'Salmon = Fish',
    'Eagle = Bird',
    'Q: Explain how plants make their own food. (4 marks)',
    'A: Plants use light energy to turn carbon dioxide and water into glucose and oxygen (photosynthesis).',
    'lines: 6',
    '',
    'You may use **bold** for key words. Do not number the questions yourself.',
    'If the text contains lesson lines starting with "===", "LO:" or "vocab:", keep them exactly where they are.',
  ].join('\n');

  const LESSON_RULES = [
    'Start each lesson with a line "=== Lesson title".',
    'Straight after it, add one line "LO: …" for each learning objective and one line "vocab: word, word, word" with the key vocabulary.',
    'Then write that lesson’s questions in the format below.',
  ].join('\n');

  const SYSTEM = 'You are an experienced UK teacher who writes clear, accurate, age-appropriate classroom worksheets. ' +
    'Use British English spelling. Questions must be factually correct, unambiguous, and have exactly the answers you mark as correct. ' +
    'Multiple-choice distractors should be plausible. Never include personal data about real pupils.';

  const TYPE_TEXT = {
    mc: 'multiple choice (one correct answer)', multi: 'tick all that apply', tf: 'true/false',
    written: 'short or extended written answers with model answers', blanks: 'fill in the blanks', match: 'matching',
  };

  function buildPrompt(o) {
    const lines = [];
    lines.push('Create a worksheet with ' + o.count + ' question' + (o.count === 1 ? '' : 's') + (o.topic ? ' on: ' + o.topic : '') + '.');
    if (o.level) lines.push('Pupils: ' + o.level + '.');
    if (o.types && o.types.length) lines.push('Use a mix of these question types only: ' + o.types.map((t) => TYPE_TEXT[t]).join('; ') + '.');
    if (o.difficulty === 'ramp') lines.push('Order the questions from easy to challenging, and group them under the headings "# Part A – Warm up", "# Part B – Core" and "# Part C – Challenge".');
    else if (o.difficulty) lines.push('Difficulty: ' + o.difficulty + '.');
    if (o.sections === false) lines.push('Do not use section headings.');
    if (o.extra) lines.push('Extra instructions from the teacher: ' + o.extra);
    if (o.source && o.sourceMode === 'guide') lines.push('Base the questions on this lesson information; you may use standard subject knowledge for this age group to assess these objectives:\n"""\n' + o.source.slice(0, 30000) + '\n"""');
    else if (o.source) lines.push('Base every question only on this source material:\n"""\n' + o.source.slice(0, 30000) + '\n"""');
    lines.push('');
    lines.push(FORMAT_RULES);
    return lines.join('\n');
  }

  // ---------- scheme of work ----------
  function outlinePrompt(text) {
    return 'The ' + (text ? 'text below is' : 'attached file is') + ' a scheme of work or unit specification for a school subject. ' +
      'Identify the unit and split it into its lessons (or weeks / topics if it is not organised by lesson), keeping the document’s own order and names. ' +
      'For each lesson give its learning objectives, the key knowledge and skills in 2–4 factual sentences (enough to write quiz questions from), and key vocabulary.\n' +
      'Reply with ONLY JSON, no commentary, in exactly this shape:\n' +
      '{"unit":"unit title","subject":"subject","yearGroup":"year group or key stage if stated, else empty","lessons":[{"title":"Lesson 1: …","objectives":["…"],"content":"…","vocabulary":["…"]}]}\n' +
      'Include at most 40 lessons.' + (text ? '\n\nScheme of work:\n"""\n' + text.slice(0, 150000) + '\n"""' : '');
  }

  function lessonPrompt(o, unit, lesson) {
    const ctx = ['Unit: ' + (unit.unit || ''), 'Lesson: ' + lesson.title];
    if (lesson.objectives && lesson.objectives.length) ctx.push('Learning objectives:\n- ' + lesson.objectives.join('\n- '));
    if (lesson.content) ctx.push('Key content: ' + lesson.content);
    if (lesson.vocabulary && lesson.vocabulary.length) ctx.push('Key vocabulary: ' + lesson.vocabulary.join(', '));
    return buildPrompt(Object.assign({}, o, {
      topic: lesson.title,
      sourceMode: 'guide',
      level: o.level || unit.yearGroup || '',
      extra: ['Every learning objective must be assessed by at least one question. Use the key vocabulary where it fits.', o.extra].filter(Boolean).join(' '),
      source: ctx.join('\n') + (o.fullText ? '\n\nExtract from the scheme of work for context:\n' + o.fullText.slice(0, 12000) : ''),
    }));
  }

  function bridgeSowPrompt(o) {
    return SYSTEM + '\n\nUse the scheme of work / unit specification in the sources. Write one worksheet for each lesson' +
      (o.lessonFilter ? ' (only these lessons: ' + o.lessonFilter + ')' : '') + ', with ' + o.count + ' questions per lesson.' +
      (o.level ? ' Pupils: ' + o.level + '.' : '') +
      (o.types && o.types.length ? ' Use a mix of these question types only: ' + o.types.map((t) => TYPE_TEXT[t]).join('; ') + '.' : '') +
      (o.extra ? ' Extra instructions: ' + o.extra : '') + '\n\n' + LESSON_RULES + '\n\n' + FORMAT_RULES;
  }

  function parseJsonObject(text) {
    const t = String(text || '').replace(/```[a-z]*\n?/gi, '').replace(/```/g, '');
    const s = t.indexOf('{'), e = t.lastIndexOf('}');
    if (s < 0 || e < s) throw new Error('The AI did not return a lesson list. Try again, or try the other AI service.');
    return JSON.parse(t.slice(s, e + 1));
  }

  function bridgePrompt(o) {
    return SYSTEM + '\n\n' + buildPrompt(o);
  }

  function tidyPrompt(text) {
    return 'Convert the worksheet below into the plain-text format described. Keep the questions and answers exactly as they are; ' +
      'work out which options are correct from any answers or answer key given. Drop anything that is not a question.\n\n' +
      FORMAT_RULES + '\n\nWorksheet to convert:\n"""\n' + text.slice(0, 30000) + '\n"""';
  }

  function answersPrompt(items, level) {
    return 'Write a concise model answer (the kind a teacher uses as a mark scheme) for each question below' + (level ? ', pitched for ' + level : '') + '. ' +
      'Match the depth to the marks available. Reply with ONLY a JSON array like [{"id":"abc","answer":"…"}] and nothing else.\n\n' +
      JSON.stringify(items.map((q) => ({ id: q.id, question: q.text, marks: q.marks })));
  }

  // Remove code fences / chatter the model may add around the format.
  function cleanOutput(text) {
    let t = String(text || '').replace(/```[a-z]*\n?/gi, '').replace(/```/g, '');
    const first = t.search(/^(={3,}|#|Q\d*[:.)]|\d+[.)])/m);
    if (first > 0) t = t.slice(first);
    return t.trim();
  }

  async function readError(res) {
    let msg = res.status + ' ' + res.statusText;
    try { const j = await res.json(); msg = (j.error && (j.error.message || j.error.type)) || msg; } catch (e) { /* ignore */ }
    if (res.status === 401 || res.status === 403) msg = 'The API key was rejected – check it in the connection settings. (' + msg + ')';
    if (res.status === 429) msg = 'Too many requests or no credit left on this API account. (' + msg + ')';
    if (res.status === 404) msg = 'Model not found – press “Load models” and pick one from the list. (' + msg + ')';
    return new Error(msg);
  }

  async function callClaude(cfg, system, user, maxTokens, att) {
    const res = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'content-type': 'application/json',
        'x-api-key': cfg.key,
        'anthropic-version': '2023-06-01',
        'anthropic-dangerous-direct-browser-access': 'true',
      },
      body: JSON.stringify({ model: cfg.model || PROVIDERS.claude.defaultModel, max_tokens: maxTokens || 8000, system, messages: [{ role: 'user', content: att ? [
        att.kind === 'pdf' ? { type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: att.base64 } }
          : { type: 'image', source: { type: 'base64', media_type: att.mime, data: att.base64 } },
        { type: 'text', text: user }] : user }] }),
    });
    if (!res.ok) throw await readError(res);
    const data = await res.json();
    return (data.content || []).filter((b) => b.type === 'text').map((b) => b.text).join('');
  }

  async function callGemini(cfg, system, user, maxTokens, att) {
    const model = (cfg.model || PROVIDERS.gemini.defaultModel).replace(/^models\//, '');
    const res = await fetch('https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent', {
      method: 'POST',
      headers: { 'content-type': 'application/json', 'x-goog-api-key': cfg.key },
      body: JSON.stringify({
        systemInstruction: { parts: [{ text: system }] },
        contents: [{ role: 'user', parts: (att ? [{ inlineData: { mimeType: att.mime, data: att.base64 } }] : []).concat([{ text: user }]) }],
        generationConfig: { temperature: 0.7, maxOutputTokens: maxTokens || 8000 },
      }),
    });
    if (!res.ok) throw await readError(res);
    const data = await res.json();
    const cand = (data.candidates || [])[0];
    if (!cand) throw new Error('Gemini returned no answer' + (data.promptFeedback && data.promptFeedback.blockReason ? ' (blocked: ' + data.promptFeedback.blockReason + ')' : '') + '.');
    return ((cand.content && cand.content.parts) || []).map((p) => p.text || '').join('');
  }

  async function complete(provider, cfg, user, opts) {
    if (!cfg || !cfg.key) throw new Error('Add your ' + PROVIDERS[provider].name + ' API key in the connection settings first.');
    const system = (opts && opts.system) || SYSTEM;
    const max = opts && opts.maxTokens;
    const att = opts && opts.attachment;
    if (provider === 'claude') return callClaude(cfg, system, user, max, att);
    if (provider === 'gemini') return callGemini(cfg, system, user, max, att);
    throw new Error(PROVIDERS[provider].name + ' can only be used by copying and pasting.');
  }

  async function listModels(provider, cfg) {
    if (!cfg || !cfg.key) throw new Error('Enter the API key first.');
    if (provider === 'claude') {
      const res = await fetch('https://api.anthropic.com/v1/models?limit=100', {
        headers: { 'x-api-key': cfg.key, 'anthropic-version': '2023-06-01', 'anthropic-dangerous-direct-browser-access': 'true' },
      });
      if (!res.ok) throw await readError(res);
      return ((await res.json()).data || []).map((m) => m.id);
    }
    if (provider === 'gemini') {
      const res = await fetch('https://generativelanguage.googleapis.com/v1beta/models?pageSize=200', { headers: { 'x-goog-api-key': cfg.key } });
      if (!res.ok) throw await readError(res);
      return ((await res.json()).models || [])
        .filter((m) => (m.supportedGenerationMethods || []).includes('generateContent') && /gemini/.test(m.name))
        .map((m) => m.name.replace(/^models\//, ''));
    }
    return [];
  }

  function parseJsonArray(text) {
    const t = String(text || '').replace(/```[a-z]*\n?/gi, '').replace(/```/g, '');
    const s = t.indexOf('['), e = t.lastIndexOf(']');
    if (s < 0 || e < s) throw new Error('The AI did not return a list of answers. Try again.');
    return JSON.parse(t.slice(s, e + 1));
  }

  global.WSAI = { PROVIDERS, SYSTEM, buildPrompt, outlinePrompt, lessonPrompt, bridgeSowPrompt, parseJsonObject, bridgePrompt, tidyPrompt, answersPrompt, cleanOutput, complete, listModels, parseJsonArray };
})(typeof window !== 'undefined' ? window : globalThis);
