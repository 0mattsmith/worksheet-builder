/*
 * Worksheet Builder – template catalogue.
 * Each template sets a house style (theme), a header, a starting structure and optional example questions.
 * `structure` and `examples` use the Quick build format (see README).
 */
(function (global) {
  'use strict';

  const SUBJECTS = [
    { id: 'maths', name: 'Maths', ico: '➗' },
    { id: 'english', name: 'English', ico: '📖' },
    { id: 'science', name: 'Science', ico: '🔬' },
    { id: 'ict', name: 'ICT', ico: '💻' },
    { id: 'stem', name: 'STEM', ico: '⚙️' },
  ];

  // ------------------------------------------------------------------ BASE DESIGNS
  // Layout, fonts and shapes. Colour comes from a separate scheme so any design can be recoloured.
  const DESIGNS = [
    { id: 'classic', name: 'Classic', desc: 'Clean, familiar, works for anything',
      theme: { bodyFont: 'Arial', headingFont: 'Arial', baseSize: 12, questionGap: 14, lineHeight: 'normal', lineStyle: 'single', headerStyle: 'classic', sectionStyle: 'underline', questionStyle: 'plain', numberBadge: false, answerArea: 'lines', numberFormat: '1.', marksFormat: '[n]', showMarks: true, showHints: true } },
    { id: 'banner', name: 'Banner', desc: 'Coloured title band and shaded sections',
      theme: { bodyFont: 'Calibri', headingFont: 'Calibri', baseSize: 12, questionGap: 14, lineHeight: 'normal', lineStyle: 'single', headerStyle: 'banner', sectionStyle: 'band', questionStyle: 'plain', numberBadge: false, answerArea: 'lines', numberFormat: '1.', marksFormat: '[n]', showMarks: true, showHints: true } },
    { id: 'modern', name: 'Modern Bold', desc: 'Solid headings, number badges, tinted questions',
      theme: { bodyFont: 'Trebuchet MS', headingFont: 'Trebuchet MS', baseSize: 12, questionGap: 16, lineHeight: 'normal', lineStyle: 'single', headerStyle: 'stripe', sectionStyle: 'solid', questionStyle: 'tinted', numberBadge: true, answerArea: 'lines', numberFormat: '1.', marksFormat: '(n marks)', showMarks: true, showHints: true } },
    { id: 'exam', name: 'Exam Paper', desc: 'Plain, formal, dotted lines, ink-friendly',
      theme: { bodyFont: 'Arial', headingFont: 'Arial', baseSize: 11, questionGap: 18, lineHeight: 'normal', lineStyle: 'dotted', headerStyle: 'minimal', sectionStyle: 'plain', questionStyle: 'plain', numberBadge: false, answerArea: 'lines', numberFormat: '1.', marksFormat: '[n marks]', showMarks: true, showHints: true } },
    { id: 'elegant', name: 'Elegant', desc: 'Serif fonts and a framed, centred title',
      theme: { bodyFont: 'Georgia', headingFont: 'Georgia', baseSize: 12, questionGap: 14, lineHeight: 'normal', lineStyle: 'single', headerStyle: 'boxed', sectionStyle: 'underline', questionStyle: 'plain', numberBadge: false, answerArea: 'lines', numberFormat: '1.', marksFormat: '[n]', showMarks: true, showHints: true } },
    { id: 'playful', name: 'Playful', desc: 'Big friendly text, badges and roomy boxes',
      theme: { bodyFont: 'Comic Sans MS', headingFont: 'Comic Sans MS', baseSize: 14, questionGap: 20, lineHeight: 'wide', lineStyle: 'dashed', headerStyle: 'banner', sectionStyle: 'band', questionStyle: 'plain', numberBadge: true, answerArea: 'box', numberFormat: '1.', marksFormat: '[n]', showMarks: true, showHints: true } },
    { id: 'minimal', name: 'Minimal', desc: 'Lots of white space, light touches of colour',
      theme: { bodyFont: 'Calibri', headingFont: 'Calibri', baseSize: 12, questionGap: 18, lineHeight: 'normal', lineStyle: 'single', headerStyle: 'minimal', sectionStyle: 'plain', questionStyle: 'plain', numberBadge: false, answerArea: 'lines', numberFormat: '1.', marksFormat: '[n]', showMarks: true, showHints: true } },
    { id: 'notebook', name: 'Notebook', desc: 'Squared-paper answers and a side stripe',
      theme: { bodyFont: 'Century Gothic', headingFont: 'Century Gothic', baseSize: 11, questionGap: 14, lineHeight: 'normal', lineStyle: 'dashed', headerStyle: 'stripe', sectionStyle: 'band', questionStyle: 'plain', numberBadge: false, answerArea: 'grid', numberFormat: '1)', marksFormat: '[n]', showMarks: true, showHints: true } },
  ];

  // ------------------------------------------------------------------ COLOUR SCHEMES
  const SCHEMES = [
    { id: 'navy', name: 'Navy', accent: '1F4E79' },
    { id: 'ocean', name: 'Ocean', accent: '0277BD' },
    { id: 'teal', name: 'Teal', accent: '00796B' },
    { id: 'forest', name: 'Forest', accent: '2E7D32' },
    { id: 'indigo', name: 'Indigo', accent: '3949AB' },
    { id: 'plum', name: 'Plum', accent: '6A1B9A' },
    { id: 'berry', name: 'Berry', accent: '8E1B45' },
    { id: 'coral', name: 'Coral', accent: 'C62828' },
    { id: 'sunset', name: 'Sunset', accent: 'D84315' },
    { id: 'bronze', name: 'Bronze', accent: '8D5A00' },
    { id: 'slate', name: 'Slate', accent: '455A64' },
    { id: 'ink', name: 'Ink saver', accent: '000000' },
  ];

  // ------------------------------------------------------------------ TOPIC SUGGESTIONS
  const TOPICS = {
    maths: ['Place value', 'Fractions', 'Decimals', 'Percentages', 'Ratio and proportion', 'Negative numbers', 'Algebraic expressions',
      'Solving equations', 'Sequences', 'Straight-line graphs', 'Angles', 'Area and perimeter', 'Volume', 'Pythagoras’ theorem', 'Probability', 'Averages and range', 'Time', 'Money'],
    english: ['Persuasive writing', 'Poetry analysis', 'Descriptive writing', 'Macbeth', 'A Christmas Carol', 'Sentence types', 'Apostrophes',
      'Speech punctuation', 'Figurative language', 'Non-fiction reading', 'Prefixes and suffixes', 'Story openings', 'Formal letters', 'Myths and legends'],
    science: ['Cells', 'Photosynthesis', 'Digestion', 'Forces', 'Energy transfers', 'Electric circuits', 'States of matter', 'Acids and alkalis',
      'Atoms and elements', 'The periodic table', 'The solar system', 'Food chains and ecosystems', 'Magnetism', 'Light and sound', 'Variation and inheritance'],
    ict: ['Binary', 'Algorithms', 'Flowcharts', 'Python basics', 'Variables and data types', 'Loops', 'Online safety', 'Networks and the internet',
      'Computer hardware', 'Spreadsheets', 'Databases', 'Cyber security', 'HTML basics', 'Logic gates'],
    stem: ['Bridges and structures', 'Simple machines', 'Renewable energy', 'Paper rockets', 'Water filtration', 'Circuits and LEDs',
      'Egg drop challenge', 'Sustainable design', 'Data logging', 'Robotics basics', 'Testing materials', 'Designing a boat'],
  };

  const designById = (id) => DESIGNS.find((d) => d.id === id) || DESIGNS[0];
  const schemeById = (id) => SCHEMES.find((c) => c.id === id) || SCHEMES[0];

  // Full theme for a template, optionally with a different design / colour scheme.
  // A different design drops the template's layout tweaks except content-related ones (answer space, marks).
  function themeFor(t, designId, schemeId) {
    const d = designById(designId || t.design);
    const c = schemeById(schemeId || t.scheme);
    const keep = {};
    Object.keys(t.overrides || {}).forEach((k) => {
      if (!designId || designId === t.design || ['answerArea', 'showMarks', 'showHints'].includes(k)) keep[k] = t.overrides[k];
    });
    return Object.assign({}, d.theme, keep, { accent: c.accent });
  }

  function tintHex(hex, amount) {
    const n = parseInt(hex, 16);
    const mix = (v) => Math.round(v + (255 - v) * amount);
    return '#' + [(n >> 16) & 255, (n >> 8) & 255, n & 255].map(mix).map((v) => v.toString(16).padStart(2, '0')).join('');
  }

  // Miniature page preview for a theme (HTML).
  function thumb(th, title) {
    const a = '#' + th.accent, t = tintHex(th.accent, 0.86), t2 = tintHex(th.accent, 0.9);
    const f = "font-family:'" + th.headingFont + "',sans-serif";
    const esc = (x) => String(x || '').replace(/[&<>"]/g, (m) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[m]);
    const name = esc(title || 'Worksheet');
    let h = '';
    switch (th.headerStyle) {
      case 'banner': h = '<div class="th-h" style="background:' + a + ';color:#fff;' + f + '">' + name + '<i style="background:rgba(255,255,255,.6)"></i></div>'; break;
      case 'boxed': h = '<div class="th-h" style="border:1.5px solid ' + a + ';color:' + a + ';text-align:center;' + f + '">' + name + '<i style="margin:2px auto 0;background:#bbb"></i></div>'; break;
      case 'stripe': h = '<div class="th-h" style="background:' + t + ';border-left:5px solid ' + a + ';color:' + a + ';' + f + '">' + name + '<i style="background:#aaa"></i></div>'; break;
      case 'minimal': h = '<div class="th-h" style="color:' + a + ';font-weight:600;border-bottom:1px solid ' + a + ';padding-left:0;' + f + '">' + name + '<i style="background:#ccc"></i></div>'; break;
      default: h = '<div class="th-h" style="color:' + a + ';padding-left:0;' + f + '">' + name + '<i style="background:#bbb"></i></div>';
    }
    const fields = '<div class="th-f"><span></span><span></span><span></span></div>';
    const sec = (txt) => {
      switch (th.sectionStyle) {
        case 'band': return '<div class="th-s" style="background:' + t + ';border-left:3px solid ' + a + ';color:' + a + '">' + txt + '</div>';
        case 'solid': return '<div class="th-s" style="background:' + a + ';color:#fff">' + txt + '</div>';
        case 'plain': return '<div class="th-s" style="color:' + a + ';padding-left:0">' + txt + '</div>';
        default: return '<div class="th-s" style="color:' + a + ';border-bottom:1px solid ' + a + ';padding-left:0">' + txt + '</div>';
      }
    };
    const num = (n) => th.numberBadge ? '<b class="th-n" style="background:' + a + ';color:#fff">' + n + '</b>' : '<b class="th-n" style="color:' + a + '">' + n + '.</b>';
    const q = (n, w) => '<div class="th-q"' + (th.questionStyle === 'tinted' ? ' style="background:' + t2 + ';border-left:2px solid ' + a + '"' : '') + '>' + num(n) + '<i style="width:' + w + '%"></i></div>';
    const bubble = '<div class="th-o"><u></u><i style="width:40%"></i></div><div class="th-o"><u></u><i style="width:30%"></i></div>';
    let area;
    if (th.answerArea === 'grid') area = '<div class="th-grid"></div>';
    else if (th.answerArea === 'box') area = '<div class="th-box"></div>';
    else area = ('<div class="th-l" style="border-bottom-style:' + (th.lineStyle === 'single' ? 'solid' : th.lineStyle) + '"></div>').repeat(3);
    return '<div class="thumb" style="font-family:\'' + th.bodyFont + '\',sans-serif">' + h + fields + sec('Part A') + q(1, 70) + bubble + q(2, 55) + area + '</div>';
  }

  const T = [];
  const add = (t) => T.push(Object.assign({ subjects: [], structure: '', examples: '', ai: null }, t));

  // ------------------------------------------------------------------ MATHS
  add({
    id: 'maths-frp', subjects: ['maths'], name: 'Fluency · Reasoning · Problem solving',
    desc: 'Three-part practice sheet with squared working space.',
    design: 'banner', scheme: 'navy', overrides: { answerArea: 'grid' },
    header: { title: 'Maths Practice', instructions: 'Show your working in the grid. Check your answers when you finish.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '# Part A – Fluency\n# Part B – Reasoning\n# Part C – Problem solving',
    examples: '# Part A – Fluency\nQ: Work out 3/4 of 20.\nlines: 2\nA: 15\nQ: Round 3.476 to one decimal place.\nlines: 1\nA: 3.5\n# Part B – Reasoning\nQ: Sam says 0.45 is bigger than 0.5 because it has more digits. Is Sam right? Explain how you know. (2 marks)\nspace: lines\nA: No – 0.5 = 0.50, which is greater than 0.45.\n# Part C – Problem solving\nQ: A cinema ticket costs £7.50. How much do 6 tickets cost? (2 marks)\nA: £45',
    ai: { types: ['written', 'mc'], difficulty: 'ramp' },
  });
  add({
    id: 'maths-quickfire', subjects: ['maths'], name: 'Mental Maths Quick-fire',
    desc: 'Compact starter or plenary with a score box.',
    design: 'minimal', scheme: 'indigo', overrides: { questionGap: 8, lineHeight: 'narrow', numberFormat: '1)', showMarks: false, showHints: false },
    header: { title: 'Quick-fire Maths', instructions: 'You have 10 minutes. No calculators.', fields: ['Name', 'Date', 'Score'], showTotal: true },
    structure: '# Quick-fire questions',
    examples: '# Quick-fire questions\nQ: 7 × 8 =\nlines: 1\nA: 56\nQ: 144 ÷ 12 =\nlines: 1\nA: 12\nQ: Half of 86 =\nlines: 1\nA: 43\nQ: 25% of 60 =\nlines: 1\nA: 15',
    ai: { types: ['written'], difficulty: 'mixed' },
  });
  add({
    id: 'maths-exam', subjects: ['maths'], name: 'Exam Practice',
    desc: 'Plain, exam-paper style with working boxes and marks.',
    design: 'exam', scheme: 'ink', overrides: { answerArea: 'box' },
    header: { title: 'Exam Practice', instructions: 'Answer **all** questions. Show all your working – marks are given for method. Calculators may be used.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '',
    examples: 'Q: Solve 3x + 5 = 20. (2 marks)\nA: x = 5\nQ: Expand and simplify 2(x + 4) + 3(x − 1). (2 marks)\nA: 5x + 5',
    ai: { types: ['written', 'mc'], difficulty: 'ramp' },
  });
  add({
    id: 'maths-primary', subjects: ['maths'], name: 'Arithmetic Drill (Primary)',
    desc: 'Big friendly text, roomy boxes and “I can” statements.',
    design: 'playful', scheme: 'forest', overrides: { showMarks: false },
    header: { title: 'Let’s practise!', instructions: 'Try your best. Show your working in the boxes.', fields: ['Name', 'Date'], showTotal: false },
    structure: '# Warm up\n# Have a go\n# Challenge\n# How did I do?\n[ ] I can add and subtract in my head\n[ ] I can check my answers',
    examples: '# Warm up\nQ: 23 + 45 =\nlines: 2\nA: 68\nQ: 10 more than 57 is [67].',
    ai: { types: ['written', 'blanks', 'tf'], difficulty: 'ramp' },
  });

  // ---------------------------------------------------------------- ENGLISH
  add({
    id: 'eng-comprehension', subjects: ['english'], name: 'Reading Comprehension',
    desc: 'A reading passage followed by retrieval, inference and language questions.',
    design: 'elegant', scheme: 'berry', overrides: {},
    header: { title: 'Reading Comprehension', instructions: 'Read the text carefully, then answer the questions. Use evidence from the text.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '# Read the text\n> Paste or type the reading text here.\n# Retrieve\n# Infer\n# Language',
    examples: '# Retrieve\nQ: Where does the story take place?\n* In a lighthouse\n- On a ship\n- In a castle\n# Infer\nQ: How do you think the keeper feels when the storm arrives? Use evidence from the text. (3 marks)\nlines: 4\n# Language\nQ: Match each word from the text to its meaning.\nsheer = very steep\nglimmer = a faint light\nrelentless = never stopping',
    ai: { types: ['mc', 'written', 'match'], difficulty: 'mixed' },
  });
  add({
    id: 'eng-spag', subjects: ['english'], name: 'Grammar & Punctuation',
    desc: 'Warm-up, practise and apply – lots of tick boxes and blanks.',
    design: 'modern', scheme: 'plum', overrides: {},
    header: { title: 'Grammar & Punctuation', instructions: 'Read each question carefully.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '# Warm up\n# Practise\n# Apply it',
    examples: '# Warm up\nQ: Which sentences are punctuated correctly?\n* The dog barked loudly.\n- the dog barked loudly.\n* Where are you going?\n- Where are you going.\n# Practise\nQ: The children [were] playing outside when it started to rain.\nQ: An adverb describes a verb.\nA: true\n# Apply it\nQ: Write a sentence that uses a fronted adverbial followed by a comma. (2 marks)\nlines: 2',
    ai: { types: ['multi', 'blanks', 'tf', 'written'], difficulty: 'ramp' },
  });
  add({
    id: 'eng-writing', subjects: ['english'], name: 'Creative Writing Planner',
    desc: 'Planning boxes, a word bank and plenty of writing space.',
    design: 'notebook', scheme: 'teal', overrides: { answerArea: 'box', showMarks: false, showHints: false, lineHeight: 'wide' },
    header: { title: 'Story Planner', instructions: 'Plan your ideas first, then write your story.', fields: ['Name', 'Class', 'Date'], showTotal: false },
    structure: '# My plan\nQ: Who is your main character? What are they like? (0 marks)\nlines: 3\nQ: Where and when is your story set? (0 marks)\nlines: 3\nQ: What problem does your character face, and how is it solved? (0 marks)\nlines: 4\n# Useful words\n> Add a list of ambitious vocabulary here.\n# Write your story\nQ: Write your story here. (0 marks)\nspace: lines\nlines: 22\n# Check your work\n[ ] I have used paragraphs\n[ ] I have used interesting vocabulary\n[ ] I have checked my spelling and punctuation',
    examples: '',
    ai: null,
  });
  add({
    id: 'eng-vocab', subjects: ['english'], name: 'Spelling & Vocabulary',
    desc: 'Matching, missing words and sentence practice.',
    design: 'playful', scheme: 'sunset', overrides: { answerArea: 'lines', showMarks: false },
    header: { title: 'Spelling & Vocabulary', instructions: 'Say each word out loud before you answer.', fields: ['Name', 'Date'], showTotal: false },
    structure: '# This week’s words\n> Add this week’s spelling words here.\n# Practise',
    examples: '# Practise\nQ: Match each word to its meaning.\nenormous = very big\nancient = very old\nfragile = easily broken\nQ: The castle was [ancient] and covered in ivy.\nwordbank: yes\nQ: Write a sentence using the word “enormous”.\nlines: 2',
    ai: { types: ['match', 'blanks', 'written'], difficulty: 'mixed' },
  });

  // ---------------------------------------------------------------- SCIENCE
  add({
    id: 'sci-practical', subjects: ['science', 'stem'], name: 'Practical Write-up',
    desc: 'Aim, hypothesis, method, results grid, conclusion and evaluation.',
    design: 'classic', scheme: 'teal', overrides: { sectionStyle: 'band', answerArea: 'box', showMarks: false, showHints: false, questionGap: 10 },
    header: { title: 'Practical Write-up', instructions: 'Complete each section during and after the practical.', fields: ['Name', 'Class', 'Date'], showTotal: false },
    structure: '# Aim\nQ: What are you trying to find out? (0 marks)\nlines: 2\n# Hypothesis\nQ: What do you predict will happen, and why? (0 marks)\nlines: 3\n# Variables\nQ: Independent variable (what you change): (0 marks)\nspace: lines\nlines: 1\nQ: Dependent variable (what you measure): (0 marks)\nspace: lines\nlines: 1\nQ: Control variables (what you keep the same): (0 marks)\nspace: lines\nlines: 2\n# Method\nQ: List your equipment and write your method as numbered steps. (0 marks)\nlines: 8\n# Results\nQ: Record your results in a table, then draw a graph. (0 marks)\nspace: grid\nlines: 8\n# Conclusion\nQ: What do your results show? Was your prediction correct? (0 marks)\nlines: 4\n# Evaluation\nQ: How could you improve the investigation? (0 marks)\nlines: 3\n# Safety check\n[ ] I wore eye protection\n[ ] I cleared away safely',
    examples: '',
    ai: null,
  });
  add({
    id: 'sci-retrieval', subjects: ['science'], name: 'Retrieval Quiz',
    desc: 'Last lesson, last week, last term – quick recall questions.',
    design: 'banner', scheme: 'forest', overrides: {},
    header: { title: 'Retrieval Quiz', instructions: 'Answer from memory first, then check your notes and correct in a different colour.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '# Last lesson\n# Last week\n# Last term',
    examples: '# Last lesson\nQ: What is the chemical symbol for sodium?\n* Na\n- So\n- S\n- Sd\n# Last week\nQ: Plants release oxygen during photosynthesis.\nA: true\n# Last term\nQ: The [nucleus] controls the activities of the cell.',
    ai: { types: ['mc', 'tf', 'blanks'], difficulty: 'mixed' },
  });
  add({
    id: 'sci-keywords', subjects: ['science'], name: 'Key Words',
    desc: 'Match scientific terms, fill the gaps, use them in a sentence.',
    design: 'minimal', scheme: 'ocean', overrides: {},
    header: { title: 'Key Words', instructions: 'Learn these words – they will come up in your assessment.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '# Key words\n# Use them',
    examples: '# Key words\nQ: Match each key word to its definition.\nsolute = the substance that dissolves\nsolvent = the liquid it dissolves in\nsolution = the mixture that is formed\n# Use them\nQ: Salt is the [solute] and water is the [solvent].\nwordbank: yes',
    ai: { types: ['match', 'blanks', 'written'], difficulty: 'mixed' },
  });
  add({
    id: 'sci-exam', subjects: ['science'], name: 'Exam Practice (Science)',
    desc: 'Exam-paper layout with dotted lines and marks.',
    design: 'exam', scheme: 'ink', overrides: {},
    header: { title: 'Exam Practice', instructions: 'Answer **all** questions in the spaces provided. You may use a calculator.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '',
    examples: 'Q: Which organelle is where aerobic respiration takes place?\n* Mitochondria\n- Ribosome\n- Nucleus\n- Cell membrane\nQ: Describe how the structure of a red blood cell is related to its function. (3 marks)\nA: No nucleus – more room for haemoglobin; biconcave shape – large surface area for oxygen diffusion; small and flexible – passes through capillaries.',
    ai: { types: ['mc', 'written'], difficulty: 'ramp' },
  });

  // -------------------------------------------------------------------- ICT
  add({
    id: 'ict-theory', subjects: ['ict'], name: 'Computing Theory Quiz',
    desc: 'Modern, clean quiz for hardware, networks and data.',
    design: 'modern', scheme: 'teal', overrides: {},
    header: { title: 'Computing Theory', instructions: 'Answer all questions.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '# Recall\n# Apply',
    examples: '# Recall\nQ: What does CPU stand for?\n* Central Processing Unit\n- Computer Power Unit\n- Central Program Utility\nQ: RAM is volatile memory.\nA: true\n# Apply\nQ: The binary number 1010 is [10] in denary.\nQ: Explain one difference between RAM and ROM. (2 marks)\nA: RAM is volatile and loses its data when power is off; ROM is non-volatile and keeps its data.',
    ai: { types: ['mc', 'tf', 'blanks', 'written'], difficulty: 'mixed' },
  });
  add({
    id: 'ict-coding', subjects: ['ict', 'stem'], name: 'Coding Task Sheet',
    desc: 'Task brief, algorithm plan, testing grid, debugging and success criteria.',
    design: 'notebook', scheme: 'indigo', overrides: { answerArea: 'box', showMarks: false, showHints: false },
    header: { title: 'Coding Task', instructions: 'Plan before you code. Test your program with different inputs.', fields: ['Name', 'Class', 'Date'], showTotal: false },
    structure: '# The task\n> Describe the program pupils will create here.\n# Plan\nQ: Write your algorithm as numbered steps or pseudocode. (0 marks)\nlines: 8\n# Test\nQ: Record your tests: input, expected output, actual output. (0 marks)\nspace: grid\nlines: 6\n# Debug\nQ: Describe one error you found and how you fixed it. (0 marks)\nspace: lines\nlines: 3\n# Success criteria\n[ ] My program runs without errors\n[ ] I have used a loop or selection\n[ ] I have tested it with normal and unexpected inputs',
    examples: '',
    ai: null,
  });
  add({
    id: 'ict-skills', subjects: ['ict'], name: 'Digital Skills Checklist',
    desc: '“I can” checklist with evidence prompts – great for self-assessment.',
    design: 'classic', scheme: 'ocean', overrides: { sectionStyle: 'band', showMarks: false, showHints: false },
    header: { title: 'Digital Skills Checklist', instructions: 'Tick each skill you can do confidently, then show your teacher.', fields: ['Name', 'Class'], showTotal: false },
    structure: '# Skills\n[ ] I can create, name and organise folders\n[ ] I can save work in the correct format\n[ ] I can search the web and judge if a source is reliable\n[ ] I can keep my passwords safe\n# Evidence\nQ: Describe one thing you did this lesson that shows these skills. (0 marks)\nlines: 3\n# Next step\nQ: Which skill do you want to get better at, and how? (0 marks)\nlines: 2',
    examples: '',
    ai: null,
  });

  // ------------------------------------------------------------------- STEM
  add({
    id: 'stem-design', subjects: ['stem'], name: 'Design Challenge',
    desc: 'Engineering design cycle: research, sketch, build, test, improve.',
    design: 'modern', scheme: 'sunset', overrides: { answerArea: 'box', showMarks: false, showHints: false },
    header: { title: 'Design Challenge', instructions: 'Work through each stage of the design cycle.', fields: ['Name', 'Team', 'Date'], showTotal: false },
    structure: '# The challenge\n> Describe the problem and any constraints (materials, size, time).\n# Research\nQ: What do you already know? What do you need to find out? (0 marks)\nlines: 4\n# Design\nQ: Sketch and label your design. (0 marks)\nlines: 14\n# Build and test\nQ: What happened when you tested your design? Record any measurements. (0 marks)\nspace: grid\nlines: 5\n# Improve\nQ: What would you change next time, and why? (0 marks)\nspace: lines\nlines: 4',
    examples: '',
    ai: null,
  });
  add({
    id: 'stem-poe', subjects: ['stem', 'science'], name: 'Predict · Observe · Explain',
    desc: 'A simple investigation log for demos and experiments.',
    design: 'notebook', scheme: 'slate', overrides: { answerArea: 'box', showMarks: false, showHints: false },
    header: { title: 'Investigation Log', instructions: 'Write your prediction before you watch.', fields: ['Name', 'Class', 'Date'], showTotal: false },
    structure: '# Predict\nQ: What do you think will happen? Why? (0 marks)\nlines: 4\n# Observe\nQ: Draw or describe what you saw. (0 marks)\nlines: 8\n# Explain\nQ: Explain what happened using what you know. (0 marks)\nspace: lines\nlines: 5',
    examples: '',
    ai: null,
  });
  add({
    id: 'stem-reflect', subjects: ['stem', 'ict'], name: 'Project Reflection',
    desc: 'Self and team assessment at the end of a project.',
    design: 'minimal', scheme: 'ocean', overrides: { showMarks: false, showHints: false },
    header: { title: 'Project Reflection', instructions: 'Be honest – this helps you improve next time.', fields: ['Name', 'Team', 'Date'], showTotal: false },
    structure: '# How did it go?\nQ: What went well? (0 marks)\nlines: 3\nQ: What was difficult, and how did you deal with it? (0 marks)\nlines: 3\nQ: How well did your team work together? (0 marks)\nlines: 3\n# Rate yourself\n[ ] I shared my ideas\n[ ] I listened to others\n[ ] I kept going when things went wrong\n# Next time\nQ: One target for your next project: (0 marks)\nlines: 2',
    examples: '',
    ai: null,
  });

  // ------------------------------------------------------- ANY SUBJECT
  add({
    id: 'gen-classic', subjects: ['*'], name: 'Classic Worksheet',
    desc: 'Clean and simple – works for anything.',
    design: 'classic', scheme: 'navy', overrides: {},
    header: { title: 'Worksheet', instructions: 'Answer all questions in the spaces provided.', fields: ['Name', 'Class', 'Date'], showTotal: true },
    structure: '', examples: '', ai: { types: ['mc', 'tf', 'blanks', 'written'], difficulty: 'mixed' },
  });
  add({
    id: 'gen-homework', subjects: ['*'], name: 'Homework Sheet',
    desc: 'Due date and parent signature fields.',
    design: 'banner', scheme: 'plum', overrides: {},
    header: { title: 'Homework', instructions: 'Hand this in by the due date. Ask for help early if you are stuck.', fields: ['Name', 'Class', 'Due date'], showTotal: true },
    structure: '# Homework tasks\n# Signed\n[ ] Parent / carer has seen this homework',
    examples: '', ai: { types: ['mc', 'written', 'blanks'], difficulty: 'mixed' },
  });
  add({
    id: 'gen-exit', subjects: ['*'], name: 'Exit Ticket',
    desc: 'Three quick questions to check learning at the end of a lesson.',
    design: 'modern', scheme: 'coral', overrides: { showMarks: false },
    header: { title: 'Exit Ticket', instructions: '', fields: ['Name', 'Date'], showTotal: false },
    structure: 'Q: One thing I learned today: (0 marks)\nlines: 2\nQ: One thing I am still unsure about: (0 marks)\nlines: 2\n# How confident do I feel?\n[ ] Very confident\n[ ] Getting there\n[ ] I need more help',
    examples: '', ai: null,
  });

  function forSubject(subject) {
    return T.filter((t) => t.subjects.includes(subject));
  }
  function general() { return T.filter((t) => t.subjects.includes('*')); }
  function byId(id) { return T.find((t) => t.id === id); }

  global.WSTemplates = { SUBJECTS, TOPICS, DESIGNS, SCHEMES, ALL: T, forSubject, general, byId, themeFor, thumb, designById, schemeById };
})(typeof window !== 'undefined' ? window : globalThis);
