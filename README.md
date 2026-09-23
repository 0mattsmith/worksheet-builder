# Worksheet Builder – a Word add-in for class worksheets

Worksheet Builder adds a **Worksheet Builder** button to Word's **Home** tab. It opens a side panel that builds worksheets in one consistent house style.

| Tab | What it does |
|---|---|
| **Templates** | A catalogue of 21 ready-made worksheets for **Maths, English, Science, ICT and STEM**, plus some that work for any subject. There are practice sheets, exam practice, reading comprehension, a practical write-up, coding task sheets, design challenges, exit tickets and more. Each template sets the look, the header and a starting structure. You can switch it to any of the **8 base designs** and **12 colour schemes** before you create the worksheet, and you can save your own templates. |
| **Wizard** | Step-by-step question builder: pick a type, fill in the form, see a live preview, then click **Insert**. It supports multiple choice, tick all that apply, true/false, short answer, extended answer, fill in the blanks (with an optional word bank) and matching. Marks are worked out for you. |
| **AI ✦** | Writes questions from a topic and optional source material. **From a scheme of work:** drop in your unit plan and it writes a worksheet for every lesson. It can also write missing model answers. It works with **Claude** and **Gemini** directly, and with **NotebookLM** and **Chalkie** by copy and paste. |
| **Quick** (quick build) | Type or paste a whole worksheet in plain text and insert every question in one click. |
| **Header** | Adds a standard title block: school, title, subject, class, Name/Class/Date lines, an instructions box, an automatic total-marks box, and a footer with page numbers. |
| **Tools** | Renumbers questions and updates totals, adds an answer key, switches between the **student copy** and the **teacher copy** (answers in red), edits or deletes the selected question, adds tick boxes and checklists, and inserts section headings or writing lines. |
| **Style** | Sets the house style. Pick a base design and a colour scheme, then fine-tune the fonts, size, spacing, header, section headings, question look, answer space (lines, working box or squared grid), numbering, marks and tick boxes. **Save as my default** keeps your style for every new worksheet. |

Every inserted item uses Word styles whose names start with `WS …`, such as `WS Question` and `WS Answer Line`. That keeps worksheets uniform, and **Apply to this document** restyles a whole worksheet in one go.

---

## Install it: publish with push.ps1 (Windows, with GitHub CLI)

If you have `git` and the GitHub CLI (`gh auth login` done), one command publishes everything:

```powershell
cd path\to\worksheet-builder
Unblock-File .\push.ps1        # first time only – files from a zip are blocked by Windows
./push.ps1 "first version"
```

On the first run it:
- creates a public repository called `worksheet-builder` in your GitHub account;
- turns on GitHub Pages;
- pushes the project;
- waits until the add-in website is live;
- writes `manifest-hosted.xml` with your site address, ready to add to Word (see step 4 below).

After that, run `./push.ps1 "what you changed"` whenever you edit anything. GitHub re-publishes the add-in automatically, and Word loads the new version the next time you open it.

Options:

| Option | What it does |
|---|---|
| `-RepoName name` | Use a different repository name |
| `-Open` | Open the add-in page in your browser when publishing finishes |
| `-NoWait` | Don't wait for publishing to finish |
| `-Force` | Replace what's already in the repository (for example, files you uploaded through the GitHub website) |
| `-Private` | Make the repository private (GitHub Pages then needs a paid plan) |

If PowerShell says running scripts is disabled, use `powershell -ExecutionPolicy Bypass -File .\push.ps1 "first version"` instead.

Colleagues can download the manifest from `https://YOURNAME.github.io/worksheet-builder/manifest.xml`.

## Install it: Option 1, try it on your own computer

You need **Node.js**. Install the free "LTS" version from https://nodejs.org.

1. Unzip this folder somewhere, for example in `Documents\worksheet-builder`.
2. Open a terminal in that folder. On Windows, right-click the folder and choose **Open in Terminal**. On a Mac, open Terminal and type `cd`, then drag the folder into the window.
3. Run these commands once:
   ```
   npm install
   npm start
   ```
   The first time, Windows or macOS asks you to trust a "localhost" development certificate. Click **Yes**. Leave the window open while you use the add-in.
4. Load the add-in into Word:
   - **Word for Windows or Mac:** open a second terminal in the same folder and run `npm run sideload`. Word opens with the add-in loaded.
   - **Word on the web (Microsoft 365):** open a document, then go to **Home → Add-ins → More Settings → Upload My Add-in**. Browse to `manifest.xml` and click **Upload**. This is remembered only in that browser.
   - **Word for Windows, without the command line:** put the manifest in a folder and share it: right-click the folder, then **Properties → Sharing → Share**, and note the `\\COMPUTER\folder` path. In Word, go to **File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs**. Paste the path, click **Add catalog**, tick **Show in Menu**, click **OK**, and restart Word. Then go to **Home → Add-ins → Advanced → SHARED FOLDER**, pick Worksheet Builder and click **Add**.
   - **Mac, by hand:** copy `manifest.xml` to `~/Library/Containers/com.microsoft.Word/Data/Documents/wef/` (create the `wef` folder if it doesn't exist), then restart Word.
5. Click **Home → Worksheet Builder**.

## Install it: Option 2, host it so it works everywhere (and for colleagues)

1. Put the contents of the `web` folder on any HTTPS web host. **GitHub Pages** is free: create a repository, upload the `web` folder's files, then turn on Pages under *Settings → Pages*.
2. Run `npm run set-url -- https://YOURNAME.github.io/REPO/`. This creates `manifest-hosted.xml`, which points at your hosted copy.
3. Upload `manifest-hosted.xml` as described in step 4 of Option 1. You no longer need `npm start`.
4. To give it to the whole department, send `manifest-hosted.xml` to your school's Microsoft 365 admin. They can deploy it from **Microsoft 365 admin centre → Settings → Integrated apps → Upload custom apps**, and it will then appear in everyone's Word.

---

## A typical workflow

1. Open a new, blank document. In **Templates**, pick your subject and a template, choose a design and colours, and click **Create worksheet**. You can also go straight to **Header → Insert / update header**.
2. Add questions with the **Wizard**. **Insert & add another** keeps the same question type ready for the next one. For a whole worksheet at once, use **Quick build**.
3. Go to **Tools → Insert / update answer key**.
4. Print the **Student copy**. Then switch to the **Teacher copy** and print a mark scheme with the answers shown in red.

## Templates, designs and colours

A **template** is the starting point for a type of worksheet. The **base design** is the layout, and the **colour scheme** is applied on top of it, so every combination works:

| Base design | Look |
|---|---|
| Classic | Arial, coloured title, underlined sections, writing lines |
| Banner | Solid colour title band and shaded section bands |
| Modern Bold | Side-stripe title, solid colour section bars, number badges, tinted question strips |
| Exam Paper | Plain and formal, dotted lines, "[2 marks]", ideal with the Ink saver colours |
| Elegant | Georgia serif, framed centred title |
| Playful | Comic Sans 14, big friendly badges, roomy working boxes (for primary) |
| Minimal | Lots of white space, light touches of colour |
| Notebook | Century Gothic, side stripe, squared-paper answer grids (great for Maths) |

`examples/design-gallery.png` shows the same worksheet in all eight designs.

Colour schemes: Navy, Ocean, Teal, Forest, Indigo, Plum, Berry, Coral, Sunset, Bronze, Slate and Ink saver (black).

**Save current look** (at the bottom of Templates) stores your current style and header as your own template for the selected subject.

## Worksheets from a scheme of work

1. Go to **AI ✦**, then **From a scheme of work**.
2. Drag your scheme of work or unit specification onto the drop zone, or click **Choose a file…**. It accepts Word (.docx), PDF, PowerPoint (.pptx), Excel (.xlsx), text or a photo of a printed plan.
3. Click **Find the lessons**. The AI lists each lesson with its learning objectives and key vocabulary. Untick any lessons you don't need.
4. Set the pupils, the number of questions per lesson, the question types and the difficulty. Then click **Write worksheets for N lessons**. It writes three lessons at a time and shows progress. If one fails, you can **Retry** it.
5. **Review** any lesson to edit it, **Insert** a single lesson, or click **Insert all lessons**. Each lesson gets its own page, with its title, learning objectives, a key vocabulary box and Name/Date lines. Question numbering restarts for each lesson, and the answer key is grouped by lesson.

The details:
- **How files are read:** Word, PowerPoint, Excel and text files are read inside Word. PDFs and photos are sent to Claude or Gemini as they are, and both can read them. The file limit is 20 MB.
- **NotebookLM and Chalkie:** add the scheme of work there as a source, then click **Copy prompt & open …**. Paste the reply back and click **Insert** to get the same lesson pack.
- **Dragging files:** if dragging onto the panel doesn't work in your version of Word, use **Choose a file…** instead.
- **Old formats:** .doc, .ppt and .xls files need to be saved as .docx, .pptx or .xlsx first.

## Using AI

Open the **AI ✦** tab and choose a service:

| Service | How it connects | What you need |
|---|---|---|
| **Claude** | Directly, from inside Word | An API key from [console.anthropic.com](https://console.anthropic.com/settings/keys). This is paid per use and billed separately from a Claude.ai subscription. |
| **Gemini** | Directly, from inside Word | An API key from [aistudio.google.com](https://aistudio.google.com/apikey). There is a free tier. |
| **NotebookLM** | Copy and paste. NotebookLM has no public API. | A normal NotebookLM account. This is ideal when you've already added your lesson sources to a notebook, because the questions stay grounded in those sources. |
| **Chalkie** | Copy and paste. Chalkie has no public API. | A Chalkie account |

**With Claude or Gemini:**
1. Paste your key under **Connection settings** and click **Test connection**.
2. Fill in the topic, the pupils, how many questions, the question types and the difficulty.
3. Optionally add source material. You can paste it, or click **Use the text selected in my document**.
4. Click **Generate questions**. The questions appear in an editable box in the Quick build format. Check them, then click **Insert**.

**With NotebookLM or Chalkie:**
1. Click **Copy prompt & open …**. The add-in copies a ready-made prompt and opens the site.
2. Paste the prompt into the site's chat.
3. Copy the reply and paste it into the **Review and edit** box, then click **Insert**.
4. If the reply comes back in a different layout, click **Tidy into worksheet format**. This needs a Claude or Gemini key, and it converts the reply automatically.

**Write missing model answers** fills in mark-scheme answers for any written questions that don't have one yet. The answers then appear in the teacher copy and the answer key.

Things to know:
- **API keys** are saved only in Word's browser storage on your computer and are sent only to that AI company. On a shared computer, click **Forget key** when you've finished.
- **Pupils' personal data:** don't put pupils' names or personal details into prompts, and follow your school's AI policy.
- **Checking:** AI can get things wrong, so check every question and answer before printing.
- **Model names** change over time. If you see "model not found", click **Load models** and pick one from the list.

## Quick build format

```
# Part A – Multiple choice                 ← section heading
Q: What is the capital of France?           ← a new question (also "1." or "Q1.")
* Paris                                     ← * marks the correct option
- London
- Berlin

Q: Which of these are prime? (2 marks)      ← two or more * = "tick all that apply"
- 4
* 7
* 11

Q: Water boils at 100 °C at sea level.
A: true                                     ← true/false question

Q: The [Sun] is at the centre of the [solar system].   ← [ ] = blanks
wordbank: yes

Q: Match each animal to its group.
Dog = Mammal                                ← matching pairs (right side is shuffled)
Salmon = Fish

Q: Explain how plants make food. (4 marks)
A: Photosynthesis …                          ← model answer (teacher copy and key)
lines: 8                                    ← optional number of writing lines
```

Other options:

```
shuffle: yes             shuffle multiple-choice options
space: grid | box | lines   answer space for a written question
(0 marks)                a prompt with no marks, e.g. in a planner
> Some text              a paragraph of text, e.g. a reading passage
[ ] I can …              a tick-box checklist item
=== Lesson 3: Friction   start a new lesson page (lesson packs)
LO: …   vocab: a, b, c   learning objectives and key vocabulary for that lesson
```

You can use `**bold**` and `*italic*` anywhere.

## Good to know

- **An add-in is missing from SHARED FOLDER, shows an old version, or keeps disappearing from the ribbon:** close Word and run `./tools/reset-addin-cache.ps1`. It checks your catalog settings and manifests, then clears Word's add-in cache. Reopen Word, go to **Home → Add-ins → Advanced → SHARED FOLDER**, click **Refresh**, and add each add-in once.

- Each question is a Word **content control**, which shows as a thin box when you click in it. You can move questions around freely, then click **Renumber & update totals**.
- To change a question, click inside it and choose **Tools → Edit in wizard**. Text you type directly inside a question box is replaced when you switch between the student and teacher copies or apply a new style. Text you type outside the question boxes is never touched.
- The question data is saved inside the .docx file, so you can reopen and edit a worksheet later, or on another computer.
- **Clickable tick boxes** work in Word for Windows, Mac and the web. If you print or export to PDF, you can switch **Style → Tick boxes** to **Print only**.
- You need Microsoft 365, Word 2021 or later, Word for Mac, or Word on the web. Some older one-off purchases of Word 2016 or 2019 for Windows use an old built-in browser that can't run the add-in.

## Files

```
manifest.xml         tells Word where the add-in lives
web/taskpane.html    the side panel
web/taskpane.js      panel logic and Word commands
web/ooxml.js         layout engine: styles, questions, tick boxes, answer key
web/ai.js            AI services: Claude, Gemini, NotebookLM and Chalkie
web/templates.js     template catalogue, base designs and colour schemes
web/docs.js          reads scheme-of-work files (Word, PowerPoint, Excel, text)
web/vendor/          fflate unzip library (MIT licence)
examples/            sample worksheets, a lesson pack and the design gallery
web/taskpane.css     panel look
server.js            local HTTPS server (npm start)
tools/set-url.js     makes manifest-hosted.xml for a hosted copy
tools/reset-addin-cache.ps1  checks your add-in catalogs and clears Word's add-in cache
push.ps1             creates the GitHub repo, publishes to GitHub Pages, pushes updates
.github/workflows/   GitHub Actions workflow that publishes the web folder
```
