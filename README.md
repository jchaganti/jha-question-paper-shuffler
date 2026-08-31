# Question Paper Shuffler

Desktop app (Electron + TypeScript) that turns one NEET-style Word question paper into
*N* shuffled sets, each with its own correct answer key.

- Questions are re-ordered **within a subject only**.
- Option labels stay where they are — `(A)`–`(D)`, `(1)`–`(4)` or `(i)`–`(iv)`, whichever the
  paper uses; only the answer text moves between them.
- The answer key at the end of each generated paper is rewritten **in the original
  table**, so its format, borders, spans and fonts are identical to the source.
- Every generated file is re-opened and verified before the run is reported as done.

---

## Getting started

```bash
npm install
```

```bash
npm start
```

`npm start` builds and launches the desktop app. Pick the paper, choose what to shuffle,
press **Dry run** to see what would happen, then **Generate**.

The window opens with a collapsed strip, *How to write the Word document so that every
question can be shuffled* — plain-language authoring guidance for whoever types the paper.
It is the practical version of the assumptions listed further down this file; the dry run
points at it whenever it has to skip a question.

### The two buttons

**Dry run** writes nothing. It reports how many questions are free to move and how many
will have their options shuffled, broken down per subject; any mistake in the exclusion
lists, such as a question number this paper does not have; and then **three numbered
findings**, always in this order and always shown, even when a count is zero:

1. **Question(s) whose options cannot be shuffled with certainty** — these could not be
   parsed, so they keep their original option order and their answer never changes. Each
   one says what stopped it and what to change in Word.
2. **Question(s) shuffled, but worth correcting in the Word document** — shuffled
   correctly, but their layout had to be worked out; see
   [Shuffled, but worth correcting](#shuffled-but-worth-correcting) below.
3. **Question(s) worth keeping in their original option order** — their option text looks
   position-dependent ("None of these", "Both (A) and (B)"), with a
   one-click **Add these N to "keep the option order"** button. Nothing is wrong with them;
   this one is your call, not the tool's.

The three answer three different questions — *what could not be read*, *what was read but
is worth tidying*, *what was read fine but may not mean the same once its options move* —
so they are never merged.

**Generate** writes the files, showing two progress bars: overall (`Set 2 of 4 · 1
generated`) and the current set's step (`Shuffling options — 36% of this set`). The steps
are: reading, shuffling options, re-ordering questions and updating the answer key,
packaging, writing, verifying.

### If `npm install` cannot download Electron

On networks that block `github.com`, install with a mirror and repair the unpack step:

```bash
ELECTRON_MIRROR=https://registry.npmmirror.com/-/binary/electron/ npm install
```

```bash
npm run fix:electron
```

(`fix:electron` extracts the already-downloaded Electron zip from the local cache. It is
a no-op once Electron is in place. Everything except launching the window — the parser,
the shuffler, the tests, the CLI — works without Electron.)

### Tests

```bash
npm test
```

265 unit and end-to-end tests. They build question papers in memory (see
`tests/support/PaperFixture.ts`), so they run without any sample document.

To review the UI without Electron and without a real paper:

```bash
npm run ui:harness
```

That writes `dist/web/renderer/harness.html` — the real HTML, CSS and renderer driven by a
stubbed bridge with sample data, so the dry-run panel, both progress bars and the results
list (including a failed verification) can be opened in any browser.

### Command line (same engine, no window)

```bash
npm run cli -- --file "C:\papers\MTP-2.docx" --inspect
```

```bash
npm run cli -- --file "C:\papers\MTP-2.docx" --dry-run --sets 4 --shuffle-questions --shuffle-options --keep-option-order 30,36
```

```bash
npm run cli -- --file "C:\papers\MTP-2.docx" --sets 4 --shuffle-questions --shuffle-options --keep-question-positions 10,11,19,20 --keep-option-order 30,36,59 --seed march-batch
```

---

## Inputs

| Input | Meaning |
| --- | --- |
| Question paper | Path of the original `.docx`. It must contain the answer key at the end. |
| Shuffle question order | Re-orders questions inside each subject. |
| Keep these question numbers in place | Comma separated printed question numbers (`10,11,19,20`). Ranges like `10-14` also work. These questions stay at their exact position. |
| Shuffle option order | Permutes the four option contents of each question. |
| Keep the option order of these questions | Comma separated printed question numbers whose options must not move. Use it for "None of these", "Both (A) and (B)", or any question where option order carries meaning. |
| Number of sets | 1–100. |
| Seed | Optional — see below. Leave it blank for a normal run. |
| Keep each question on one page | On by default. A question that no longer fits at the bottom of a page starts on the next page instead of being split across two. See below. |

Exclusion lists always use the **printed question numbers of the original paper**
(1…180 for the sample papers), not per-subject numbering.

### The Seed field

The shuffle is not truly random: it comes from a pseudo-random generator, and the seed is
that generator's starting point. Same seed + same paper + same settings = the same sets,
question for question and option for option.

- **Leave it blank** for normal use. A fresh random seed is invented for the run, so every
  run produces different sets.
- **The seed used is always reported back** — shown under "Generated sets" with a *Reuse
  this seed* button, printed by the CLI, and recorded in `_generation-report.pdf`.
- **Type a seed** (anything: `march-batch`, `2026-08-02-mock-2`, or a seed you were given
  back) to reproduce a run exactly.

Why that matters in practice:

- a set was deleted, or you need more copies of *Set 03* — regenerate the identical file
  instead of issuing a paper that differs from the one students already sat;
- a dispute about which paper was used on a given date — the seed in the report, with the
  original paper, reproduces it byte for byte;
- comparing two configurations (different exclusion lists, say) without the shuffle itself
  changing underneath you.

Each set inside a run derives its own seed from the run seed (`<seed>#set1`, `<seed>#set2`,
…), which is why the sets in one run differ from each other. Those derived values are noted
per set in the report; the one you re-enter is the run seed.

A seed is not a secret or a password — anyone with the seed and the original paper can
reproduce the sets. The report already contains the full answer mapping, so keep the whole
`question-sets - …` folder as confidential as the paper itself.

### Shuffled, but worth correcting

Several option layouts are readable only because the parser has a fallback for them. Those
questions are shuffled normally and their answer letters are correct — but a fallback is a
*reading* of an ambiguous document, not a certainty, so the dry run lists them for whoever
types the paper. They used to be reported as skips; now that they succeed they are still
reported, grouped by problem with the fix stated once:

| Reported as | What it means | Fix in Word |
| --- | --- | --- |
| Some options are lettered by Word and the rest typed by hand | Option (A) is an automatically lettered list item; (B), (C), (D) are typed. | Letter all four the same way — all by Word, or all typed. |
| An option label has no tab in front of it | `…(i), (iv) and (v) (D) …` — the tab before `(D)` is missing, so the label was read from the punctuation in front of it. | Press Tab before each option label. |
| The four option labels mix capital and small letters | `(A) (b) (C) (d)`. | Use one case for all four. |
| More than one lettered list could have been the options | A question has two four-item lettered lists; the bracketed one was taken as the options. | Letter only the options `(A) (B) (C) (D)`; give other lists a different style, such as `A.` or `(i) (ii)`. |

Counts on the sample papers, out of 180 questions (100 for Animal Kingdom):

| Paper | Options not shuffled | Shuffled, worth correcting |
| --- | --- | --- |
| Group A | 6 | 4 |
| Group B | 8 | 2 |
| Group C | 9 | 3 |
| Group D | 6 | 4 |
| Animal Kingdom | 5 | 5 |

A question is reported whatever the settings say — the document is worth fixing even when
that question is in the "keep the option order" list. The same grouping goes into
`_generation-report.pdf` and into `--dry-run` / `--inspect` on the command line.

### Keep each question on one page

Shuffling moves questions of different heights into new positions, so a question that sat
comfortably on one page in the original can end up straddling a page break — stem at the
bottom of page 4, options at the top of page 5. With this option on (the default) that
cannot happen: a question that no longer fits moves down as a whole and starts at the top
of the next page.

Nothing is measured — where the page breaks fall is Word's decision, made from font
metrics and image sizes at layout time. Instead each question is marked with the two flags
Word itself provides for this (`w:keepNext` between its paragraphs, `w:keepLines` inside
each one, and `w:cantSplit` on the rows of any table inside the question), and Word does
the rest when it lays the paper out.

Measured on the sample papers, one set each, same seed with the option off and on:

| Paper | Questions split across two pages | Pages |
| --- | --- | --- |
| Group A, option off | 13 | 30 |
| Group A, option on | **0** | 32 |
| Group C, option off | 20 | 30 |
| Group C, option on | **0** | 32 |
| Animal Kingdom, option off | 18 | 26 |
| Animal Kingdom, option on | **0** | 30 |

So the cost is a few extra pages, which is why it is a checkbox rather than a fixed rule.
Two limits, both deliberate:

- A question taller than a page cannot be kept together. Word ignores the request in that
  case and breaks the question as it would have anyway — nothing is ever lost.
- Blank paragraphs that trail a question are treated as spacing, not content, and stay out
  of the chain. The page break is free to fall in that gap, which keeps the page count from
  growing more than it must.

Turn it off with `--allow-page-splits` on the command line.

### Each subject and the answer key start a new page

Always, with no setting to turn it off. Both are structural: a subject that begins halfway
down a page reads as a continuation of the previous one, and an answer key printed under
the last question is too easy to hand out with the paper.

Papers achieve this with a run of blank paragraphs, which works for the original but not
after shuffling: the text reflows, and blank paragraphs are only worth whatever space is
left on the page. Group C has no blank paragraphs at all before its answer key, so its
generated sets ran the key on straight after the last option. The generated paper therefore
states the intent — `w:pageBreakBefore` on each subject heading and on the "ANSWER KEY"
heading (or, in a paper that has no such heading, on the first paragraph inside the key
table).

Nothing is added when a break is already there — a typed page break, a next-page section
break, or the property itself — so no paper gains a blank page. Two more cases are left
alone:

- **The first subject when it opens the document.** All five NEET samples start with
  `PHYSICS` as the very first paragraph (the course/test/date line lives in the Word page
  *header*), so it already opens page 1 and asking for a break there is how a document
  gains a leading blank page. A paper with a cover line above its first subject *does* get
  the break, so the cover keeps a page to itself.
- **A paper with no subject headings.** Its single section covers the whole paper and its
  first paragraph is the paper's own title, not a subject, so there is nothing to mark.

Measured with Word on a Group C set, same shuffle, with the breaks stripped and applied:

| | Without | With |
| --- | --- | --- |
| PHYSICS | p1, opens the page | p1, opens the page |
| CHEMISTRY | p7, **mid-page** | **p8**, opens the page |
| BIOLOGY | p14, **mid-page** | **p16**, opens the page |
| ANSWER KEY | p32, opens the page | p33, opens the page |
| Pages | 33 | 34 |
| Pages with no text | none | none |

Across the sample folder: the five multi-subject papers get a break on `CHEMISTRY` and
`BIOLOGY`; the four single-topic papers are untouched.

## Outputs

Written next to the source paper, in a new folder:

```
<paper folder>/
  question-sets - 31-08-2026-13-21/       <- one folder per run; never overwritten
    <paper name> - 31-08-2026-13-21-Set-01.docx
    <paper name> - 31-08-2026-13-21-Set-02.docx
    _generation-report.pdf
```

The date and time are `DD-MM-YYYY-HH-MM` on a 24-hour clock, read from the computer's own
clock when the run starts. **The folder and every set inside it carry the same stamp**, so
a batch stays together in a folder listing even when the run crosses a minute boundary, and
sets from two runs of the same paper can never be confused. The minutes are separated with
`-` rather than the `:` of a clock reading because Windows does not allow `:` in a file
name; the report, which is not a file name, writes it `31-08-2026 at 13:21`. Two runs
started in the same minute are kept apart by a `-02`, `-03` on the second folder onwards.

Each generated paper:

- has its questions renumbered automatically by Word (numbering is a Word list, so
  re-ordering blocks renumbers them);
- carries the answer key for **that** set, in the original table, with the set label
  appended to the answer-key title (`MODEL TEST PAPER-2 (A)  –  SET 01`);
- keeps every embedded MathType/OLE equation, image, table, style, header and footer
  from the source, because the tool edits `word/document.xml` and copies every other part
  of the package byte-for-byte;
- keeps every question whole on one page (see below);
- starts each subject, and the answer key, on a page of its own (see below).

`_generation-report.pdf` records, per run: the run seed, the paper structure, the questions
whose options could not be shuffled (and why), the questions worth excluding, the
verification results, and the full `new question -> original question -> answer`
mapping for every set.

The report names options the way the paper names them, reading the style from the paper's
own answer key: a paper whose key is written `A B C D` gets `A`, `C` and `A→C, B→D, …`,
while one written `1 2 3 4` gets `1`, `3` and `1→3, 2→4, …`. The mapping table can
therefore be read straight against the generated set's key without translating anything.

---

## How it works

```
DocxPackage        read/write the .docx zip; only word/document.xml is modified
  NumberingIndex   which Word lists number questions (decimal) vs letter options
  PaperParser      body nodes -> subjects -> question blocks (+ answer key table)
  OptionSetParser  the four options, whichever way the paper labels them:
    OptionBlockParser        typed "(A) ... (B) ..." -> label / content / padding atoms
    AutoLetteredOptionParser one option per Word-numbered paragraph
  ShufflePlanner   pure plan: question order per subject + option permutation per question
  SetBuilder       applies a plan: swap option contents, re-order blocks, rewrite the key
  PageFlowGuard    marks each question so a page break cannot cut it in half, and starts
                   each subject and the answer key on a page of its own
  SetVerifier      re-opens the written file and proves it is correct
GenerationService  inspect / dryRun / generate; used by both the UI and the CLI
```

The pieces are independent: `ShufflePlanner` knows nothing about Word, `SetBuilder`
knows nothing about randomness, and `SetVerifier` re-derives everything from the written
file rather than trusting the builder.

### Question blocks

A "question block" is the numbered paragraph plus every body node up to the next
numbered paragraph — option paragraphs, tables, blank spacing paragraphs and anchored
pictures included. Shuffling moves whole blocks, so a question's diagram and match-the-
columns table always travel with it.

### Option slots

Papers label options in two quite different ways, so each layout has its own reader behind
one interface (`OptionSet`), and the rest of the application does not care which it holds.

**Typed labels** — `(A) … (B) …` written into the paragraph text. Option text is not stored
one-option-per-run in real papers: a single run can hold a tab *and* the next option's
label, and an option's value can be an OLE equation. The parser therefore breaks the option
paragraphs into "atoms" (one inline element each, carrying the original run properties),
identifies the `(A)`–`(D)` labels among them, and splits each option into

- **label** — never moved,
- **lead** — the tabs between the label and the text, never moved,
- **content** — the part that gets permuted,
- **padding** — trailing tabs and blank paragraphs that keep the columns aligned, never
  moved.

Labels are searched for in several passes, safest first (see assumption 6): upper case
before lower case, exact position before relaxed position. The first pass that yields a
clean A, B, C, D wins, and a pass that yields only B, C, D is accepted when a single
Word-lettered list paragraph supplies option (A).

Adjacent runs with identical properties are merged again afterwards, so the generated XML
stays close in size and shape to the original.

**Auto-lettered** — one option per paragraph, with the `(A)` produced by Word's numbering
rather than typed. Here the letter comes from the paragraph's *position in the list*, so
shuffling moves the content between the four paragraphs and leaves each paragraph's
numbering properties exactly where they are; Word then re-letters them (A) to (D) in place.
Nothing has to be renumbered by hand, and the same question can safely hold a lettered
statement list alongside its options.

Typed labels are tried first because they are unambiguous. When both readers refuse a
question, the more specific of the two complaints is the one reported.

### Answer key

The original key is read into `question number -> letter` before anything is changed.
For every position in the new paper the tool knows which original question landed there
and how its options were permuted, so the new letter is
`letter of the slot that now holds the originally correct content`. Letters are written
back into the existing cells; no table is ever rebuilt.

### Verification

After writing each file the tool re-opens it and checks:

1. the question count is unchanged;
2. every question in the generated file can be matched to an original question by
   content (the fingerprint is invariant under option shuffling, and includes symbol
   glyphs, embedded-object ids and math text, so `-11` and `11` are not confused);
3. each question still offers exactly its original four options — nothing lost,
   duplicated or altered;
4. the letter in the new key points at the *same text* that was correct in the original.

A failing check is shown in the UI and written to the report; the file is still kept so
it can be inspected.

---

## Assumptions

These held for all five sample papers. They are validated at parse time and produce a
clear error message when they do not hold.

1. **Question numbers are Word automatic numbering**, not typed text. Papers may use
   several numbering lists (the samples use one for Physics starting at 1 and another for
   Chemistry + Biology starting at 46); a list is treated as question numbering when it is
   decimal *and* starts exactly where the previous one stopped.
2. **Subjects are announced by a paragraph whose entire text is the subject name**
   (`PHYSICS`, `CHEMISTRY`, `BIOLOGY`, `BOTANY`, `ZOOLOGY`, `MATHEMATICS`, `MATHS`). A
   paper without such headings is treated as one single subject, and shuffling then spans
   the whole paper.
3. **The paper ends with an answer key** introduced by an `ANSWER KEY` paragraph (or, as a
   fallback, the first table whose cells pair a number with an answer). The key must
   cover every question exactly once; the number of key entries is what defines the
   question count. Everything from the `ANSWER KEY` heading onwards is never shuffled.

   An answer box may name its option by letter (`A`, `(A)`, `a.`), by position as a digit
   (`1`, `2.`) or by position as a roman numeral (`iii)`, `(IV)`). The style is read off
   the box itself and the new answer is written back in that same style, so a key comes
   back written the way its author wrote it. A key names its options one way throughout,
   so the style used by most of its boxes is taken as that table's style and a box written
   another way is not read — which is also what tells an answer key apart from any other
   grid of numbers in the document.

   A number box may be written `1`, `1.` or `1)` — the ordinal punctuation authors
   habitually type names question 1 and nothing else, so all three are read. Anything else
   after the number (`57,` for `57.`) is a slip rather than a convention: the box no longer
   names a question, and the tool reports which box to retype instead of deciding for
   itself which trailing characters are decoration. A key that skips a number inside its
   own run is reported as incomplete before the question numbering is examined, so the
   error names the key rather than blaming the numbering.
4. **Printed question numbers come from the key**, in ascending order, matched to the
   questions in document order. Because questions only move inside a subject, the numbers
   printed at each position never change.
5. **Blank paragraphs at the end of a subject are page spacing**, not part of the last
   question, so they stay behind when that question moves. Blank paragraphs *inside* a
   question (used to reserve room for a floating diagram) travel with it.
6. **A typed option label is `(A)`/`A)` at the start of a paragraph or straight after a
   tab** — and, failing that, after a space whose preceding character is not alphanumeric
   (papers do lose the tab: `…(i), (iv) and (v) (D) …`). Requiring a non-letter is what lets
   `(C) Both (A) and (B)`, `Assertion (A):` and `1s22s2(C)` be read correctly.

   The label *scheme* is deduced per question rather than assumed: letters `(A)…(D)`,
   digits `(1)…(4)`, or roman numerals `(i)…(iv)`. Schemes are tried in order of
   decreasing certainty that a run of labels really is the options — upper-case letters,
   lower-case letters, either case, digits, then roman numerals — and the first that
   yields a clean run of four wins. That ordering is what makes a match-the-columns
   question work: one that lists items as `(a)…(d)` and answers as `(A)…(D)` yields the
   answers, and one that lists items as `(i)…(iv)` and answers as `(1)…(4)` also yields
   the answers. Whichever scheme is found, the labels themselves are never moved, so each
   generated set keeps the tokens its author typed. Every scheme also recognises its fifth
   label, so a five-option question is refused rather than having its fifth option folded
   into the fourth.

   Where pass order alone is not enough, **the answer key decides**: the scheme the key
   names its answers in is tried first. That settles the match-the-columns question whose
   column entries are lettered `(a)…(d)` above real options written `(1)…(4)` — reading
   letters first would shuffle the column entries. Nothing is guessed; the paper states
   that its answers are digits, so the digit labels are its options.

7. **A floating picture anchored among the options is layout, not text.** Word places such
   a picture from the *page*, so the words in its text box are not part of the sentence a
   reader sees, and the picture does not travel with the text beside it. It is therefore
   left out when labels are looked for — otherwise a diagram's stray letters (`O A B Cl`)
   sit in front of the next label and hide it — and it stays in its own paragraph while
   the answers move around it. A picture in the *middle* of one answer is refused instead:
   there the words would move to another slot and the picture would stay behind. Questions
   shuffled with a picture among their options are listed in the report so the author can
   confirm that no picture belonged to one particular option.
8. **Options may be lettered by Word instead of typed** — either all four (four list
   paragraphs, first is (A)), or just the first, with `(B)`, `(C)`, `(D)` typed; in the
   latter case a single-item lettered list at or before `(B)` is option (A). A question may also carry a lettered
   *statement* list, so the list written `(%1)` wins over one written `%1.`, and upper case
   wins over lower case; when that still leaves two candidates the question is skipped
   rather than guessed at.
9. **Only the four option contents move; labels, tabs, spacing and paragraph boundaries do
   not.** Whitespace at the edges of an option is separator, not content — including a
   space inside the same run as the option text, which is how four options on one line are
   often parted (`…(iii) (D)…`). Letting it travel would glue the next label onto whatever
   landed there. Shuffling therefore changes the order of words in a question and nothing
   else; the tests assert that the multiset of words is identical before and after.
   Column alignment is preserved structurally, but because options differ in length the
   text after a tab stop can sit slightly differently, and a paper may gain or lose a page
   through reflow.
10. **A question is skipped, never guessed at.** If the four options cannot be identified
   with certainty, that question keeps its original option order and is listed in the
   report. This happens for options laid out inside a table, options that continue onto
   another paragraph, an option whose whole answer is a floating picture or that has one in
   the middle of its words, questions where two lettered lists could equally be the options, and
   labels whose opening bracket was inserted from a symbol font — that bracket prints as `(`
   but holds no text, so where the option before it ends cannot be established. The missing
   bracket is the whole signal there: only a label that reads as `A)` is suspect. A label
   that reads as `(A)` has its bracket, so a symbol in front of it is the previous option's
   *content* — an answer of `60Ω` or `15°` sitting last in its column, which after a shuffle
   can land in front of any label — and the question shuffles normally.
   Where a fallback in assumption 6, 7 or 8 *did* make the options readable, the question is
   shuffled but still reported — see
   [Shuffled, but worth correcting](#shuffled-but-worth-correcting). Nothing the tool works
   out from an inconsistent layout stays invisible.
11. **Semantics are the user's call, and the test is position dependence.** The tool never
    decides that an option "should not" move. It flags a question only when an option's
    meaning points at something *outside itself* — the options above it ("None of these"),
    another option by name ("Both (A) and (B)"), or wording printed once as directions
    above a run of questions — and the UI offers to add those to the exclusion list in one
    click. It is not flagged for being a *kind* of question: the standard Assertion-Reason
    options ("Only statement I is true", "Assertion and Reason are true and Reason is the
    correct explanation of Assertion") each state their own meaning in full, so they move
    safely and are left alone. The flag reads the **options**, never the stem — a stem that
    happens to say "all of the above" says nothing about whether the options can move.
12. **Every set differs** from the original paper and from the other sets in the same run
    (checked over 50 attempts), and no question keeps its original four options in the same
    order unless it was excluded or unshufflable.
13. **Page breaks are Word's to place, not the tool's.** "Keep each question on one page"
    adds `w:keepNext` / `w:keepLines` / `w:cantSplit` and lets Word lay the paper out; the
    tool never measures a page and never moves content to make one fit. The only breaks it
    asks for are `w:pageBreakBefore` on each subject heading and on the answer key, and only
    where no break is there already. Flags in the source document are left alone.

## Will it work on my paper?

Press **Dry run**: it states the question count, the subjects it found, what would be
shuffled, and every question whose options it will not touch — without writing anything.
That is the answer for a specific paper. In general:

| Variation | Behaviour |
| --- | --- |
| `.doc` (Word 97-2003 binary) | **Rejected.** Save as `.docx` first. |
| Continuous numbering across subjects, one or several Word lists | Supported (both sample formats). |
| Labels `(A)` / `A)` / `(a)` | Supported. |
| Labels written as digits `(1)…(4)` or `1)…4)` | Supported — the scheme is deduced from the question and the labels stay as typed. |
| Labels written as roman numerals `(i)…(iv)` or `(I)…(IV)` | Supported — tried after letters and digits, so an item list is not mistaken for the options. |
| 1, 2 or 4 options per line, or one option per paragraph | Supported. |
| Options with equations, superscripts, symbols, inline images | Supported - the run moves with its content. |
| A floating picture anchored among the options (a diagram beside or between them) | Supported for both typed and Word-lettered labels - the answers move, the picture stays exactly where it is, and the question is **reported** so you can check it. |
| A floating picture in the middle of one option's answer | Question keeps its option order and is listed in the report - the words would move and the picture would not. |
| Options whose answers are **inline** pictures (`In line with text`) | Supported - the picture moves with its option like any other content. |
| Options whose answers are **floating** pictures | Question keeps its option order and is listed in the report, naming the fix: set each option picture to *In line with text*. |
| Spacer paragraphs, a `SECTION B` instruction or the next question's artwork after the last option | Supported - a blank paragraph ends the option list, so trailing material is not read as part of the last option. |
| Stem spread over several paragraphs, match-the-columns tables, diagrams | Supported - they travel with the question. |
| No subject headings (or headings the tool does not recognise) | Treated as **one** subject, so shuffling spans the whole paper. Visible in the summary panel as a single `ALL` row - check it. |
| Options auto-lettered by Word (one option per list paragraph, no typed "(A)") | Supported. |
| One option lettered by Word with the other three typed, at any position (`(A) (B) [Word letters this one] (D)`) | Supported, and **reported** as worth correcting. |
| A missing tab before a label (`…and (v) (D) …`) | Supported, and **reported** as worth correcting. |
| Two options on a line pushed apart with spaces instead of a tab (`…in transmission␣␣␣␣␣(B) …`) | Question keeps its option order and is listed in the report, naming each label to put a tab in front of. |
| Labels mixing case (`(A) (b) (C) (d)`) | Supported, and **reported** as worth correcting. |
| Item list `(a)…(d)` above the answer options `(A)…(D)` | Supported — the upper-case run is the options. |
| A lettered statement list ("A. …") next to the option list ("(A) …") | Supported — the bracketed list is the options — and **reported** as worth correcting. |
| Two lists that could equally be the options | Question keeps its option order and is listed in the report. |
| Options laid out inside a table, or a mix of one auto-lettered option and typed labels | Question keeps its option order and is listed in the report. |
| 3 or 5 options, options continuing onto another paragraph, an option whose whole answer is a floating picture | Question keeps its option order and is listed in the report. |
| A label whose opening bracket was put in with *Insert → Symbol* | Question keeps its option order and is listed in the report. |
| An option whose *content* is a symbol (`60Ω`, `15°`), sitting right before the next label | Supported - the label has its own bracket, so the symbol is content and moves with it. |
| Answer key numbers written `1.` or `1)` | Supported. |
| Answer key answers written `A`, `(A)`, `a`, `1`, `2.`, `iii)` or `(IV)` | Supported — the key's own style is deduced, and both the rewritten key and the generation report use it. |
| Question numbers typed by hand instead of Word numbering | **Hard error** naming the lists it did find. |
| Each subject restarting numbering at 1 | **Hard error** - a key entry would no longer identify one question. |
| Answer key missing, or with number and answer in one cell | **Hard error** - the key is not detected. |
| A stray character in a key number box (`57,`) | **Hard error** naming the box to retype - the key is one answer short. |
| Fewer than five questions | **Hard error** - too small to recognise a key table. |

The design rule behind that table: the tool either does the right thing, refuses one
question and says so, or refuses the whole paper and says why. It is not designed to
guess. `tests/variations.test.ts`, `tests/autoLetteredOptions.test.ts` and
`tests/labelHeuristics.test.ts`, `tests/optionLabelFormats.test.ts` and
`tests/answerKeyFormats.test.ts` lock each row above in place.

The one case self-verification cannot catch is a paper structured so differently that the
*parse itself* is wrong in a consistent way (verification re-uses the same parser). That is
why the dry run leads with the question count and subject breakdown: if those numbers match
your paper, the parse is right.

## Known limitations

- Options that live inside a table are reported and left alone rather than shuffled, as are
  the few questions where a floating picture's text box runs into an option label. In the
  four MTP sample papers this affects 6–9 of 180 questions; in the `ANIMAL KINGDOM` sample,
  5 of 100.
- Only four options (A–D) are supported. A five-option question is detected and skipped,
  not mangled.
- Tab-aligned option columns can shift by a few millimetres, and total page count can
  change by a page, purely because the option texts have different widths.
- Two questions that are byte-identical (same stem *and* same options) cannot be told
  apart when verifying; the check reports them instead of failing silently.
- The tool does not renumber anything that was typed by hand (for example a stem that
  says "in question 14 above").

## Repository layout

```
src/shared/types.ts        contracts shared by UI, main process and CLI
src/shared/layoutNotes.ts  grouping of option-layout notes for display
src/core/docx/             .docx package + XML helpers
src/core/parse/            numbering, paper structure, answer key
src/core/options/          the two option layouts, atoms, applying a permutation
src/core/shuffle/          seeded RNG and the pure shuffle planner
src/core/generate/         set builder, page flow, output folder, report, orchestration
src/core/verify/           post-generation verification
src/main/                  Electron main process + preload bridge
src/renderer/              UI (HTML/CSS/TS, no framework)
src/cli/                   headless entry point
tests/                     unit + end-to-end tests with an in-memory paper fixture
scripts/                   build helpers, Electron repair, diagnostics
```

Diagnostics used while developing, handy when a new paper misbehaves:

```bash
node scripts/diag-question.cjs "paper.docx" 35
```

```bash
node scripts/diag-dupes.cjs "paper.docx"
```
