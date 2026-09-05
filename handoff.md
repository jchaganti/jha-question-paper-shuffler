# Handoff — Question Paper Shuffler

Written 2026-08-30. Covers the project brief, the architecture, every decision taken so
far and the exact next steps. Read this plus `README.md` before changing anything.

---

## 1 · Project brief

An **Electron + TypeScript desktop app** that takes one NEET-style Word question paper
(`.docx`) and produces *N* shuffled sets, each with its own correct answer key.

- Questions are re-ordered **only within their own section** - a subject, or a `SECTION A` /
  `SECTION B` division of one (item 28).
- Option labels `(A)`–`(D)` stay put; only the option *contents* move between them.
- The answer key is **edited in place** — only the letter inside each existing cell is
  rewritten, so the table keeps its geometry, borders, spans and fonts.
- Only `word/document.xml` is modified. Every other zip part is copied byte-for-byte.

**Users**: the people who type these papers. They are not technical. Every message the
tool emits is written in their terms and says what to change in Word.

### The invariant everything else serves

> Shuffling changes **no character** of any question. It only re-orders whole option
> contents and whole question blocks.

The multiset of words inside a question's option region must be identical before and
after. This is enforced at three levels: `splitLeadCorePad` (whitespace ownership),
`SetVerifier` (per-file re-parse and fingerprint match), and tests.

### The standing rule that governs all design decisions

The user stated this and it overrides convenience:

> "I do not want to add any hacks in the code when something can be fixed in the input
> question paper by the user manually. We should strictly follow this rule. We want to
> standardize the input. So if there is something non-standard, we will let user know how
> to standardize it and also add it to the best practices section. So if the current fix
> is a hack, lets not fix it - rather we will let the user know to fix it in the input
> question paper manually."

**How to apply it.** Before writing a fix, ask: *is this reading a standard form
correctly, or is it guessing past an ambiguity?*

- Reading `1.` and `1)` as question number 1 — **not a hack**. One unambiguous reading,
  and it is the form authors actually type. Implemented.
- Reading `57,` as 57 — **a hack**. Requires deciding which trailing characters are
  decoration. Refused and reported instead.
- A "loose fingerprint" fallback so verification passes on a malformed document —
  **a hack**. Was built, then reverted, because it made a defective document report PASS.

When you refuse something, you owe the user three things: a precise error naming the
exact box/label to fix, an entry in the in-app **best practices** list
(`src/renderer/index.html`), and a row in the README's *Will it work on my paper?* table.

---

## 2 · Build, run, test

```bash
npm run build        # tsc main (CommonJS) + renderer (ESM) + copy static assets
npm start            # build, then launch Electron
npx vitest run       # 326 tests, 26 files
npm run portable     # release/Question Paper Shuffler <v> (portable).zip - shareable, no archiver needed
npm run dist         # release/... Setup <v>.exe - needs a 7za the machine will run, see item 26
npm run cli -- --file "paper.docx" --shuffle-questions --shuffle-options --dry-run
npm run regress      # every paper in C:\ps\q-paper; add -- --generate to also verify sets
npm run audit        # every dry-run report in the corpus, checked against itself
node scripts/ui-harness.mjs   # dist/web/renderer/harness.html — real UI, stubbed backend
```

Sample papers live in `C:\ps\q-paper` (12 files). They are the real regression corpus —
the unit tests do not cover the awkwardness these documents contain.

---

## 3 · Core architecture

### Pipeline

```
DocxPackage        unzip; hand out word/document.xml + numbering.xml as XmlParts
      ↓
NumberingIndex     numId -> {format, start, levelText}
      ↓
PaperParser        body nodes -> subjects -> sections -> QuestionBlocks (+ AnswerKeyTable)
      ↓
OptionSetParser    per question: typed-label reader, else auto-lettered reader
      ↓
ShufflePlanner     seeded permutations (questions within section; options within question)
      ↓
SetBuilder         re-order blocks, apply option permutations, rewrite key, page flow
      ↓
DocxPackage.save   rezip, replacing only document.xml
      ↓
SetVerifier        re-open the written file and prove four properties
```

### Module map (`src/`)

| Path | Responsibility |
| --- | --- |
| `core/docx/` | zip + XML DOM helpers (`xml.ts` has `NS`, `visibleText`, `firstChild`, …) |
| `core/docx/SymbolFont.ts` | Adobe's Symbol encoding, so a `w:sym` reads as the character it prints |
| `core/parse/PaperParser.ts` | the structural reader; throws `PaperParseError` with user-facing text |
| `core/parse/AnswerKeyTable.ts` | finds the key grid, reads/writes letters, reports spoilt boxes |
| `core/parse/NumberingIndex.ts` | Word numbering definitions |
| `core/options/Atoms.ts` | the movable unit; `splitLeadCorePad` decides what is layout vs content |
| `core/options/OptionBlockParser.ts` | typed `(A)…(D)` labels, multi-pass label search |
| `core/options/AutoLetteredOptionParser.ts` | options lettered by Word numbering |
| `core/options/OptionShuffleApplier.ts` | rewrites paragraphs; label/lead/pad stay, core moves |
| `core/generate/SetBuilder.ts` | builds one set end to end |
| `core/generate/PageFlow.ts` | `keepNext`/`keepLines`/`cantSplit`/`pageBreakBefore` |
| `core/generate/Signatures.ts` | question fingerprint, invariant under option shuffling |
| `core/parse/BoundaryPictures.ts` | floats anchored across a question boundary, which pin both questions |
| `core/verify/SetVerifier.ts` | re-parses the written file; the safety net |
| `core/report/Pdf.ts` | the minimal PDF writer - fonts, text, rules, xref |
| `core/report/ReportPdf.ts` | typesets a `ReportDocument` onto pages |
| `shared/answerStyle.ts` | the only boundary between internal A-D slots and the paper's own option names |
| `main/` | Electron main, preload bridge, IPC channels |
| `renderer/` | UI. **`renderer.ts` must have zero runtime imports** |
| `cli/cli.ts` | headless driver, used for all regression runs |

### Key data structures

- **`Atom`** — one inline element wrapped in its own run carrying cloned `rPr`. Fields:
  `node`, `paragraphIndex`, `text`, `blank`, `floatingGraphic`, `unreadableSymbol`. A
  `w:sym` in the Symbol font contributes its decoded character to `text` (item 27); one in a
  picture font contributes `''` but is still `blank: false`, so it is never trimmed away as
  whitespace.
- **`OptionSlot`** — `labelAtoms` / `leadAtoms` / `coreAtoms` / `padAtoms`. **Only
  `coreAtoms` move.** Label, the tabs after it, and the trailing padding belong to the
  *position*, which is what keeps columns aligned and separators in place.
- **`PaperSection`** — the unit a question may move within: a subject, or one `SECTION`/
  `PART` division of one. `subject` + `group` + `label`; `headingNode` only on the division
  that opens its subject, since that is what drives the page break.
- **`QuestionBlock`** — a slice of body nodes, plus `printedNumber` (taken from the key).

### Constraints that bite

- **ECMA-376 child order.** `w:pPr` and `w:trPr` children must appear in schema order or
  Word refuses the file. `PageFlow.ts` has `PPR_ORDER` / `TRPR_ORDER` and splices new
  flags into the right position. Never append.
- **CSP** in the renderer is `default-src 'none'; script-src 'self'; style-src 'self';
  img-src 'self' data:`. No inline `<style>`, no `style=` attributes, no inline
  `<script>`. `'self'` does resolve file:// subresources in this Electron setup — that is
  how `styles.css`, `renderer.js` and `paper-stack.svg` all load.
- **No bundler for the renderer.** `renderer.js` is loaded directly by the browser.
  Type-only imports are erased and fine; a runtime import breaks the app, and
  `scripts/ui-harness.mjs` throws if one appears.
- **Zip timestamps differ between runs.** To prove determinism, compare the hash of
  `word/document.xml`, never the hash of the `.docx`.

---

## 4 · What was built, in order

Commits, oldest first:

| Commit | Work |
| --- | --- |
| `0d81e6b` | Initial commit |
| `cb7469e` | Bug fixes |
| `fb7c25d` | Each subject starts a new page |
| `acffc97` | Option-separator fix when four options share a line |
| `26aa84e` | Answer key fixes (this session) |
| `5719b9d` | Colour customization (this session) |

### Earlier in the session (pre-compaction)

1. **Questions no longer split across pages.** `PageFlowGuard.keepQuestionsWhole` sets
   `keepLines` on every paragraph of a block and `keepNext` on all but the last, plus
   `cantSplit` on table rows. Measured with Word COM: shuffling had created 13 (Group A),
   20 (Group C) and 18 (Animal Kingdom) split questions; all eliminated for 2–4 extra
   pages. Toggle: *Keep each question on one page*, CLI `--allow-page-splits`.
2. **Answer key always starts a new page.** Papers push the key down with blank
   paragraphs, which stops working once text reflows — Group C had none at all.
   `pageBreakBefore` states the intent. Skipped when a break already exists, so no paper
   gains a blank page.
3. **Each subject starts a new page**, by the same rule. Without it CHEMISTRY started
   mid-p7 and BIOLOGY mid-p14; with it, p8 and p16.
4. **Option-separator defect.** Two independent causes:
   - *Ours*: trailing whitespace inside a text atom travelled with the option content,
     gluing the next label — `all of the above(D)`. Fixed by `splitLeadCorePad`, which
     trims edge whitespace into the slot's lead/pad. Without it, four otherwise-clean
     papers lost characters.
   - *Theirs*: labels whose opening bracket is a Symbol-font character. Refused and
     reported, per the standing rule.

### This session

5. **Answer key not found on *Alternating Current*.** The key existed; its number boxes
   read `1.`, `2.`, `3.` and the reader only accepted bare digits.
   `AnswerKeyTable.NUMBER_RE` now accepts an optional trailing `.` or `)`.
6. **Precise diagnosis for a spoilt key box.** One cell reads `57,`. Added
   `AnswerKeyTable.spoiltNumbers` (sorted into question order, not column-scan order) and
   `gaps`, and `PaperParser.describeKeyGaps`. A hole inside the key's own run is now
   reported **before** the numbering is resolved, so the error names the key instead of
   blaming the question numbering. Message:

   > The answer key is incomplete. The key has no entry for question 57. A number box in
   > the key reads "57," - a stray character was typed after the number. Retype that box
   > as the plain number.

7. **Fixed a real defect in the symbol-bracket check** (introduced by item 4). It fired on
   any `w:sym` sitting before a label, but an option whose *content* is a symbol — `60Ω`,
   `15°`, `2 cosec 2θ` — puts one there too, and a shuffle can move one in front of any
   label. In one case the tool shuffled a question and then could not re-read its own
   output. `LabelHit.bracketed` now records whether the opening bracket was in the text;
   only a label that read `A)` is suspect. Genuine cases still refused (MTP-2-XI-2023
   keeps all 32); false positives gone, recovering ~3 questions per NEET paper.
8. **UI: pastel wash, watermark, colour picker.** Four palettes (Periwinkle, Mint, Peach,
   Warm sand), each with a dark variant, selected by `data-palette` on `<html>` because
   CSP forbids writing custom properties from JS. `paper-stack.svg` is three fanned papers
   with a *different* answer filled in on each. `--watermark-tint` hue-rotates it per
   palette. Choice persisted in `localStorage` (every access wrapped — it throws on the
   `data:` URL the harness runs from).

   **Superseded by item 24**: the picker, the other three palettes and the `localStorage`
   handling are gone. Warm sand's values sit in `:root`; the wash, the watermark and its
   tint are unchanged.

---

### This session (option/answer format deduction)

9. **The option label scheme and the answer-key answer style are now deduced from the
   paper instead of assumed.** The user's instruction: *"do not assume any format in the
   code. Deduce it based on what is there in the input paper and use it."*

   - **Answer key** (`AnswerKeyTable.ts`). An answer box may name its option by letter
     (`A`, `(A)`, `a.`), by position as a digit (`1`, `2.`) or by position as a roman
     numeral (`iii)`, `(IV)`). `parseAnswerCellText` records the scheme, case and
     decoration per box; `setAnswer` renders the new answer back through
     `renderAnswerCellText`, so the key keeps the style its author used. A bracketed key
     used to be a **hard error** and is now supported.
   - **Option labels** (`OptionBlockParser.ts`). The letter-only regex trio became
     `SCHEME_TOKENS` / `SCHEME_PATTERNS` over six schemes. `SEARCH_PASSES` runs them
     upper-letter → lower-letter → any-letter → digit → lower-roman → upper-roman, each
     strict-then-relaxed. Slots stay named A–D internally; labels are never moved, so the
     page keeps its own tokens.

   **Why that pass order matters.** It is the whole defence against reading an *item* list
   as the options. Questions routinely list items as `(i)…(iv)` or `(a)…(d)` above answers
   written another way, so the scheme most easily confused with an item list is asked last
   and only wins when nothing better yields a clean run of four. `optionLabelFormats.test.ts`
   pins both directions.

   **Key detection needed a matching guard.** Accepting digits as answers meant any grid of
   number pairs could read as a key — and `findAnswerKeyStart` scans individual tables.
   `AnswerKeyTable.scan` now tallies the schemes of a table's candidate pairs and keeps
   only the majority scheme, because a real key names its options one way throughout. A box
   dropped this way surfaces through the existing `gaps` check, which names the question.

   **Result: `FST-1-Rep-2024` was never missing a key.** It writes its answers as digits
   `1`–`4` and labels its options `(1)`–`(4)` — the filename even says `1-2-3-4`. It went
   from *blocked* to 200 questions, **166 options shuffled, verification PASSED**. Every
   other paper's *options shuffled* count is unchanged from §6, which is the check that
   proves no false-positive refusal appeared and no check stopped firing.

10. **The generation report speaks the paper's vocabulary too.** `_generation-report.md`
    still printed `A`/`B`/`C`/`D` in *Original answer*, *New answer* and *Option mapping*
    for a paper written `(1)`–`(4)`. The parse/render pair moved out of `AnswerKeyTable`
    into **`src/shared/answerStyle.ts`** (`AnswerStyle`, `parseAnswer`, `renderAnswer`,
    `PLAIN_LETTER_STYLE`) — dependency-free, so main, CLI and renderer can all use it.
    `AnswerKeyTable` now exposes `.style`; `PaperSummary.answerStyle` carries it to the
    report; `describePermutation(permutation, style)` renders the mapping column; and
    `SetVerifier`'s key-broken detail (which also lands in the report) uses it.

    **A-D remains the single internal vocabulary** — slots, permutations and
    `OPTION_LETTERS` are unchanged. `answerStyle.ts` is the only boundary between that and
    what the user reads. Sort the mapping column by *slot*, not rendered text: `"(10)"`
    sorts before `"(2)"` as a string, and lower-case letters sort after upper-case ones.

    Verified on the corpus: only FST-1's report changed to digits; all ten letter-keyed
    papers still read `A | C | A→A, B→C, …`. Note FST-1's key cells hold **bare** digits
    (`1`, `4`) while its option labels on the page are bracketed `(1)` — the report follows
    the *key*, so the three columns match the key cells they describe.

    193 tests, 15 files. New: `answerKeyFormats.test.ts`, `optionLabelFormats.test.ts`, and
    a "report names options the way the paper names them" block in `generation.test.ts`.

11. **Set file names carry the run's date and time**:
    `<paper> - 31-08-2026-13-21-Set-01.docx`. `setFileName` takes a `Date`;
    `fileNameTimestamp` renders `DD-MM-YYYY-HH-MM` in **local** time (the clock the user
    read when they pressed Generate).

    Two things worth keeping: the timestamp is read **once per run** in
    `GenerationService.generate` (`runStartedAt`) and passed to every set, so a run
    crossing a minute boundary still produces one consistently named batch; and the same
    instant now feeds `GenerationResult.generatedAt` and the report's `Generated:` line,
    which previously called `new Date()` again inside `ReportWriter` and could disagree
    with the names.

    **The user asked for `13:21`; the code writes `13-21`.** Windows forbids `:` in a file
    name (with `\ / * ? " < > |`), so a literal clock reading could not be saved and Word
    could not open it. `setFileName.test.ts` asserts the name is free of all of them - keep
    that test if the format is ever revisited.

    205 tests, 16 files. New: `setFileName.test.ts`.

12. **A floating picture among the options no longer blocks the question.** This was one
    root cause behind three different-looking complaints, all in `FST-1`:
    *"Labels read as 134"*, *"Option (D) continues on another paragraph"* (Q70, Q74), and
    *"Option (A) has no content"*.

    A floating (anchored) picture is placed **from the page**, so two facts follow, and the
    code now encodes both:

    - Its text box is **not in the reading flow**. `makeAtom` gives a floating atom
      `text: ''`, so a diagram's stray caption (`O A B Cl`) can no longer sit in front of
      `(2)` and hide it. Offsets stay consistent because every consumer reads `atom.text`.
    - It **never moves**. `splitLeadCorePad` treats a floating atom at either edge as
      layout (like a tab), and `OptionShuffleApplier.push` always returns a floating atom
      to `atom.paragraphIndex`, whatever slot's layout it ended up in.

    **The rule self-polices, which is why it is safe.** If the pictures *are* the options,
    every core comes out empty and the existing "has no content" check refuses the question
    (FST-1 Q52, Q60 - correctly). If they are diagrams, all four options have real text and
    the question shuffles. A picture left *inside* a core is still refused
    (`option-contains-floating-graphic`), now with a message saying it is in the middle of
    the answer. Questions shuffled with a picture among their options emit the new
    `floating-picture-in-option-area` layout note, so nothing worked out this way is
    invisible.

13. **The answer key now breaks ties between label schemes.** `passesFor(answerScheme)`
    reorders `SEARCH_PASSES` so the family the key uses is tried first.
    `OptionSetParser` takes the scheme and passes it to `OptionBlockParser`.

    This fixes a genuine mis-read found while testing: FST-1 Q80/Q132 are
    match-the-columns questions whose column entries are lettered `(a)`-`(d)` above real
    options written `(1)`-`(4)`. Pass order alone put lower-case letters first, so the tool
    read the *column entries* as the options. It then failed for an unrelated reason, so
    nothing was corrupted - but on another paper it could have shuffled the wrong list.

    **`SetVerifier` must pass the same scheme** (it does, from the generated file's own
    key) or it would re-read the options differently from how they were written.

14. **Every user-facing message now names options the way the paper does.** `labelAsTyped`
    renders a label as typed (`(1)`, `B)`, `(iii)`); `shownLabel(i)` in `build` uses the
    found labels, falling back to the scheme's own token. The dry run for a digit paper now
    says *"Option (1) has no content"* and *"Expected labels 1,2,3,4"*. A question that
    genuinely uses `(a)`-`(d)` is still reported as `(a)` - the deduction is per question.

    **Corpus effect** (options shuffled, vs the §6 baseline). Every paper improved or held:
    Alternating Current 93→95, Extra MTP-1 40→45, FST-1 blocked→**180**, Group A 174→175,
    Group B 172→174, Group C 171→172, Group D 174→175, Nano Physics 9→11; ANIMAL KINGDOM
    95, MTP-2-XI 155, Nano Chemistry 12 unchanged. All verifications PASS except ANIMAL
    KINGDOM Set 02 (pre-existing open issue 1).

    213 tests, 17 files. New: `floatingPictures.test.ts`.

15. **An option may *be* a picture — but only an inline one can move.** Investigated
    directly: every graphic in FST-1's image-option questions (Q52, Q60, Q63, …) is
    `wp:anchor` / VML `position:absolute` — **zero inline drawings**. Their anchors use
    `positionV relativeFrom="paragraph"`, so they *would* travel with a moved run.

    **But which picture belongs to which option cannot be read from the file.** Q52 has
    pictures in paragraphs 1,3,5,7 with labels in 3,5,7,9; Q63 has pictures in 2,4,6,8 with
    labels in 2,5,7,9. The pairing differs question to question, and settling it needs
    Word's layout engine (paragraph positions plus EMU offsets). Any rule invented here
    would risk pairing an answer with the wrong picture — the worst failure this tool has,
    since the key would then point at the wrong structure. **Refused, per the standing
    rule**, but the message now names the fix that genuinely works: set each option picture
    to *In line with text*. `imageOptions.test.ts` proves inline pictures shuffle correctly,
    including a picture mixed with words in the same option.

    If this is ever revisited, the missing input is a layout pass, not a cleverer heuristic.

16. **A blank paragraph ends the option list** (`splitTrailingBlock`). The last option is
    the only one with no following label to bound it, so it swallowed whatever trailed the
    block — spacer paragraphs, a `SECTION B (Attempt any 10 questions)` instruction, the
    next question's artwork — and was refused as *"continues on another paragraph"*
    (FST-1 Q35, Q85). Same reading the parser already applies to a subject's trailing blank
    paragraphs. A genuine continuation follows immediately with no blank between, so it is
    still refused. Blankness is taken from the block's **paragraph list**, not from the
    content atoms: a truly empty paragraph has no atoms at all.

17. **The dry run reports three numbered findings**, always in that order and always shown
    even at zero, in both the CLI and the UI (`renderFinding`, `.finding` in `styles.css`):
    (1) could not be read, (2) read but worth tidying, (3) read fine but position-dependent.
    Finding 3 used to render as a `.notice` box, which made it read as a different kind of
    thing; "kept because you asked" moved inside it.

    **Corpus effect**: FST-1 180→**182**, MTP-2-XI 155→**158**; every other paper
    unchanged; all verifications PASS except ANIMAL KINGDOM Set 02 (open issue 1).
    FST-1's unshufflable count is down from 34 at the start of this work to 18.

    223 tests, 18 files. New: `imageOptions.test.ts`.

18. **The option advisor now tests position dependence, not question type.** The user's
    call: *"if the options have no dependency on the options, then no need to even flag
    them. We will flag only position dependent ones."*

    Assertion-Reason questions were flagged for *being* Assertion-Reason - the regex ran
    over the whole block, so the "Assertion:/Reason:" lines in the **stem** triggered it.
    But the standard four options each state their own meaning ("Only statement I is true",
    "Assertion and Reason are true and Reason is the correct explanation of Assertion"), so
    they move safely and the key is remapped like any other question. Surveyed the corpus
    first: **51 AR questions, every one self-describing, zero position-dependent.**

    `OptionAdvisor.advise(paper, optionParser)` now reads the parsed **options**:
    - Catch-all and cross-reference are matched against the option text, not the stem - a
      stem saying "all of the above" says nothing about whether the options can move.
    - Assertion-Reason is flagged only when the options are **not** self-describing
      (`SELF_DESCRIBING` vocabulary: assertion/reason/statement/true/false/correct/
      incorrect/wrong/right/explanation). That is the legend layout, where the wording is
      printed once as directions above a run of questions and the options are bare tokens
      whose meaning lives outside them - moving those genuinely breaks the question.
    - When the options cannot be parsed it falls back to the whole block, erring towards
      flagging; such a question is already in finding 1 and `suggestedForExclusion` drops it.

    **Corpus effect**: AR flags **51 → 2** (both unparseable, so neither reaches finding 3);
    catch-all **47 → 47** and cross-reference **6 → 6**, i.e. reading options instead of the
    stem lost nothing real. FST-1's finding 3 went 16 → 7.

    **Watch the vocabulary**: FST-1 Q66 writes "Both statements are wrong", not "false" -
    `wrong`/`right` and the plural forms are in `SELF_DESCRIBING` for that reason. A paper
    using some other phrasing would be flagged spuriously again; widen the list, do not
    weaken the rule.

    231 tests, 19 files. New: `optionAdvisor.test.ts`.

19. **The auto-lettered reader had been left behind.** Item 12 gave floating pictures their
    proper treatment in `OptionBlockParser` (typed labels) but `AutoLetteredOptionParser`
    kept the old blanket `atoms.some(floatingGraphic)` refusal - so "Extra MTP-1" Q3, Q4,
    Q27, Q29 were refused while Q12, Q15, Q19, Q20 shuffled, purely because the first group
    lets Word letter its options and the second types the labels. Same picture, different
    code path. It now uses `splitLeadCorePad` exactly as the typed path does.

    **`AutoLetteredOptionSet` changed shape with it.** It used to snapshot raw paragraph
    children and swap them wholesale; it now holds `{lead, core, pad}` per paragraph and
    emits `pPr + own.lead + source.core + own.pad` through `mergeAdjacentRuns`.
    **`signatures` must be `slotSignature(slot.core)`, not the whole paragraph** - an
    anchored picture stays behind, so a signature that counted it would differ after every
    shuffle and `SetVerifier` would report a fault that is not there.

20. **A Word-lettered option can sit anywhere, not just first.** "Extra MTP-1" Q28 types
    `(A)`, `(B)`, `(D)` and lets Word letter the third, so no `(C)` is in the text and the
    reader said `Expected labels A,B,C,D but found "ABD"`. (The user asked whether the
    picture's "A" and "B" caused it - it did not; there is no picture in that question.)

    `letteredFirstOption` became `letteredOptionFor(paragraphs, found, gap)`, and
    `missingSlot` finds the one absent letter in an otherwise in-order run of three. The
    candidate paragraph must be a single-item lettered list **positioned in the gap** -
    after the label before it and before the label after it - which is what tells the
    missing option from any other lettered list in the question. `build` splices the
    zero-width label range at `lettered.slot` instead of unshifting it.

    **Corpus effect**: Extra MTP-1 45 → **50 of 50**, FST-1 182 → 183, Alternating Current
    95 → 96; everything else unchanged, all verifications PASS except ANIMAL KINGDOM Set 02.

    239 tests, 19 files. `letteredOptionIndex` added to the fixture.

21. **Two options on a line pushed apart with spaces, not a tab.** "Alternating Current"
    Q75 writes `…High power loss in transmission␣×14␣(B) Low power loss…`. Both the strict
    and the relaxed pass refuse a label that follows spaces preceded by a letter or digit -
    that shape is also `Both (A) and (B)` - so only the paragraph-initial `(C)` was found
    and the report said `Expected labels A,B,C,D but found "C"`, which names nothing to fix.

    Per the standing rule this is **not** worked around: a run of spaces used as column
    alignment is exactly the input the tool asks authors to standardise. What changed is
    the refusal. `labelsAfterSpaceRun` re-scans for labels rejected *only* by that
    separator (≥ `COLUMN_GAP_SPACES` = 2 spaces, so a glued `1s22s2(C)` is not one), and
    `explainFailure` reports the new `label-after-spaces-not-tab` reason **only when adding
    them accounts for all four options** - full `ABCD`, or a single gap that
    `letteredOptionFor` can fill from a Word-lettered paragraph. Without that guard a
    question whose real fault is elsewhere would be blamed on a stray space run.

    Only Q75 in the whole corpus hits it; every shuffled count is unchanged.

    242 tests, 19 files.

22. **The output folder carries the run's date and time too.**
    `question-sets - 31-08-2026-13-21`, from the same `runStartedAt` that names every file
    inside it, so a folder and its sets plainly describe one run and runs sort by when they
    were made. `OutputFolderResolver.create` takes the `Date`; the old `NN` counter now only
    breaks a tie between two runs started in the same minute (`-02`, `-03`, …), which is
    what keeps the "a run never overwrites an earlier one" guarantee. The folder name does
    **not** carry the paper name - two papers shuffled in the same minute land in
    `… -13-21` and `… -13-21-02`, and the files inside are named by paper.

23. **The report is a PDF, written by hand.** `_generation-report.md` became
    `_generation-report.pdf`. `ReportWriter.build` now returns a `ReportDocument` (headings,
    paragraphs, `facts`, `table`) and `renderReportPdf` typesets it, so *what the report
    says* and *how it looks* are testable apart.

    `src/core/report/Pdf.ts` writes the PDF: three standard Type 1 fonts (nothing embedded),
    uncompressed content streams, a real xref table. Chosen over Electron's `printToPDF`
    because the CLI - which runs every regression - has no Electron, and over a library
    because the project ships two dependencies and a page of text and rules needs neither.
    The output is deterministic, so the same run gives the same bytes.

    Two traps it already fell into, both pinned by tests in `tests/reportPdf.test.ts`:
    - **Measure what you will print.** WinAnsi has no `→`, so `A→C` is spelled `A->C`;
      `textWidth` and `encode` both go through `winAnsi()`, or the mapping column would
      have been measured three characters short and run into the next column.
    - **A cell that fits exactly must not wrap.** A column is sized from its widest cell, so
      that cell measures as exactly the width available - one floating-point step turned
      "Subject" into "Subjec" over "t". Hence `CELL_SLACK`.

    Column widths: `weight` marks the columns carrying prose. Those take any spare width,
    and those give width back when the table is too wide - a subject column keeps its size,
    because "CHEMISTRY" broken over two lines is worse than a taller reason column.

    `tests/support/pdfText.ts` reads the text back out, so the tests still assert on the
    file that was written rather than on the model behind it.

24. **One palette, no picker.** The Colour menu is gone and Warm sand's values live in
    `:root` (with the dark-scheme block). This closes open issue 2 - there is no stored
    choice left to apply after load, so nothing can flash.

    265 tests, 20 files.

25. **A picture anchored across a question boundary pins both questions.** Reported as
    "images not properly placed" for Alternating Current Q42.

    The question's XML in the generated set is byte-identical to the source apart from the
    page-flow flags, and Word reports the same shapes at the same offsets - yet the artwork
    was gone. Word's object model gave the answer: one of the graphs, `Group 29`, is
    anchored at character 11652, which is **question 41's last paragraph**, and drawn 20pt
    below it - straight over question 42's `(A)/(B)` row. A float is drawn *downwards* from
    its anchor, and Word attaches that anchor to whichever paragraph was nearest when the
    picture was dropped. Move the two questions apart and the picture leaves with 41.

    Which question such a picture belongs to is a fact about Word's *layout*, not about the
    file, so it is not guessed. `questionsPinnedByPictures` (`core/parse/BoundaryPictures.ts`)
    finds a floating graphic in a block's **last** paragraph - the only place one can be
    drawn over the *next* question - and holds that question and its successor at their
    original positions. Everything else shuffles around them. A picture anchored anywhere
    earlier is drawn over its own question's paragraphs, which travel with it, so it is
    ignored.

    Wired in as ordinary question exclusions, so the planner needed no new concept. The
    numbers are kept out of `optionsKeptByUser` and out of the "this paper has no question
    N" warning - a number the *tool* pinned must never be reported back as the user's.

    **Dry-run finding 4** was added for it. Findings 1-3 are about a question's *options*;
    this one is about its *position*, which is a separate axis and so a separate finding.

    **Corpus cost**: 2 questions of 100 (Alternating Current), 15 of 200 (FST-1), 2 of 180
    (Group A); every other paper unaffected. All verifications still PASS.

    **How it was proved**, and how to re-check: Word COM exports the page to PDF
    (`ExportAsFixedFormat`), and a scratch script inflates the page content stream and counts
    the drawing operators in the band the artwork occupies. Before the fix that band held
    `m:7 c:11 re:8 l:5`; after it, `m:13 c:60 re:13 l:9` - identical to the source. Counting
    operators beats looking at a thumbnail; the Bézier count alone is unreliable because the
    page watermark falls in different bands depending on where the question sits.

    288 tests, 22 files.

26. **Windows builds, so the app can be handed to the people who type the papers.**

    - `npm run portable` → `release/Question Paper Shuffler <version> (portable).zip`
      (~110 MB). Extract anywhere, run the exe; nothing installed, no admin rights, no
      registry. A `READ ME FIRST.txt` goes in the zip. Zipped with **JSZip**, already a
      dependency, so this build needs no archiver on the machine - which matters, see below.
    - `npm run dist` → `release/Question Paper Shuffler Setup <version>.exe`, an assisted
      per-user NSIS installer with shortcuts and an uninstall entry. `electron-builder.yml`.
    - `npm run icon` draws `build/icon.ico` (`scripts/make-icon.mjs`): three fanned papers,
      the watermark's idea, rendered by writing pixels and deflating a PNG by hand - the
      same "no new dependency for a small job" call as the PDF writer. 7 sizes, 16-256 px.

    **The installer cannot be built on this machine.** electron-builder compresses the
    installer payload with a bundled `7za.exe`, and the endpoint-protection agent here
    quarantines it - the file is *deleted* as it runs and the build dies with `spawn EPERM`.
    Established by elimination: `electron.exe` and NSIS's own `makensis.exe` both run fine
    from the same directories, unsigned, so it is that binary and not the path or the
    signature. Windows Defender's service is not even running (`Get-MpPreference` fails with
    0x800106ba), so it is a third-party agent.

    Two things came out of chasing it, both kept:
    - `toolsets: { nsis: 1.2.1 }` in `electron-builder.yml`. The default NSIS toolset is a
      `.7z` and needed 7za merely to *unpack*; the 1.2.1 bundle is a `.tar.gz`, unpacked
      in-process by Node. That removed one of the two 7za uses.
    - The other use - creating `app.7z` - has no way around it: on Windows
      `targets/archive.ts` always goes through 7za, `useZip` included, and the MSI target
      needs 7za to unpack WiX. So the fix is to give it a 7za that is allowed:
      `ELECTRON_BUILDER_7ZIP_PATH="C:\Program Files\7-Zip\7z.exe" npm run dist` after
      `winget install 7zip.7zip`, or an exclusion from whoever administers the machine.

    **Not done, deliberately**: pre-extracting toolsets into electron-builder's cache and
    hand-writing its `.state` files. It would have worked today and broken on the next
    upgrade - the standing rule's "is this reading a standard form, or guessing past an
    ambiguity?" applies to build tooling too.

    Neither build is signed, so SmartScreen warns on first run. README says what to tell
    recipients. Both builds were smoke-tested: packaged app launches with the right window
    title and icon, and the app runs from a freshly extracted copy of the zip.

27. **A `w:sym` in the Symbol font is now read, not treated as unreadable ink.**
    `src/core/docx/SymbolFont.ts`.

    The refusal in item 4/7 rested on a claim that turned out to be false: *"that character
    carries no readable text."* `<w:sym w:font="Symbol" w:char="F028"/>` states its code in
    the font's own encoding, and Symbol's encoding is published - the same Adobe
    specification that gives `Pdf.ts` its widths. 0x28 is `parenleft`. Reading it is
    deduction, not a guess, so the standing rule never applied to this case.

    `symbolCharacter(font, char)` decodes Symbol and nothing else; `symbolTextOf` in
    `xml.ts` is used by both `visibleText` and `Atoms.atomText`, so every reading of the
    document sees the character. `Atom.symbol` became `Atom.unreadableSymbol` - true only
    for a `w:sym` that decoded to nothing - and the refusal now fires only for those. The
    fix message names picture fonts instead of Insert > Symbol.

    **Do not use the naive low-byte mapping.** `F071` is *not* `q`: Symbol 0x71 is θ. Half
    the codes coincide with ASCII and half do not, and a lower-case letter invented out of a
    physics variable would fabricate an option label. The table is the real encoding, written
    as the runs of consecutive codes that have a character to read; the gaps (0x60,
    0xBD-0xBE, 0xE6-0xF0, 0xF3+) are glyph *pieces* for drawing tall brackets, which are ink
    rather than characters and still read as nothing.

    Corpus effect: MTP-2-XI-2023 **158 → 191** shuffled, all 33 `label-bracket-is-a-symbol`
    refusals gone; every other paper unchanged to the question, all twelve still verify.
    Option text also stopped losing characters everywhere - `at an angle θ` used to read
    `at an angle `.

    `scripts/regress.mjs` was written for this and kept: it copies the corpus to a scratch
    directory, dry-runs each paper and prints shuffled counts and refusal reasons, with
    `--generate` to write two sets each and verify them.

28. **Questions shuffle within a *section*, not within a subject.**
    `PaperParser.splitIntoGroups`, `PaperModel.PaperSection.group` / `.label`.

    Papers divide a subject into `SECTION A (All questions are compulsory)` and
    `SECTION B (Attempt any 10 questions)`, and Biology often adds `PART 1` / `PART 2` above
    those. Those divisions are rules, not decoration: a compulsory question shuffled into
    the attempt-any-10 section changes what the candidate is asked to do. The tool was
    shuffling across them.

    A second defect fell out of the same cause. A question block runs from one numbered
    paragraph to the next, so the `SECTION B` heading was *inside the last question of
    section A* and travelled with it when that question moved. It also read as a
    continuation of that question's last option - which is why MTP-2-XI-2023 went from 191
    to **193** options shuffled when this landed: Q35 and Q85 stopped being refused with
    `option-spans-paragraphs`.

    `PaperSection` is now the *division*, not the subject, and carries `subject`, `group`
    (`PART 2 SECTION A`) and `label` (`BIOLOGY - PART 2 SECTION A`). The planner, builder
    and verifier needed no change: they were already generic over sections. `headingNode` -
    which drives "each subject starts a new page" - is set only on the division that opens
    its subject, so `SECTION B` does not gain a page break.

    A heading with no questions under it (`BIOLOGY PART 1` sitting directly above
    `SECTION - A`) does **not** open a run of its own; it joins the heading below it. So
    Biology comes out as four runs, not six with two empties.

    The summary types were renamed with it: `SubjectSummary`/`DryRunSubject` are now
    `GroupSummary`/`DryRunGroup` with a `group` field, and `PaperSummary.subjects` /
    `DryRunReport.subjects` are `.groups`. "Subject" had stopped being true - one subject
    contributes four of these.

    Corpus: three papers are divided this way. MTP-2-XI-2023 (8 sections), FST-1-Rep-2024
    (4 subjects x 2), Extra MTP-1 (2, with no subject heading at all - labelled plain
    `SECTION A` / `SECTION B`, since `ALL` is a stand-in rather than a name). Every paper
    still verifies, and a generated set re-parses into exactly the same eight sections.

29. **Any mixture of Word-lettered and typed option labels is read.**
    `OptionBlockParser.letteredOptionsFor` + the `merged` helper.

    The old reader handled two shapes: all four lettered by Word (the auto-lettered reader),
    or exactly *one* lettered among three typed. Real papers do neither in several places -
    `(A) (B) (C)` lettered with the `(D)` typed (4 questions), two and two (1), and one
    lettered `(A)` sharing a line with a typed `(B)`.

    Reading is now a **merge, not a search**: typed labels and lettered paragraphs are laid
    out in document order and must come to exactly four, each typed label at the position its
    own letter names. That single rule replaces the old per-gap search and covers every
    mixture. It is also the safeguard - three lettered statements plus three typed labels
    merge to six, and are refused.

    Two things were learned the hard way, both now pinned by tests:
    - **A paragraph can be two options.** `(A) 1 amp <tab> (B) 1.5 amp` is one paragraph that
      is a list item *and* carries a typed label. An early version excluded such paragraphs
      as candidates and cost *Alternating Current* 76 questions. A lettered entry sorts
      before any typed label in the same paragraph, because Word draws its number at the
      start of the line.
    - **Narrow before wide.** Q99 of the same paper has two lettered lists - `(a) (b)` for
      its two situations and `(A)` for option A. Accepting both gives six entries and fails.
      `FORMATS_OF_SCHEME` is therefore a *ladder*: same case first, then the family; within
      each, bracketed `(%1)` before plain `%1.`. The first tier that comes to four wins.

    Corpus: **+6** questions, no losses. ANIMAL KINGDOM 96->97, PCB Group B 174->175,
    MTP-2-XI-2023 193->197. All twelve papers still verify.

    Still refused, correctly: FST-1 Q170 letters three options `(A) (B) (C)` with Word and
    types the fourth as `(4)`. Letters and digits are not one naming scheme, and
    `FORMATS_OF_SCHEME` keeps a decimal label from matching a letter list.

30. **An option split across two paragraphs stays refused - with a message that names the
    key to press.** MTP-2-XI-2023 Q170 and Q195.

    Reading it is unambiguous; *writing* it is not. `OptionShuffleApplier` buckets atoms by
    paragraph and rewrites each paragraph in place, so an option's content has to have a line
    to move into. Moving two paragraphs' worth into a one-paragraph slot would either join
    the lines or leave an empty paragraph behind - a layout change, not just a re-ordering,
    and the invariant this project is built on is that shuffling changes nothing but order.

    The input fix is one keystroke and keeps the paper looking identical: **Shift+Enter**
    (a `w:br`) instead of Enter. A `w:br` is an atom inside the paragraph, so it travels with
    the content and the option still prints on two lines. `mixedLabels.test.ts` proves both
    halves - the refusal, and that the same question shuffles once the break is a `w:br`.

    **Not done, and this is the open question if it comes back:** supporting a *uniform*
    multi-paragraph shape, where all four options have the same paragraph count (Q195 is
    label+continuation four times over). Paragraph *k* of the source would map to paragraph
    *k* of the destination, and layout would be preserved exactly. It is buildable and worth
    exactly **one question** in the whole corpus; the ragged shape (Q170) would still be
    refused, because there is no destination paragraph to map to.

31. **Every refusal now says what is wrong *and* what to change, and they are grouped.**
    `shared/skipReasons.ts`, `SkippedOptionShuffle.fix`, `SkippedOptionGroup`.

    Finding 1 used to print `Q125 [BIOLOGY] options-not-found: No paragraph follows the
    question stem.` - a slug that means nothing to whoever types the paper, and a sentence
    that does not say what to do. The layout notes beside it already had `detail` + `fix`
    and were grouped by problem; the skips now work the same way, so the two findings answer
    the same question in the same shape.

    - `SkipReason` gained a plain-language `SKIP_REASON_LABEL`. **The slug is never shown**,
      and a test asserts it never reaches the report PDF.
    - Every one of the 15 refusal sites was rewritten as cause + remedy. The remedy names a
      key or a menu the author can find: *Shift+Enter*, *Convert to Text*, *Layout Options >
      In line with text*.
    - `groupSkippedOptions` gathers them by reason, most-affected first, with each question's
      own detail underneath - so an odd one out inside a group of 24 is still findable.

    Two things fell out of writing the tests, both real defects:
    - A question whose options are *entirely* inside a table returned "nothing follows the
      question stem", because `paragraphs.length === 0` returned before the table check.
      Telling an author nothing follows a stem that plainly has four options would send them
      hunting. `optionsInsideTable` is now checked in both places.
    - One `fix` said 'write those "A." rather than "(A)"' - in a paper whose options are
      `(1)`-`(4)`. That breaks the rule that user-facing text uses the paper's own option
      names, and `dryRun.test.ts` caught it. The example was dropped; `expected` in the same
      sentence is already rendered in the paper's scheme.

32. **The dry run now accounts for every question, and the two exclusion lists no longer
    survive a change of paper.** `shared/accounting.ts`, `PaperSummary.questionNumbers`,
    `DryRunReport.questionAccounting` / `.optionAccounting`, `scripts/audit-report.mjs`.

    Reported as two separate complaints about one report, on the ANIMAL KINGDOM paper:

    - *"3 question(s) whose options cannot be shuffled with certainty"* beside a headline of
      *"options would be shuffled for 77 questions"* out of 100. Both numbers were correct.
      The other 20 were the user's own exclusion list, mentioned only inside finding 3, so
      the report read as a contradiction.
    - *"This paper has no question 102, 112, 117, 124, 158, 193"* on a 100-question paper.

    One root cause. Picking a new paper cleared the summary, dry-run and results panels but
    **not** the two exclusion fields, and the "Add these N to keep the option order" button
    merges rather than replaces. The field held the union of three papers' suggestions -
    MTP-2-XI-2023 (13), ANIMAL KINGDOM (7), Alternating Current (7) - which is 20 numbers in
    1..100 and exactly those 6 above it. 97 - 20 = 77. Reproduced number for number before
    changing anything.

    - `browseButton` empties both lists when the chosen path differs from the current one,
      and says so. A question number only means something in the paper it came from.
    - The warning now says *why* a number is foreign ("check the list belongs to this paper,
      and clear anything left over from another one"), because the numbers alone do not.
    - The generation report gained the same note: a completed run made with a stale list left
      no record of it anywhere.
    - `accountFor` splits the paper into three disjoint outcomes - shuffled, kept by the tool,
      kept because you asked - which always sum to the question count, and words the sum
      once. The tool's list wins where the two overlap: taking such a question off the list
      would change nothing, so calling it the user's doing sends them to the wrong place.
    - `PaperSummary.questionNumbers` is now the single source of which numbers exist. The
      section table, the totals and every "not in this paper" check read from it.

    The sum is worded in the **main process** and shipped as `ShuffleAccounting.summary`,
    because `renderer.ts` must have zero runtime imports (item 8). `scripts/ui-harness.mjs`
    inlines the compiled `accounting.js` with its `export` keywords stripped rather than
    re-implementing the arithmetic - a stub that computes it differently is a fourth answer.

    `npm run audit` checks ~40 identities per report across the corpus under four settings
    (44 reports). It was sabotage-tested: dropping one term from `accountFor` makes it report
    44 contradictions. It also found one **wrong count in `regress.mjs`** - the `pinned`
    column summed group memberships, and FST-1 has a question held by two different pictures,
    so it read 16 for 15 questions. Now distinct; that is the only corpus number that moved.

## 5 · Decisions worth not relitigating

| Decision | Why |
| --- | --- |
| Accept `1.` / `1)` in the key, refuse `57,` | The first is one unambiguous reading and the form authors type; the second needs a guess about which characters are decoration |
| Report key gaps before resolving numbering | The key's entry count *defines* the question count, so a short key produces a confident, wrong complaint about the numbering |
| `splitLeadCorePad` kept | Justified by measurement, not taste: without it four clean papers lost characters |
| Loose-signature verifier fallback **reverted** | It made a defective document report PASSED — the exact failure mode verification exists to prevent |
| Symbol check keyed on the *missing bracket*, not on the symbol | The missing bracket is the actual signal; keying on the symbol punishes legitimate `60Ω` options |
| A refusal owes a `fix`, not just a `detail` | A question the tool skips is only useful if the person who typed the paper can act on it; the fix names the key or the menu |
| Skips are grouped by problem, like the layout notes | One paper has 24 questions with the same table problem; printing the same paragraph 24 times buries the one that failed for its own reason |
| The `SkipReason` slug is never shown to the user | It is a name for the code to switch on. A test asserts no slug reaches the report |
| Mixed labels are read by merging, never by searching for the missing one | One rule covers every mixture, and "does it come to exactly four, in order?" is its own safeguard against a lettered statement list |
| A paragraph may hold both a Word-lettered option and a typed one | `(A) 1 amp <tab> (B) 1.5 amp` is one paragraph; excluding such paragraphs cost 76 questions in one paper |
| An option split by a paragraph break is refused, not joined | Joining it or leaving an empty paragraph is a layout change; Shift+Enter is a one-keystroke input fix that prints identically |
| A section heading is read from its whole paragraph, never from a word inside one | "cross-section" appears in a real question stem; requiring the paragraph to be *only* `SECTION B (...)` is what keeps prose out |
| A part or section heading with no questions under it joins the one below | `BIOLOGY PART 1` sits directly above `SECTION - A`; opening a run for it would put two empty rows in every report |
| Symbol-font characters are decoded; picture fonts are not | Symbol has a published encoding, so reading it is deduction. Wingdings numbers drawings, so there is genuinely nothing to read and the refusal stands |
| The Symbol table is the real encoding, never the low byte | `F071` is θ, not `q`. A naive mapping would invent a lower-case option label out of a physics variable |
| Option schemes tried letters → digits → roman | Order of decreasing certainty that a run of labels is the *options*; an `(i)…(iv)` item list must never outrank real answers |
| A key table must use one answer scheme throughout | It is what separates an answer key from any other grid of numbers, now that digits are valid answers |
| Slots stay named A–D internally whatever the page writes | One canonical vocabulary for permutations and the key; labels are never moved, so the page keeps its own tokens regardless |
| The two option readers must gain capabilities together | Typed-label and Word-lettered layouts are the same paper feature; fixing one left four questions refused for no reason the user could see |
| An auto-lettered signature covers the core only, not the paragraph | An anchored picture stays behind; counting it would make the verifier fail every shuffled question |
| The advisor flags position dependence, not question type | An Assertion-Reason option that states its own meaning moves safely; flagging all 51 in the corpus was noise that trained the user to ignore finding 3 |
| The advisor reads options, never the stem | Only an *option* can depend on where it sits; a stem saying "all of the above" is about the question |
| Floating-picture options refused, not paired by position | The picture-to-option pairing differs question to question and needs a layout pass; a wrong pairing puts the key on the wrong picture |
| A blank paragraph ends the option list | Matches the existing subject rule, and keeps a genuine continuation (no blank between) refused |
| Three numbered dry-run findings, shown even at zero | They answer three different questions; a predictable report is read, a variable one is skimmed |
| A floating picture contributes no text and never moves | Both follow from Word placing it from the page; together they turn three unrelated-looking refusals into shuffled questions |
| Refuse only a picture *inside* a core, not one at the edges | If the pictures are the options, every core comes out empty and the existing check refuses anyway - the rule self-polices |
| The answer key breaks ties between label schemes | A match-the-columns list `(a)-(d)` above real `(1)-(4)` options otherwise wins on pass order alone |
| One timestamp per run, not per file | A batch must stay together in a folder listing; per-file stamping splits it whenever a run crosses a minute |
| `-` between hours and minutes, not `:` | Windows forbids `:` in a file name - a literal clock reading could not be saved |
| A run of spaces before a label is refused, not read | It is also the shape of "Both (A) and (B)"; the fix is a keystroke in Word, and the refusal now names the labels |
| The PDF is written by hand, not by Electron or a library | Every regression run goes through the CLI, which has no Electron; a page of text and rules needs no third dependency, and the bytes stay deterministic |
| `ReportDocument` between the writer and the page | What the report says is testable without a PDF, and the typesetting without a generation run |
| Measuring and drawing share one WinAnsi pass | `→` prints as `->`; measuring the original string would reserve two characters too few and overrun the column |
| Only weighted (prose) columns give up width | A subject column that wraps reads worse than a taller reason column, and only prose loses nothing by taking another line |
| A float in a block's last paragraph pins that question and the next | It is drawn downwards over the following question, and which question it belongs to is a layout fact Word alone has; holding the pair costs 19 questions in the whole corpus |
| Dry-run finding 4 kept separate from 1-3 | Findings 1-3 are about a question's options; this one is about its position |
| The portable zip is built with JSZip, not an archiver | It is already a dependency, and the machine's endpoint agent will not run `7za.exe` |
| electron-builder's toolset cache is left alone | Pre-extracting and hand-writing its `.state` files would work today and break on the next upgrade |
| One palette, written into `:root` | The picker was a setting nobody needed to change; removing it also removed the startup flash it caused |
| Watermark on `body::before`, not `body` | Lets the theme control its opacity and tint independently of the wash |

---

## 6 · Current state

- **326 tests pass**, 26 files. `npx vitest run`.
- Renderer type-checks; UI harness builds; Electron launches clean.

### Paper corpus (`C:\ps\q-paper`)

| Paper | Q | Options shuffled | Status |
| --- | --- | --- | --- |
| ANIMAL KINGDOM (TEST-1) | 100 | 97 | Works — **but see open issue 1** |
| Extra MTP-1-XI-2023 | 50 | 50 | Works — every question |
| MTP-2-PCB-XI-2027 Group A | 180 | 175 | Works |
| MTP-2-PCB-XI-2027 Group B | 180 | 175 | Works |
| MTP-2-PCB-XI-2027 Group C | 180 | 172 | Works |
| MTP-2-PCB-XI-2027 Group D | 180 | 175 | Works |
| MTP-2-XI-2023 | 200 | 197 | Works — 8 sections; items 27, 28 and 29. Only Q125, Q170, Q195 left |
| Nano MTP 2 Physics | 12 | 11 | Works |
| Nano MTP Chemistry-1 | 12 | 12 | Works |
| Alternating Current XII-2024 | 100 | 97 | Works — the `57,` box and Q75’s space-run separator have since been retyped in Word |
| FST-1-Rep-2024 | 200 | 183 | Works — unblocked by the format deduction below |
| CURRENT ELECTRICITY (Wheatstone) | — | — | Blocked: question numbers typed by hand |

---

## 7 · Open issues

### 1. Shuffling can destroy the boundary that made a question readable — **highest priority**

`ANIMAL KINGDOM` fails verification on *some seeds*. Reproduce:

```bash
npm run cli -- --file "<copy of ANIMAL KINGDOM…docx>" --shuffle-questions --shuffle-options --sets 2 --seed regress
```

Set 02 reports `1 question(s) could not be matched by content`.

**Cause.** The question is laid out `(C) →(i), (ii) and (iii) (D) →all of the above` —
no tab before `(D)`. It is only readable because `isLabelPosition`'s relaxed pass accepts
a label after a space when the preceding character is non-alphanumeric — here the `)` of
`(iii)`. After the shuffle puts `all of the above` in slot C, the preceding character is
`e`, the relaxed pass rejects `(D)`, and the file no longer reads as four options.

**Confirmed pre-existing** — reproduced on the committed code with the changes stashed.
It is not caused by anything done this session.

**Options.**

- **(a) Recommended — re-parse each shuffled block and roll back.** After
  `OptionShuffleApplier.apply`, re-run the option parser on that block; if it no longer
  yields `A,B,C,D`, restore the original order and report the question as skipped. This is
  the existing whole-file verification applied per question, so it is a safety net rather
  than a special case, and it generalises to causes nobody has hit yet. Costs a re-parse
  per shuffled question.
- **(b) Refuse up front** any question whose labels were found only by a relaxed pass
  *and* whose separator is content-dependent. Cheaper, but blunter, and it would skip
  questions that shuffle perfectly well.
- **(c) Treat as input-only.** The tool already reports `label-not-after-tab` with the fix
  *"Press Tab before each option label."* Under the standing rule this is defensible —
  but it leaves the tool writing a file it cannot re-read, which is worse than refusing.

**This needs the user's decision** before implementing, because (a) changes coverage
numbers and (c) leaves a known bad output in place. My recommendation is (a).

### 2. Palette flash at startup — **closed**

The colour picker is gone (item 24) and Warm sand is written straight into `:root`, so
there is no stored choice to apply after the page loads and nothing to flash.

### 3. One paper unsupported by design

`CURRENT ELECTRICITY` has hand-typed question numbers and is correctly refused with a clear
message. Decide whether to support it, or document it as out of scope permanently.

(`FST-1-Rep-2024` used to sit here as "no key table". It was never missing a key: its key
writes answers as digits `1`–`4` and its options are labelled `(1)`–`(4)`. Both are now
deduced, and the paper works.)

---

## 8 · Exact next steps

1. ~~Retype the `57,` box in *Alternating Current*.~~ Done in Word; that paper now works.
2. **Get a decision on open issue 1**, then implement — most likely (a), a post-apply
   re-parse with rollback in `SetBuilder`/`OptionShuffleApplier`, plus a new `SkipReason`
   such as `'shuffle-would-break-labels'`, a fixture test, and a full corpus re-run.
3. **Re-run the whole corpus** after any parser change and compare *options shuffled* per
   paper against the table in §6. A silent drop means a new false-positive refusal; a
   silent rise means a check stopped firing.
4. Optional: the table-grid option reader that was offered and declined earlier —
   options laid out inside a table are currently skipped.

### How to re-run the corpus

`scripts/regress.mjs` does it. It copies every paper to a scratch directory first — the CLI
writes its output folder next to the source file, and the corpus folder must not be
littered — then prints one row per paper: questions, options shuffled, questions pinned by a
picture, and a breakdown of the refusal reasons.

```bash
npm run build:main && node scripts/regress.mjs
```

Add `--generate` to write two sets per paper and verify them (slower, and the only way to
catch a shuffle that destroys the boundary it was read from). `--dir` points it at another
folder, `--seed` changes the permutations.

Compare the *options shuffled* column against the table in §6. A silent drop means a new
false-positive refusal; a silent rise means a check stopped firing.

---

## 9 · Traps

- Do **not** append to `w:pPr` / `w:trPr`; splice in schema order or Word rejects the file.
- Do **not** compare `.docx` hashes to prove determinism; compare `word/document.xml`.
- `OptionSet.signatures` is a getter over **parse-time** atoms. After `apply`, run merging
  can empty an atom's node, so signatures read *after* apply are misleading. Assert on the
  document (paragraph text, or text with symbols rendered), not on the model.
- A `w:sym` from a picture font contributes `''` to text but is **not** blank. Any new
  whitespace logic must keep it inside `coreAtoms` or symbols get left behind.
- Symbol's codes are not ASCII. `F071` is θ, not `q`. Never map a `w:char` by its low byte.
- The relaxed label pass reads a boundary from surrounding punctuation. That boundary is
  not stable across a shuffle — this is exactly open issue 1.
- Never write generated sets into `C:\ps\q-paper`; copy the paper out first.
