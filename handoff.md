# Handoff — Question Paper Shuffler

Written 2026-08-30. Covers the project brief, the architecture, every decision taken so
far and the exact next steps. Read this plus `README.md` before changing anything.

---

## 1 · Project brief

An **Electron + TypeScript desktop app** that takes one NEET-style Word question paper
(`.docx`) and produces *N* shuffled sets, each with its own correct answer key.

- Questions are re-ordered **only within their own subject**.
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
npx vitest run       # 242 tests, 19 files
npm run cli -- --file "paper.docx" --shuffle-questions --shuffle-options --dry-run
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
PaperParser        body nodes -> subjects -> QuestionBlocks (+ AnswerKeyTable)
      ↓
OptionSetParser    per question: typed-label reader, else auto-lettered reader
      ↓
ShufflePlanner     seeded permutations (questions within subject; options within question)
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
| `core/verify/SetVerifier.ts` | re-parses the written file; the safety net |
| `shared/answerStyle.ts` | the only boundary between internal A-D slots and the paper's own option names |
| `main/` | Electron main, preload bridge, IPC channels |
| `renderer/` | UI. **`renderer.ts` must have zero runtime imports** |
| `cli/cli.ts` | headless driver, used for all regression runs |

### Key data structures

- **`Atom`** — one inline element wrapped in its own run carrying cloned `rPr`. Fields:
  `node`, `paragraphIndex`, `text`, `blank`, `floatingGraphic`, `symbol`. A `w:sym` has
  `text: ''` but `blank: false`, so it is never trimmed away as whitespace.
- **`OptionSlot`** — `labelAtoms` / `leadAtoms` / `coreAtoms` / `padAtoms`. **Only
  `coreAtoms` move.** Label, the tabs after it, and the trailing padding belong to the
  *position*, which is what keeps columns aligned and separators in place.
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

## 5 · Decisions worth not relitigating

| Decision | Why |
| --- | --- |
| Accept `1.` / `1)` in the key, refuse `57,` | The first is one unambiguous reading and the form authors type; the second needs a guess about which characters are decoration |
| Report key gaps before resolving numbering | The key's entry count *defines* the question count, so a short key produces a confident, wrong complaint about the numbering |
| `splitLeadCorePad` kept | Justified by measurement, not taste: without it four clean papers lost characters |
| Loose-signature verifier fallback **reverted** | It made a defective document report PASSED — the exact failure mode verification exists to prevent |
| Symbol check keyed on the *missing bracket*, not on the symbol | The missing bracket is the actual signal; keying on the symbol punishes legitimate `60Ω` options |
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
| Palettes as CSS blocks + one attribute | CSP allows no inline style; also keeps all colour in one file |
| Watermark on `body::before`, not `body` | Lets the theme control its opacity and tint independently of the wash |

---

## 6 · Current state

- **242 tests pass**, 19 files. `npx vitest run`.
- Renderer type-checks; UI harness builds; Electron launches clean.

### Paper corpus (`C:\ps\q-paper`)

| Paper | Q | Options shuffled | Status |
| --- | --- | --- | --- |
| ANIMAL KINGDOM (TEST-1) | 100 | 95 | Works — **but see open issue 1** |
| Extra MTP-1-XI-2023 | 50 | 50 | Works — every question |
| MTP-2-PCB-XI-2027 Group A | 180 | 175 | Works |
| MTP-2-PCB-XI-2027 Group B | 180 | 174 | Works |
| MTP-2-PCB-XI-2027 Group C | 180 | 172 | Works |
| MTP-2-PCB-XI-2027 Group D | 180 | 175 | Works |
| MTP-2-XI-2023 | 200 | 158 | Works (symbol-bracket questions skipped) |
| Nano MTP 2 Physics | 12 | 11 | Works |
| Nano MTP Chemistry-1 | 12 | 12 | Works |
| Alternating Current XII-2024 | 100 | 96 | Works — the `57,` box has since been retyped in Word |
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

### 2. Palette flash at startup

CSP forbids an inline script, so the palette is applied by the deferred module script; a
non-default choice can show one frame of Periwinkle. Fix if wanted: have the main process
substitute the saved palette into the `<html>` tag before loading the page. Cosmetic.

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
4. Optional: palette flash (issue 2), and the table-grid option reader that was offered
   and declined earlier — options laid out inside a table are currently skipped.

### How to re-run the corpus

There is no committed script for this; it was done inline. Copy each paper to a temp
directory first — the CLI writes `question-sets-NN` next to the source file, and you must
not litter the user's folder.

```bash
for f in "C:/ps/q-paper/"*.docx; do
  b=$(basename "$f" .docx); cp "$f" "./$b.docx"
  node dist/cli/cli.js --file "./$b.docx" --shuffle-questions --shuffle-options \
    --sets 2 --seed regress 2>&1 | grep -E "verification|^Error:"
done
```

Consider promoting this to `scripts/regress.mjs` — it has been retyped several times.

---

## 9 · Traps

- Do **not** append to `w:pPr` / `w:trPr`; splice in schema order or Word rejects the file.
- Do **not** compare `.docx` hashes to prove determinism; compare `word/document.xml`.
- `OptionSet.signatures` is a getter over **parse-time** atoms. After `apply`, run merging
  can empty an atom's node, so signatures read *after* apply are misleading. Assert on the
  document (paragraph text, or text with symbols rendered), not on the model.
- A `w:sym` contributes `''` to text but is **not** blank. Any new whitespace logic must
  keep it inside `coreAtoms` or symbols get left behind.
- The relaxed label pass reads a boundary from surrounding punctuation. That boundary is
  not stable across a shuffle — this is exactly open issue 1.
- Never write generated sets into `C:\ps\q-paper`; copy the paper out first.
