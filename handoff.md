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
npx vitest run       # 160 tests, 13 files
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

## 5 · Decisions worth not relitigating

| Decision | Why |
| --- | --- |
| Accept `1.` / `1)` in the key, refuse `57,` | The first is one unambiguous reading and the form authors type; the second needs a guess about which characters are decoration |
| Report key gaps before resolving numbering | The key's entry count *defines* the question count, so a short key produces a confident, wrong complaint about the numbering |
| `splitLeadCorePad` kept | Justified by measurement, not taste: without it four clean papers lost characters |
| Loose-signature verifier fallback **reverted** | It made a defective document report PASSED — the exact failure mode verification exists to prevent |
| Symbol check keyed on the *missing bracket*, not on the symbol | The missing bracket is the actual signal; keying on the symbol punishes legitimate `60Ω` options |
| Palettes as CSS blocks + one attribute | CSP allows no inline style; also keeps all colour in one file |
| Watermark on `body::before`, not `body` | Lets the theme control its opacity and tint independently of the wash |

---

## 6 · Current state

- **160 tests pass**, 13 files. `npx vitest run`.
- Working tree clean at `5719b9d`.
- Renderer type-checks; UI harness builds; Electron launches clean.

### Paper corpus (`C:\ps\q-paper`)

| Paper | Q | Options shuffled | Status |
| --- | --- | --- | --- |
| ANIMAL KINGDOM (TEST-1) | 100 | 95 | Works — **but see open issue 1** |
| Extra MTP-1-XI-2023 | 50 | 40 | Works |
| MTP-2-PCB-XI-2027 Group A | 180 | 174 | Works |
| MTP-2-PCB-XI-2027 Group B | 180 | 172 | Works |
| MTP-2-PCB-XI-2027 Group C | 180 | 171 | Works |
| MTP-2-PCB-XI-2027 Group D | 180 | 174 | Works |
| MTP-2-XI-2023 | 200 | 155 | Works (32 genuine symbol-bracket questions skipped) |
| Nano MTP 2 Physics | 12 | 9 | Works |
| Nano MTP Chemistry-1 | 12 | 12 | Works |
| Alternating Current XII-2024 | 100 | 93 | **One cell edit away** — retype `57,` as `57.` |
| FST-1-Rep-2024 | — | — | Blocked: no answer key table in the document |
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

### 3. Two papers unsupported by design

`FST-1-Rep-2024` has no key table; `CURRENT ELECTRICITY` has hand-typed numbers. Both are
correctly refused with clear messages. Decide whether to support either, or document them
as out of scope permanently.

---

## 8 · Exact next steps

1. **Tell the user to retype the `57,` box in *Alternating Current*** as `57.`, in Word,
   then re-run. Nothing to code. (Note: they had that file open in Word — a `~$…` lock
   file appeared in `C:\ps\q-paper` mid-session.)
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
