import type { PinnedQuestionGroup } from '../../shared/types';
import { NS } from '../docx/xml';
import { hasFloatingGraphic } from '../options/Atoms';
import type { ParsedPaper } from './PaperModel';

/**
 * Finds questions that must keep their position because a floating picture is anchored at
 * the boundary between them.
 *
 * A floating picture is placed from the paragraph it is anchored to and drawn *downwards*
 * from there, as far as its height takes it - which can be well past the end of that
 * paragraph, and so past the end of the question. Word attaches the anchor to whichever
 * paragraph was nearest when the picture was dropped, so a picture illustrating question 42
 * is quite often anchored in the last paragraph of question 41.
 *
 * While the two questions stay next to each other the page looks right. Move either of them
 * and the picture travels with its anchor: question 42 loses its artwork, and question 41
 * carries away a picture that means nothing beside it.
 *
 * Which question such a picture belongs to cannot be read from the file - that is a fact
 * about where Word *lays it out*, and needs Word. So the pair is kept together instead:
 * both questions keep their original positions, everything else shuffles around them, and
 * the report names the picture to re-anchor to lift the restriction.
 *
 * Only the **last** paragraph of a question is examined. A picture anchored anywhere earlier
 * is drawn over that question's own paragraphs, which travel with it.
 */
export function questionsPinnedByPictures(paper: ParsedPaper): PinnedQuestionGroup[] {
  const groups: PinnedQuestionGroup[] = [];

  for (const section of paper.sections) {
    section.blocks.forEach((block, index) => {
      const paragraphs = block.nodes.filter(
        (node) => node.namespaceURI === NS.w && node.localName === 'p',
      );
      const last = paragraphs[paragraphs.length - 1];
      if (!last || !hasFloatingGraphic(last)) return;

      const next = section.blocks[index + 1]?.printedNumber;
      const numbers = next === undefined ? [block.printedNumber] : [block.printedNumber, next];
      groups.push({
        questionNumbers: numbers,
        subject: section.subject,
        detail:
          `A floating picture is anchored in the last paragraph of question ${block.printedNumber}, ` +
          `so it is drawn over whatever follows it` +
          `${next === undefined ? '' : ` - question ${next}`}. Moving either question away from ` +
          'the other would take the picture with it, leaving one question without its artwork.',
        fix:
          `In Word, click the picture at the end of question ${block.printedNumber} and drag its ` +
          'anchor marker into the question the picture illustrates - or set its Layout Options to ' +
          '"In line with text", which anchors it exactly where it sits.',
      });
    });
  }

  return groups;
}

/** The question numbers `questionsPinnedByPictures` holds in place, without repeats. */
export function pinnedQuestionNumbers(groups: readonly PinnedQuestionGroup[]): number[] {
  return [...new Set(groups.flatMap((group) => group.questionNumbers))].sort((a, b) => a - b);
}
