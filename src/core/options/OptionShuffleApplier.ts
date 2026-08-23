import type { Node } from '../docx/dom';
import { NS, firstChild, replaceChildren } from '../docx/xml';
import type { Atom } from './Atoms';
import type { OptionBlock } from './OptionBlockParser';
import { mergeAdjacentRuns } from './RunMerger';

/**
 * Rewrites the option paragraphs so that slot *i* shows the content that used to be in
 * slot `permutation[i]`.
 *
 * Labels, tabs and paragraph boundaries stay exactly where they were, so the printed
 * layout (two options per line, four per line, one per line...) is preserved; only the
 * answer text moves between slots.
 */
export class OptionShuffleApplier {
  apply(block: OptionBlock, permutation: readonly number[]): void {
    if (permutation.length !== block.slots.length) {
      throw new Error(
        `Permutation of length ${permutation.length} cannot be applied to ${block.slots.length} options`,
      );
    }

    const buckets = new Map<number, Atom[]>();
    for (const entry of block.paragraphs) buckets.set(entry.paragraphIndex, []);
    const push = (atom: Atom, paragraphIndex: number): void => {
      buckets.get(paragraphIndex)?.push(atom);
    };

    for (const atom of block.prefixAtoms) push(atom, atom.paragraphIndex);

    block.slots.forEach((slot, index) => {
      // Label, the tabs after it and the trailing padding all belong to the *position*;
      // only the answer text comes from the source slot.
      for (const atom of slot.labelAtoms) push(atom, slot.paragraphIndex);
      for (const atom of slot.leadAtoms) push(atom, slot.paragraphIndex);
      const source = block.slots[permutation[index]!];
      if (!source) throw new Error(`Permutation refers to option ${permutation[index]}, which does not exist`);
      for (const atom of source.coreAtoms) push(atom, slot.paragraphIndex);
      for (const atom of slot.padAtoms) push(atom, atom.paragraphIndex);
    });

    for (const entry of block.paragraphs) {
      const pPr = firstChild(entry.paragraph, NS.w, 'pPr');
      const nodes: Node[] = [];
      if (pPr) nodes.push(pPr);
      for (const atom of buckets.get(entry.paragraphIndex) ?? []) nodes.push(atom.node);
      replaceChildren(entry.paragraph, mergeAdjacentRuns(nodes));
    }
  }
}
