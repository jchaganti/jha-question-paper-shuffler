import type { GenerationRequest } from '../../shared/types';
import { SeededRandom, shuffleInPlace, type IRandom } from './Random';

/** A plan is a pure description of one set: what moves where. */
export interface SetPlan {
  readonly setNumber: number;
  readonly seed: string;
  /** Per subject: `order[newPosition] = originalPosition`. */
  readonly sectionOrders: readonly (readonly number[])[];
  /** Printed question number -> `permutation[newSlot] = originalSlot`. */
  readonly optionPermutations: ReadonlyMap<number, readonly number[]>;
}

export interface PlannerInput {
  /** Printed question numbers per subject, in original document order. */
  readonly sections: readonly (readonly number[])[];
  /** Questions whose options can actually be shuffled (parsed successfully). */
  readonly shufflableOptionQuestions: readonly number[];
  readonly request: GenerationRequest;
  /**
   * Seed for the whole run. When omitted, the request's seed is used, or a fresh random
   * one is invented. Passing it in lets the caller record exactly which seed was used.
   */
  readonly baseSeed?: string;
}

const IDENTITY_4 = [0, 1, 2, 3];

/** The seed a run will actually use: the one the user typed, or a fresh random one. */
export function resolveSeed(requested: string | undefined): string {
  const trimmed = requested?.trim();
  return trimmed ? trimmed : SeededRandom.newSeed();
}

/**
 * Builds one plan per set.
 *
 * Rules implemented here:
 *  - questions only ever move inside their own subject;
 *  - a question listed in `questionExclusions` keeps its exact position;
 *  - a question listed in `optionExclusions` keeps its option order;
 *  - a plan never repeats the original paper or an earlier set (best effort, 50 tries).
 */
export class ShufflePlanner {
  plan(input: PlannerInput): SetPlan[] {
    const { request } = input;
    const baseSeed = input.baseSeed ?? resolveSeed(request.seed);
    const seen = new Set<string>([identitySignature(input)]);
    const plans: SetPlan[] = [];

    for (let setNumber = 1; setNumber <= request.setCount; setNumber++) {
      let plan: SetPlan | undefined;
      for (let attempt = 0; attempt < 50; attempt++) {
        const seed = `${baseSeed}#set${setNumber}${attempt === 0 ? '' : `#retry${attempt}`}`;
        const candidate = this.planOne(input, setNumber, seed);
        const signature = signatureOf(candidate);
        if (!seen.has(signature)) {
          seen.add(signature);
          plan = candidate;
          break;
        }
        plan = candidate; // keep the last attempt if uniqueness is impossible
      }
      plans.push(plan!);
    }

    return plans;
  }

  private planOne(input: PlannerInput, setNumber: number, seed: string): SetPlan {
    const random = new SeededRandom(seed);
    const { request } = input;

    const questionExclusions = new Set(request.questionExclusions);
    const optionExclusions = new Set(request.optionExclusions);

    const sectionOrders = input.sections.map((questionNumbers) =>
      request.shuffleQuestions
        ? orderWithFixedPositions(questionNumbers, questionExclusions, random)
        : questionNumbers.map((_, index) => index),
    );

    const optionPermutations = new Map<number, readonly number[]>();
    if (request.shuffleOptions) {
      for (const questionNumber of input.shufflableOptionQuestions) {
        if (optionExclusions.has(questionNumber)) continue;
        optionPermutations.set(questionNumber, nonIdentityPermutation(random));
      }
    }

    return { setNumber, seed, sectionOrders, optionPermutations };
  }
}

/** Shuffles only the positions that are not pinned by the exclusion list. */
export function orderWithFixedPositions(
  questionNumbers: readonly number[],
  excludedNumbers: ReadonlySet<number>,
  random: IRandom,
): number[] {
  const freePositions: number[] = [];
  questionNumbers.forEach((questionNumber, index) => {
    if (!excludedNumbers.has(questionNumber)) freePositions.push(index);
  });

  const shuffledSources = shuffleInPlace([...freePositions], random);
  const order = questionNumbers.map((_, index) => index);
  freePositions.forEach((position, i) => {
    order[position] = shuffledSources[i]!;
  });
  return order;
}

function nonIdentityPermutation(random: IRandom): number[] {
  for (let attempt = 0; attempt < 12; attempt++) {
    const candidate = shuffleInPlace([...IDENTITY_4], random);
    if (candidate.some((value, index) => value !== index)) return candidate;
  }
  return [1, 0, 3, 2];
}

function signatureOf(plan: SetPlan): string {
  const orders = plan.sectionOrders.map((order) => order.join(',')).join('|');
  const options = [...plan.optionPermutations.entries()]
    .sort((a, b) => a[0] - b[0])
    .map(([questionNumber, permutation]) => `${questionNumber}:${permutation.join('')}`)
    .join(';');
  return `${orders}//${options}`;
}

function identitySignature(input: PlannerInput): string {
  return signatureOf({
    setNumber: 0,
    seed: '',
    sectionOrders: input.sections.map((numbers) => numbers.map((_, index) => index)),
    optionPermutations: new Map(),
  });
}
