import { describe, expect, it } from 'vitest';
import { SeededRandom, shuffleInPlace } from '../src/core/shuffle/Random';
import { ShufflePlanner, orderWithFixedPositions } from '../src/core/shuffle/ShufflePlanner';
import type { GenerationRequest } from '../src/shared/types';

const request = (overrides: Partial<GenerationRequest> = {}): GenerationRequest => ({
  sourceFile: 'paper.docx',
  shuffleQuestions: true,
  shuffleOptions: true,
  questionExclusions: [],
  optionExclusions: [],
  setCount: 3,
  seed: 'test-seed',
  ...overrides,
});

const input = (overrides: Partial<Parameters<ShufflePlanner['plan']>[0]> = {}) => ({
  sections: [
    [1, 2, 3, 4, 5, 6, 7, 8, 9, 10],
    [11, 12, 13, 14, 15],
  ],
  shufflableOptionQuestions: [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15],
  request: request(),
  ...overrides,
});

describe('SeededRandom', () => {
  it('is deterministic for a given seed', () => {
    const first = [...Array(20)].map((_, i) => new SeededRandom('abc').nextInt(100 + i));
    const second = [...Array(20)].map((_, i) => new SeededRandom('abc').nextInt(100 + i));
    expect(first).toEqual(second);
  });

  it('produces a permutation, never losing or duplicating an item', () => {
    const items = [...Array(50)].map((_, i) => i);
    const shuffled = shuffleInPlace([...items], new SeededRandom('x'));
    expect([...shuffled].sort((a, b) => a - b)).toEqual(items);
  });
});

describe('orderWithFixedPositions', () => {
  it('keeps excluded questions at their own position', () => {
    const order = orderWithFixedPositions([1, 2, 3, 4, 5], new Set([2, 4]), new SeededRandom('seed'));
    expect(order[1]).toBe(1); // question 2 stays at index 1
    expect(order[3]).toBe(3); // question 4 stays at index 3
    expect([...order].sort((a, b) => a - b)).toEqual([0, 1, 2, 3, 4]);
  });

  it('is a no-op when every question is excluded', () => {
    const order = orderWithFixedPositions([1, 2, 3], new Set([1, 2, 3]), new SeededRandom('seed'));
    expect(order).toEqual([0, 1, 2]);
  });
});

describe('ShufflePlanner', () => {
  it('produces one plan per set, each different from the original and from each other', () => {
    const plans = new ShufflePlanner().plan(input());
    expect(plans).toHaveLength(3);

    const signatures = plans.map((plan) => JSON.stringify(plan.sectionOrders));
    expect(new Set(signatures).size).toBe(3);
    for (const plan of plans) {
      expect(plan.sectionOrders[0]).not.toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    }
  });

  it('never moves a question between subjects', () => {
    const plans = new ShufflePlanner().plan(input());
    for (const plan of plans) {
      expect([...plan.sectionOrders[0]!].sort((a, b) => a - b)).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
      expect([...plan.sectionOrders[1]!].sort((a, b) => a - b)).toEqual([0, 1, 2, 3, 4]);
    }
  });

  it('honours both exclusion lists', () => {
    const plans = new ShufflePlanner().plan(
      input({ request: request({ questionExclusions: [3, 7], optionExclusions: [1, 2] }) }),
    );
    for (const plan of plans) {
      expect(plan.sectionOrders[0]![2]).toBe(2);
      expect(plan.sectionOrders[0]![6]).toBe(6);
      expect(plan.optionPermutations.has(1)).toBe(false);
      expect(plan.optionPermutations.has(2)).toBe(false);
      expect(plan.optionPermutations.has(3)).toBe(true);
    }
  });

  it('leaves question order alone when only options are shuffled', () => {
    const plans = new ShufflePlanner().plan({
      ...input(),
      request: request({ shuffleQuestions: false, setCount: 1 }),
    });
    expect(plans[0]!.sectionOrders[0]).toEqual([0, 1, 2, 3, 4, 5, 6, 7, 8, 9]);
    expect(plans[0]!.optionPermutations.size).toBe(15);
  });

  it('leaves options alone when only questions are shuffled', () => {
    const plans = new ShufflePlanner().plan({
      ...input(),
      request: request({ shuffleOptions: false, setCount: 1 }),
    });
    expect(plans[0]!.optionPermutations.size).toBe(0);
  });

  it('never returns the identity permutation for an option set', () => {
    const plans = new ShufflePlanner().plan(input({ request: request({ setCount: 10 }) }));
    for (const plan of plans) {
      for (const permutation of plan.optionPermutations.values()) {
        expect(permutation).not.toEqual([0, 1, 2, 3]);
        expect([...permutation].sort()).toEqual([0, 1, 2, 3]);
      }
    }
  });

  it('reproduces identical plans for the same seed', () => {
    const first = new ShufflePlanner().plan(input());
    const second = new ShufflePlanner().plan(input());
    expect(JSON.stringify(first.map((p) => [p.sectionOrders, [...p.optionPermutations]]))).toBe(
      JSON.stringify(second.map((p) => [p.sectionOrders, [...p.optionPermutations]])),
    );
  });

  it('produces different plans without a seed', () => {
    const first = new ShufflePlanner().plan(input({ request: request({ seed: undefined }) }));
    const second = new ShufflePlanner().plan(input({ request: request({ seed: undefined }) }));
    expect(JSON.stringify(first.map((p) => p.sectionOrders))).not.toBe(
      JSON.stringify(second.map((p) => p.sectionOrders)),
    );
  });
});
