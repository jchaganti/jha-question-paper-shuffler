import { randomUUID } from 'node:crypto';

/** Minimal randomness contract, so shuffling can be tested deterministically. */
export interface IRandom {
  /** Integer in [0, maxExclusive). */
  nextInt(maxExclusive: number): number;
}

/**
 * Deterministic PRNG (splitmix32) keyed by a text seed.
 *
 * Every run records its seed in the report: re-running with the same seed, source
 * file and options reproduces the same sets, which makes a generated paper auditable.
 */
export class SeededRandom implements IRandom {
  private state: number;

  constructor(readonly seed: string) {
    this.state = hashSeed(seed);
  }

  static newSeed(): string {
    return randomUUID();
  }

  nextInt(maxExclusive: number): number {
    if (maxExclusive <= 0) return 0;
    return Math.floor(this.nextFloat() * maxExclusive);
  }

  private nextFloat(): number {
    this.state = (this.state + 0x9e3779b9) | 0;
    let z = this.state;
    z = Math.imul(z ^ (z >>> 16), 0x21f0aaad);
    z = Math.imul(z ^ (z >>> 15), 0x735a2d97);
    z = z ^ (z >>> 15);
    return (z >>> 0) / 4294967296;
  }
}

function hashSeed(seed: string): number {
  let h = 2166136261;
  for (let i = 0; i < seed.length; i++) {
    h ^= seed.charCodeAt(i);
    h = Math.imul(h, 16777619);
  }
  return h | 0;
}

/** In-place Fisher-Yates shuffle. */
export function shuffleInPlace<T>(items: T[], random: IRandom): T[] {
  for (let i = items.length - 1; i > 0; i--) {
    const j = random.nextInt(i + 1);
    const a = items[i]!;
    const b = items[j]!;
    items[i] = b;
    items[j] = a;
  }
  return items;
}
