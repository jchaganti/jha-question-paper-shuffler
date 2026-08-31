/**
 * The generation report as structure rather than as text.
 *
 * `ReportWriter` builds one of these and `renderReportPdf` typesets it. Keeping the two
 * apart means what the report *says* can be tested without going through a PDF, and the
 * typesetting can be tested without going through a whole generation run.
 */

/** A `label: value` line - the run's settings and each set's facts. */
export interface ReportFact {
  readonly label: string;
  readonly value: string;
  /** Fixed-width, for a file name, a path or a seed that must be copied exactly. */
  readonly mono?: boolean;
  /** Drawn in green or red: a verification result. */
  readonly tone?: 'ok' | 'bad';
}

export interface ReportColumn {
  readonly header: string;
  /** Relative share of any width left over once every column has its natural width. */
  readonly weight?: number;
  readonly mono?: boolean;
  readonly align?: 'left' | 'right';
}

export type ReportBlock =
  /** A section title. `level: 2` is a set's own heading, smaller than a section's. */
  | { readonly kind: 'heading'; readonly text: string; readonly level?: 1 | 2 }
  | { readonly kind: 'paragraph'; readonly text: string }
  | { readonly kind: 'facts'; readonly items: readonly ReportFact[] }
  | {
      readonly kind: 'table';
      readonly columns: readonly ReportColumn[];
      readonly rows: readonly (readonly string[])[];
    };

export interface ReportDocument {
  readonly title: string;
  /** One line under the title: which paper, generated when. */
  readonly subtitle: string;
  readonly blocks: readonly ReportBlock[];
}
