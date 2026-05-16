/** 1-based column index → Excel letters (1 → A). */
export function colIndexToLetters(index: number): string {
  let n = index;
  let s = '';
  while (n > 0) {
    n -= 1;
    s = String.fromCharCode((n % 26) + 65) + s;
    n = Math.floor(n / 26);
  }
  return s;
}

export function lettersToColIndex(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    n = n * 26 + (ch.charCodeAt(0) - 64);
  }
  return n;
}

export function formatA1(row: number, col: number): string {
  return `${colIndexToLetters(col)}${row}`;
}

const A1 = /^([A-Za-z]+)(\d+)$/;

export function parseA1(address: string): { row: number; col: number } {
  const m = address.trim().match(A1);
  if (!m) throw new Error(`Bad A1 address: ${address}`);
  return { col: lettersToColIndex(m[1]), row: parseInt(m[2], 10) };
}

export function splitRangeSpecs(raw: string): string[] {
  return raw.trim().split(/\s+/).filter(Boolean);
}

export function rangeBounds(spec: string): { c0: number; r0: number; c1: number; r1: number } {
  const [a, b] = spec.split(':');
  const p0 = parseA1(a);
  const p1 = parseA1(b);
  const c0 = Math.min(p0.col, p1.col);
  const c1 = Math.max(p0.col, p1.col);
  const r0 = Math.min(p0.row, p1.row);
  const r1 = Math.max(p0.row, p1.row);
  return { c0, r0, c1, r1 };
}

/** Intersect with 1-based inclusive grid [1..rowCount] × [1..colCount]. */
export function clipRangeBounds(
  b: { c0: number; r0: number; c1: number; r1: number },
  rowCount: number,
  colCount: number,
): { c0: number; r0: number; c1: number; r1: number } | null {
  const c0 = Math.max(1, b.c0);
  const r0 = Math.max(1, b.r0);
  const c1 = Math.min(colCount, b.c1);
  const r1 = Math.min(rowCount, b.r1);
  if (c0 > c1 || r0 > r1) return null;
  return { c0, r0, c1, r1 };
}
