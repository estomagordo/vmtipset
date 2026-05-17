/** Mall sheet anchors for free-text fields (merged G4:K5, G6:K7 in VM-tipset 2026). */
export const META_FREE_TEXT_CELLS = new Set(['G4', 'G6']);

export function isMetaFreeTextCell(address: string): boolean {
  return META_FREE_TEXT_CELLS.has(address.trim().toUpperCase().replace(/\$/g, ''));
}
