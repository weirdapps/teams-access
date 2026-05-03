// src/output/table.ts
export interface ColumnSpec {
  key: string;
  header: string;
  maxWidth?: number;
}

function cellOf(row: Record<string, unknown>, col: ColumnSpec): string {
  const v = row[col.key];
  if (v == null) return '';
  let s = String(v);
  if (col.maxWidth && s.length > col.maxWidth) {
    s = s.slice(0, col.maxWidth - 1) + '…';
  }
  return s;
}

export function renderTable(rows: Array<Record<string, unknown>>, columns: ColumnSpec[]): string {
  const headerCells = columns.map((c) => c.header);
  const bodyRows = rows.map((r) => columns.map((c) => cellOf(r, c)));

  const widths = columns.map((c, i) =>
    Math.max(headerCells[i].length, ...bodyRows.map((r) => r[i].length), 0),
  );

  const fmt = (cells: string[]) =>
    cells
      .map((c, i) => c.padEnd(widths[i]))
      .join('  ')
      .trimEnd();

  const lines: string[] = [];
  lines.push(fmt(headerCells));
  lines.push(fmt(widths.map((w) => '-'.repeat(w))));
  for (const r of bodyRows) lines.push(fmt(r));
  return lines.join('\n');
}

export function writeTable(rows: Array<Record<string, unknown>>, columns: ColumnSpec[]): void {
  process.stdout.write(renderTable(rows, columns) + '\n');
}
