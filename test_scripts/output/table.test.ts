// test_scripts/output/table.test.ts
import { describe, it, expect } from 'vitest';
import { renderTable } from '../../src/output/table';

describe('renderTable', () => {
  it('renders headers and rows aligned to widest cell per column', () => {
    const out = renderTable(
      [
        { name: 'Alice', role: 'Manager' },
        { name: 'Bob', role: 'Eng' },
      ],
      [
        { key: 'name', header: 'Name' },
        { key: 'role', header: 'Role' },
      ],
    );
    const lines = out.split('\n');
    expect(lines[0]).toContain('Name');
    expect(lines[0]).toContain('Role');
    expect(lines[1]).toMatch(/^-+\s+-+$/);
    expect(lines[2]).toContain('Alice');
    expect(lines[2]).toContain('Manager');
    expect(lines[3]).toContain('Bob');
  });

  it('substitutes empty string for null/undefined cells', () => {
    const out = renderTable(
      [{ a: null, b: undefined }],
      [
        { key: 'a', header: 'A' },
        { key: 'b', header: 'B' },
      ],
    );
    expect(out).not.toContain('null');
    expect(out).not.toContain('undefined');
  });

  it('truncates long cell values to maxWidth and appends ellipsis', () => {
    const out = renderTable(
      [{ s: 'this is a very long string' }],
      [{ key: 's', header: 'S', maxWidth: 10 }],
    );
    expect(out).toMatch(/this is a…/);
  });
});
