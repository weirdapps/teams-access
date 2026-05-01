// src/util/time.ts

const UNIT_MS: Record<string, number> = {
  m: 60_000,
  h: 3_600_000,
  d: 86_400_000,
  w: 604_800_000,
};

const DURATION_RE = /^(\d+)([mhdw])$/;
const NOW_OFFSET_RE = /^now\s*([+\-])\s*(\d+)([mhdw])$/;

export function parseRelativeTime(input: string): Date {
  const s = input.trim();

  if (s === 'now') return new Date();

  const isoLike = /^\d{4}-\d{2}-\d{2}/;
  if (isoLike.test(s)) {
    const d = new Date(s);
    if (isNaN(d.getTime())) {
      throw new Error(`Unrecognised time format (invalid ISO): "${input}"`);
    }
    return d;
  }

  const dur = DURATION_RE.exec(s);
  if (dur) {
    const n = Number(dur[1]);
    const ms = n * UNIT_MS[dur[2]];
    return new Date(Date.now() - ms);
  }

  const offset = NOW_OFFSET_RE.exec(s);
  if (offset) {
    const sign = offset[1] === '+' ? 1 : -1;
    const n = Number(offset[2]);
    const ms = sign * n * UNIT_MS[offset[3]];
    return new Date(Date.now() + ms);
  }

  throw new Error(`Unrecognised time format: "${input}"`);
}
