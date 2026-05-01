// test_scripts/util/time.test.ts
import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { parseRelativeTime } from '../../src/util/time';

describe('parseRelativeTime', () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.setSystemTime(new Date('2026-05-01T12:00:00Z'));
  });
  afterEach(() => vi.useRealTimers());

  it('parses ISO timestamps verbatim', () => {
    const d = parseRelativeTime('2026-04-25T08:30:00Z');
    expect(d.toISOString()).toBe('2026-04-25T08:30:00.000Z');
  });

  it('parses "now" as the current instant', () => {
    expect(parseRelativeTime('now').toISOString()).toBe('2026-05-01T12:00:00.000Z');
  });

  it('parses bare durations as offsets from now', () => {
    expect(parseRelativeTime('24h').toISOString()).toBe('2026-04-30T12:00:00.000Z');
    expect(parseRelativeTime('3d').toISOString()).toBe('2026-04-28T12:00:00.000Z');
    expect(parseRelativeTime('1w').toISOString()).toBe('2026-04-24T12:00:00.000Z');
    expect(parseRelativeTime('30m').toISOString()).toBe('2026-05-01T11:30:00.000Z');
  });

  it('parses "now - 24h" and "now + 14d"', () => {
    expect(parseRelativeTime('now - 24h').toISOString()).toBe('2026-04-30T12:00:00.000Z');
    expect(parseRelativeTime('now + 14d').toISOString()).toBe('2026-05-15T12:00:00.000Z');
  });

  it('throws on unrecognised format', () => {
    expect(() => parseRelativeTime('yesterday')).toThrow(/unrecognised time/i);
    expect(() => parseRelativeTime('2 hours ago')).toThrow();
  });
});
