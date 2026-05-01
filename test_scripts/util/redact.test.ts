// test_scripts/util/redact.test.ts
import { describe, it, expect } from 'vitest';
import { redactBody } from '../../src/util/redact';

describe('redactBody', () => {
  it('truncates long strings to 200 chars with "…(truncated)"', () => {
    const long = 'x'.repeat(500);
    const out = redactBody(long);
    expect(out.length).toBeLessThanOrEqual(220);
    expect(out).toMatch(/…\(truncated\)$/);
  });

  it('returns short strings unchanged', () => {
    expect(redactBody('hello')).toBe('hello');
  });

  it('strips Bearer tokens from the body', () => {
    const out = redactBody('Authorization: Bearer eyJ.foo.bar in error');
    expect(out).not.toContain('eyJ.foo.bar');
    expect(out).toContain('Bearer [REDACTED]');
  });

  it('strips long base64-like blobs (>40 chars of [A-Za-z0-9+/_-])', () => {
    const blob = 'A'.repeat(60);
    const out = redactBody(`prefix ${blob} suffix`);
    expect(out).not.toContain(blob);
    expect(out).toContain('[REDACTED]');
  });
});
