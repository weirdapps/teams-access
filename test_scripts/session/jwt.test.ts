// test_scripts/session/jwt.test.ts
import { describe, it, expect } from 'vitest';
import { decodeJwt } from '../../src/session/jwt';

describe('decodeJwt', () => {
  it('extracts payload claims from a valid token', () => {
    // header={alg:HS256,typ:JWT}, payload={upn:"x@y.com",exp:1900000000}
    const token =
      'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.' +
      'eyJ1cG4iOiJ4QHkuY29tIiwiZXhwIjoxOTAwMDAwMDAwfQ.' +
      'sig';
    const payload = decodeJwt(token);
    expect(payload.upn).toBe('x@y.com');
    expect(payload.exp).toBe(1900000000);
  });

  it('throws on malformed token (wrong segment count)', () => {
    expect(() => decodeJwt('a.b')).toThrow(/three segments/i);
  });

  it('throws on non-base64-url payload', () => {
    expect(() => decodeJwt('a.!!!.c')).toThrow();
  });
});
