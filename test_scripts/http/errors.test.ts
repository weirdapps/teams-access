// test_scripts/http/errors.test.ts
import { describe, it, expect } from 'vitest';
import { GraphHttpError, AuthRequiredError, UpstreamError } from '../../src/http/errors';

describe('http errors', () => {
  it('GraphHttpError carries status, code, message, requestId', () => {
    const e = new GraphHttpError(400, 'BadRequest', 'invalid query', 'req-123');
    expect(e.status).toBe(400);
    expect(e.graphCode).toBe('BadRequest');
    expect(e.requestId).toBe('req-123');
    expect(e.message).toContain('invalid query');
  });

  it('AuthRequiredError extends GraphHttpError with retryable=false', () => {
    const e = new AuthRequiredError('token expired', 'req-x');
    expect(e).toBeInstanceOf(GraphHttpError);
    expect(e.status).toBe(401);
    expect(e.retryable).toBe(false);
  });

  it('UpstreamError extends GraphHttpError with retryable=true', () => {
    const e = new UpstreamError(503, 'Throttled', 'try again later', 'req-y');
    expect(e).toBeInstanceOf(GraphHttpError);
    expect(e.retryable).toBe(true);
  });
});
