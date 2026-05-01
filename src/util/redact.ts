// src/util/redact.ts

const MAX_LEN = 200;
const BEARER_RE = /Bearer\s+[A-Za-z0-9+/=._-]+/g;
const LONG_BLOB_RE = /[A-Za-z0-9+/_-]{40,}/g;

export function redactBody(input: string): string {
  let s = input;
  s = s.replace(BEARER_RE, 'Bearer [REDACTED]');
  s = s.replace(LONG_BLOB_RE, '[REDACTED]');
  if (s.length > MAX_LEN) {
    s = s.slice(0, MAX_LEN) + '…(truncated)';
  }
  return s;
}
