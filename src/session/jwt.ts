// src/session/jwt.ts

export interface JwtPayload {
  upn?: string;
  unique_name?: string;
  oid?: string;
  tid?: string;
  exp?: number;
  iat?: number;
  appid?: string;
  scp?: string;
  [key: string]: unknown;
}

export function decodeJwt(token: string): JwtPayload {
  const parts = token.split('.');
  if (parts.length !== 3) {
    throw new Error('JWT must have three segments separated by "."');
  }
  const payloadB64Url = parts[1];
  const payloadB64 = payloadB64Url.replace(/-/g, '+').replace(/_/g, '/');
  const padded = payloadB64 + '='.repeat((4 - (payloadB64.length % 4)) % 4);
  const json = Buffer.from(padded, 'base64').toString('utf8');
  return JSON.parse(json) as JwtPayload;
}
