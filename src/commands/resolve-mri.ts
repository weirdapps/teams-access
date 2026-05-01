// src/commands/resolve-mri.ts — Translate Teams MRI to {id, email, displayName} via Graph /users/{aad-oid}.

import { GraphClient } from '../http/graph-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';

export interface ResolveMriOptions {
  session: Session;
  httpTimeoutMs: number;
  mri: string;
}

export interface ResolveMriResult {
  id: string;            // aad-oid
  email: string | null;  // mail field; null for guests
  displayName: string;
}

// aad-oid is canonically a hex GUID, but Graph also returns synthetic ids for
// guests (e.g. "x"); keep this lenient — caller passes through to Graph anyway.
const MRI_RE = /^8:orgid:([A-Za-z0-9-]+)$/;

export async function runResolveMri(opts: ResolveMriOptions): Promise<ResolveMriResult> {
  const m = MRI_RE.exec(opts.mri);
  if (!m) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: `Invalid MRI: expected "8:orgid:<aad-oid>", got "${opts.mri}"`,
    });
  }
  const oid = m[1];

  const graph = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
  try {
    const user = await graph.get<{
      id: string;
      mail?: string | null;
      displayName: string;
      userPrincipalName: string;
    }>(`/users/${encodeURIComponent(oid)}`);
    return {
      id: user.id,
      email: user.mail ?? null,
      displayName: user.displayName,
    };
  } catch (e) {
    if (e instanceof AuthRequiredError) {
      throw new ExitWithCode(ExitCode.AuthRequired, { code: 'auth_required', message: e.message });
    }
    if (e instanceof GraphHttpError) {
      // Surface 404 distinctly so callers can mark permanent_fail.
      throw new ExitWithCode(ExitCode.Upstream, {
        code: 'upstream',
        message: e.message,
        status: e.status,
      });
    }
    throw e;
  }
}
