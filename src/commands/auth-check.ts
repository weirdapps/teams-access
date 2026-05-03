// src/commands/auth-check.ts
import { GraphClient } from '../http/graph-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';
import type { User } from '../http/types';
import { decodeJwt } from '../session/jwt';

export interface AuthCheckOptions {
  session: Session;
  httpTimeoutMs: number;
}

export interface AuthCheckResult {
  status: 'ok';
  tokenExpiresAt?: string;
  account: { upn?: string; displayName?: string; oid?: string };
}

export async function runAuthCheck(opts: AuthCheckOptions): Promise<AuthCheckResult> {
  const client = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
  try {
    const me = await client.get<User>('/me');
    let tokenExpiresAt: string | undefined;
    try {
      const claims = decodeJwt(opts.session.bearerToken);
      if (claims.exp) tokenExpiresAt = new Date(claims.exp * 1000).toISOString();
    } catch {
      /* ignore */
    }
    return {
      status: 'ok',
      tokenExpiresAt,
      account: {
        upn: me.userPrincipalName,
        displayName: me.displayName,
        oid: me.id,
      },
    };
  } catch (e) {
    if (e instanceof AuthRequiredError) {
      throw new ExitWithCode(ExitCode.AuthRequired, {
        code: 'auth_required',
        message: 'Cached session is no longer accepted. Run `teams-cli login`.',
        requestId: e.requestId,
      });
    }
    if (e instanceof GraphHttpError) {
      throw new ExitWithCode(ExitCode.Upstream, {
        code: 'upstream',
        message: e.message,
        status: e.status,
      });
    }
    throw e;
  }
}
