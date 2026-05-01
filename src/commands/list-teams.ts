// src/commands/list-teams.ts — Graph /me/joinedTeams.

import { GraphClient } from '../http/graph-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';
import type { Team } from '../http/types';

export interface ListTeamsOptions {
  session: Session;
  httpTimeoutMs: number;
}

export interface ListTeamsResult {
  teams: Team[];
}

export async function runListTeams(opts: ListTeamsOptions): Promise<ListTeamsResult> {
  const c = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
  try {
    const r = await c.get<{ value: Team[] }>('/me/joinedTeams');
    return { teams: r.value };
  } catch (e) {
    if (e instanceof AuthRequiredError) {
      throw new ExitWithCode(ExitCode.AuthRequired, {
        code: 'auth_required',
        message: e.message,
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
