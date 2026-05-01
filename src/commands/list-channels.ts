// src/commands/list-channels.ts — Graph /teams/{id}/channels (single team or all).

import { GraphClient } from '../http/graph-client';
import { AuthRequiredError, GraphHttpError } from '../http/errors';
import { ExitCode, ExitWithCode } from '../util/exit-codes';
import type { Session } from '../session/store';
import type { Channel, Team } from '../http/types';
import { writeSession } from '../session/store';

export interface ListChannelsOptions {
  session: Session;
  httpTimeoutMs: number;
  teamId?: string;
  allTeams?: boolean;
}

export interface ChannelWithTeam extends Channel {
  teamId: string;
  teamDisplayName?: string;
}

export interface ListChannelsResult {
  channels: ChannelWithTeam[];
}

/** Side effect: cache each team's General channel id into session (lazy). */
function cacheGeneralChannel(session: Session, teamUuid: string, channels: Channel[]): boolean {
  const general = channels.find(c => c.displayName === 'General');
  if (!general) return false;
  if (!session.generalChannelByTeamId) session.generalChannelByTeamId = {};
  if (session.generalChannelByTeamId[teamUuid] === general.id) return false; // no change
  session.generalChannelByTeamId[teamUuid] = general.id;
  return true;
}

export async function runListChannels(opts: ListChannelsOptions): Promise<ListChannelsResult> {
  if (!opts.teamId && !opts.allTeams) {
    throw new ExitWithCode(ExitCode.InvalidInput, {
      code: 'invalid_input',
      message: 'Must provide either --team-id <id> or --all-teams',
    });
  }
  const c = new GraphClient(opts.session, { httpTimeoutMs: opts.httpTimeoutMs });
  try {
    let dirty = false;

    if (opts.teamId) {
      const r = await c.get<{ value: Channel[] }>(`/teams/${encodeURIComponent(opts.teamId)}/channels`);
      dirty = cacheGeneralChannel(opts.session, opts.teamId, r.value) || dirty;
      if (dirty) writeSession(opts.session);
      return { channels: r.value.map(ch => ({ ...ch, teamId: opts.teamId! })) };
    }

    // --all-teams: fan out
    const teams = await c.get<{ value: Team[] }>('/me/joinedTeams');
    const out: ChannelWithTeam[] = [];
    for (const team of teams.value) {
      const r = await c.get<{ value: Channel[] }>(`/teams/${encodeURIComponent(team.id)}/channels`);
      dirty = cacheGeneralChannel(opts.session, team.id, r.value) || dirty;
      for (const ch of r.value) {
        out.push({ ...ch, teamId: team.id, teamDisplayName: team.displayName });
      }
    }
    if (dirty) writeSession(opts.session);
    return { channels: out };
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
