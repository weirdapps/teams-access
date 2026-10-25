// test_scripts/auth/nav-budget.test.ts
//
// The headless navigation loop used to be fire-and-forget, so the browser closed
// 40s after the first bearer, while the loop was still on the Teams Calendar
// view. The Outlook /mail/ Graph provoker never ran. The loop is now awaited
// within the login budget and stops early once both required audiences are in.

import { describe, it, expect } from 'vitest';

import { awaitWithin, hasRequiredAudiences } from '../../src/auth/browser-capture';

describe('hasRequiredAudiences', () => {
  it('needs both Graph and chatsvcagg', () => {
    expect(hasRequiredAudiences(['https://chatsvcagg.teams.microsoft.com'])).toBe(false);
    expect(hasRequiredAudiences(['https://graph.microsoft.com'])).toBe(false);
    expect(
      hasRequiredAudiences([
        'https://api.spaces.skype.com',
        'https://graph.microsoft.com',
        'https://chatsvcagg.teams.microsoft.com',
      ]),
    ).toBe(true);
  });

  it('accepts Graph by app id', () => {
    expect(
      hasRequiredAudiences([
        '00000003-0000-0000-c000-000000000000',
        'https://chatsvcagg.teams.microsoft.com',
      ]),
    ).toBe(true);
  });
});

describe('awaitWithin', () => {
  it('returns true when the loop finishes inside the budget', async () => {
    expect(await awaitWithin(new Promise((r) => setTimeout(r, 5)), 1000)).toBe(true);
  });

  it('returns false when the budget runs out first', async () => {
    expect(await awaitWithin(new Promise(() => {}), 20)).toBe(false);
  });

  it('does not wait at all on a spent budget', async () => {
    expect(await awaitWithin(new Promise(() => {}), 0)).toBe(false);
    expect(await awaitWithin(new Promise(() => {}), -500)).toBe(false);
  });
});
