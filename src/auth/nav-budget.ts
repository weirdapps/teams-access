// src/auth/nav-budget.ts
//
// Bounded wait for the headless navigation loop in browser-capture.ts, kept here
// so it is unit-tested: browser-capture.ts drives Playwright and is excluded
// from coverage.

/**
 * Wait for the navigation loop, but never past `budgetMs`. Resolves true when
 * the loop finished, false when the budget ran out first.
 */
export async function awaitWithin(done: Promise<unknown>, budgetMs: number): Promise<boolean> {
  if (budgetMs <= 0) return false;
  let timer: ReturnType<typeof setTimeout> | undefined;
  const expired = new Promise<false>((r) => {
    timer = setTimeout(() => r(false), budgetMs);
  });
  try {
    return await Promise.race([done.then(() => true as const), expired]);
  } finally {
    clearTimeout(timer);
  }
}
