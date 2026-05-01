// test_scripts/util/exit-codes.test.ts
import { describe, it, expect } from 'vitest';
import { ExitCode, ExitWithCode } from '../../src/util/exit-codes';

describe('exit codes', () => {
  it('defines canonical numeric codes', () => {
    expect(ExitCode.Success).toBe(0);
    expect(ExitCode.Internal).toBe(1);
    expect(ExitCode.InvalidInput).toBe(2);
    expect(ExitCode.Config).toBe(3);
    expect(ExitCode.AuthRequired).toBe(4);
    expect(ExitCode.Upstream).toBe(5);
    expect(ExitCode.Io).toBe(6);
  });

  it('ExitWithCode carries code + structured payload', () => {
    const e = new ExitWithCode(ExitCode.Config, { code: 'config_error', missingSetting: 'X' });
    expect(e.code).toBe(3);
    expect(e.payload.code).toBe('config_error');
    expect(e.payload.missingSetting).toBe('X');
  });
});
