// src/util/exit-codes.ts

export const ExitCode = {
  Success: 0,
  Internal: 1,
  InvalidInput: 2,
  Config: 3,
  AuthRequired: 4,
  Upstream: 5,
  Io: 6,
} as const;
export type ExitCodeValue = (typeof ExitCode)[keyof typeof ExitCode];

export interface ErrorPayload {
  code: string;
  message?: string;
  [key: string]: unknown;
}

export class ExitWithCode extends Error {
  constructor(
    public readonly code: ExitCodeValue,
    public readonly payload: ErrorPayload,
  ) {
    super(payload.message ?? payload.code);
    this.name = 'ExitWithCode';
  }
}
