// src/http/errors.ts

export class GraphHttpError extends Error {
  public readonly retryable: boolean = false;
  constructor(
    public readonly status: number,
    public readonly graphCode: string,
    message: string,
    public readonly requestId?: string,
  ) {
    super(`Graph ${status} ${graphCode}: ${message}` + (requestId ? ` [req=${requestId}]` : ''));
    this.name = 'GraphHttpError';
  }
}

export class AuthRequiredError extends GraphHttpError {
  constructor(message: string, requestId?: string) {
    super(401, 'Unauthenticated', message, requestId);
    this.name = 'AuthRequiredError';
  }
}

export class UpstreamError extends GraphHttpError {
  public override readonly retryable = true;
  constructor(status: number, graphCode: string, message: string, requestId?: string) {
    super(status, graphCode, message, requestId);
    this.name = 'UpstreamError';
  }
}
