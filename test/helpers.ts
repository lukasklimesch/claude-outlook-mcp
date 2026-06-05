// Shared test utilities.

import type { OutlookRunner, PsResult } from "../src/types.js";

export interface RecordedCall {
  script: string;
  payload?: Record<string, unknown>;
}

/**
 * A fake {@link OutlookRunner} that records every call and returns canned
 * responses, letting us exercise the entire server/dispatch flow without a
 * Windows host or a live Outlook.
 */
export class FakeRunner implements OutlookRunner {
  public calls: RecordedCall[] = [];
  private responder: (script: string, payload?: Record<string, unknown>) => PsResult;

  constructor(
    responder: (
      script: string,
      payload?: Record<string, unknown>,
    ) => PsResult = () => ({ ok: true, data: [] }),
  ) {
    this.responder = responder;
  }

  async run<T = unknown>(
    script: string,
    payload?: Record<string, unknown>,
  ): Promise<PsResult<T>> {
    this.calls.push({ script, payload });
    return this.responder(script, payload) as PsResult<T>;
  }

  last(): RecordedCall {
    const c = this.calls[this.calls.length - 1];
    if (!c) throw new Error("FakeRunner: no calls recorded");
    return c;
  }
}
