import { describe, expect, test } from "bun:test";
import { extractTrailingJson, parsePsOutput } from "../src/util/parse.js";

describe("parsePsOutput", () => {
  test("parses a success envelope", () => {
    const r = parsePsOutput('{"ok":true,"data":[{"subject":"hi"}]}');
    expect(r.ok).toBe(true);
    expect(Array.isArray(r.data)).toBe(true);
  });

  test("parses an error envelope", () => {
    const r = parsePsOutput('{"ok":false,"error":"boom","category":"outlook"}');
    expect(r.ok).toBe(false);
    expect(r.error).toBe("boom");
    expect(r.category).toBe("outlook");
  });

  test("wraps a bare value as success data", () => {
    const r = parsePsOutput('{"version":"16.0"}');
    expect(r.ok).toBe(true);
    expect((r.data as any).version).toBe("16.0");
  });

  test("empty output is a structured error", () => {
    const r = parsePsOutput("   ");
    expect(r.ok).toBe(false);
    expect(r.category).toBe("empty");
  });

  test("non-JSON output is a parse error", () => {
    const r = parsePsOutput("totally not json at all");
    expect(r.ok).toBe(false);
    expect(r.category).toBe("parse");
  });

  test("recovers JSON that follows incidental warning text", () => {
    const out = 'WARNING: slow\n{"ok":true,"data":{"n":1}}';
    const r = parsePsOutput(out);
    expect(r.ok).toBe(true);
    expect((r.data as any).n).toBe(1);
  });

  test("strips a UTF-8 BOM", () => {
    const r = parsePsOutput('﻿{"ok":true,"data":5}');
    expect(r.ok).toBe(true);
    expect(r.data).toBe(5);
  });
});

describe("extractTrailingJson", () => {
  test("ignores braces inside strings", () => {
    const v = extractTrailingJson('noise {"a":"}{","b":2} trailer') as any;
    expect(v.a).toBe("}{");
    expect(v.b).toBe(2);
  });

  test("handles escaped quotes within strings", () => {
    const v = extractTrailingJson('{"a":"he said \\"hi\\"","b":1}') as any;
    expect(v.a).toBe('he said "hi"');
    expect(v.b).toBe(1);
  });

  test("returns undefined when no JSON present", () => {
    expect(extractTrailingJson("plain text")).toBeUndefined();
  });

  test("is not hijacked by a stray bracket before the real object", () => {
    const v = extractTrailingJson('a[b] noise {"ok":true,"data":1}') as any;
    expect(v.ok).toBe(true);
    expect(v.data).toBe(1);
  });

  test("returns the longest valid JSON value", () => {
    const v = extractTrailingJson('{"x":1} and {"x":1,"y":2,"z":3}') as any;
    expect(v).toEqual({ x: 1, y: 2, z: 3 });
  });
});
