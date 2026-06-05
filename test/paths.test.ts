import { describe, expect, test } from "bun:test";
import {
  defaultDownloadDir,
  normalizeWindowsPath,
} from "../src/util/paths.js";

describe("normalizeWindowsPath", () => {
  test("converts slashes and strips surrounding quotes", () => {
    expect(normalizeWindowsPath('"C:/Users/Joe/file.txt"')).toBe(
      "C:\\Users\\Joe\\file.txt",
    );
    expect(normalizeWindowsPath("C:/a/b")).toBe("C:\\a\\b");
  });
});

describe("defaultDownloadDir", () => {
  test("honors explicit override", () => {
    expect(
      defaultDownloadDir({ OUTLOOK_MCP_DOWNLOAD_DIR: "D:/Saved" }),
    ).toBe("D:\\Saved");
  });

  test("falls back to the user's Downloads folder", () => {
    expect(defaultDownloadDir({ USERPROFILE: "C:\\Users\\Joe" })).toBe(
      "C:\\Users\\Joe\\Downloads",
    );
  });

  test("falls back to cwd when nothing is set", () => {
    expect(defaultDownloadDir({})).toBe(".");
  });
});
