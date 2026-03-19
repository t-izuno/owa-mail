import { describe, it, expect } from "vitest";
import { parseArgs } from "./index.js";

describe("parseArgs", () => {
  it("コマンドとpositional引数を解析する", () => {
    const result = parseArgs(["read", "item123"]);
    expect(result.cmd).toBe("read");
    expect(result.rest).toEqual(["item123"]);
    expect(result.flags).toEqual({});
  });

  it("値付きフラグを解析する", () => {
    const result = parseArgs(["list", "--folder", "inbox", "--count", "50"]);
    expect(result.cmd).toBe("list");
    expect(result.flags.folder).toBe("inbox");
    expect(result.flags.count).toBe("50");
  });

  it("--unread をbooleanフラグとして解析する", () => {
    const result = parseArgs(["list", "--unread"]);
    expect(result.cmd).toBe("list");
    expect(result.flags.unread).toBe(true);
  });

  it("--unread の後にフラグが続いても正しく解析する", () => {
    const result = parseArgs(["list", "--unread", "--folder", "sent"]);
    expect(result.cmd).toBe("list");
    expect(result.flags.unread).toBe(true);
    expect(result.flags.folder).toBe("sent");
  });

  it("引数がない場合cmdがundefinedになる", () => {
    const result = parseArgs([]);
    expect(result.cmd).toBeUndefined();
    expect(result.rest).toEqual([]);
    expect(result.flags).toEqual({});
  });

  it("値なしの未知フラグはtrueになる", () => {
    const result = parseArgs(["list", "--verbose"]);
    expect(result.flags.verbose).toBe(true);
  });

  it("--save-dir のようなハイフン付きフラグを解析する", () => {
    const result = parseArgs(["attachment", "item123", "--save-dir", "/tmp/mail"]);
    expect(result.cmd).toBe("attachment");
    expect(result.rest).toEqual(["item123"]);
    expect(result.flags["save-dir"]).toBe("/tmp/mail");
  });
});
