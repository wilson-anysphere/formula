// @vitest-environment jsdom

import { describe, expect, it } from "vitest";

import { collabConnectionOptionsFromShareLink, parseCollabShareLink, serializeCollabShareLink } from "./collabLink.js";

describe("collabLink", () => {
  it("serializes and parses collaboration links (roundtrip)", () => {
    const link = serializeCollabShareLink(
      {
        wsUrl: "ws://127.0.0.1:1234",
        docId: "doc-123",
        token: " secret-token ",
      },
      { baseUrl: "http://localhost:4174/" },
    );

    const parsed = parseCollabShareLink(link, { baseUrl: "http://localhost:4174/" });
    expect(parsed).toEqual({
      wsUrl: "ws://127.0.0.1:1234",
      docId: "doc-123",
      token: "secret-token",
    });

    // Tokens should live in the hash fragment (not query params).
    const url = new URL(link);
    expect(url.searchParams.get("token")).toBeNull();
    expect(url.hash).toContain("token=secret-token");
  });

  it("constructs CollabSession connection options from a share link", () => {
    const link = "http://localhost:4174/?collab=1&wsUrl=ws%3A%2F%2Fexample.com%3A1234&docId=my-doc#token=my-token";
    const conn = collabConnectionOptionsFromShareLink(link, { baseUrl: "http://localhost:4174/" });
    expect(conn).toEqual({ wsUrl: "ws://example.com:1234", docId: "my-doc", token: "my-token" });
  });

  it("accepts legacy token-in-query share links", () => {
    const link =
      "http://localhost:4174/?collab=1&wsUrl=ws%3A%2F%2Fexample.com%3A1234&docId=my-doc&token=my-token";
    const parsed = parseCollabShareLink(link, { baseUrl: "http://localhost:4174/" });
    expect(parsed).toEqual({ wsUrl: "ws://example.com:1234", docId: "my-doc", token: "my-token" });
  });
});
