import { describe, expect, it } from "vitest";
import type { SharedConversationSource } from "./shared-conversation-index";
import { buildSharedConversationsByMri } from "./shared-conversation-index";

describe("buildSharedConversationsByMri", () => {
  it("indexes each conversation once per participant mri", () => {
    const result = buildSharedConversationsByMri(
      [
        {
          id: "dm-1",
          conversation: { id: "dm-1" },
          title: "Pat Lee",
          preview: "Hey",
          kind: "dm",
          isFavorite: false,
          avatarMri: "8:orgid:pat",
          sideTime: "10:30",
          searchText: "pat lee hey dm",
        },
        {
          id: "group-1",
          conversation: {
            id: "group-1",
            members: [
              { id: "8:orgid:pat", role: "User", isMri: true },
              { id: "8:orgid:self", role: "User", isMri: true },
              { id: "8:orgid:pat", role: "User", isMri: true },
            ],
          },
          title: "Design",
          preview: "Review notes",
          kind: "group",
          isFavorite: false,
          sideTime: "09:15",
          searchText: "design review notes group",
        },
      ] satisfies SharedConversationSource[],
      {},
      {},
      {},
    );

    expect(result["8:orgid:pat"]).toEqual([
      {
        id: "dm-1",
        title: "Pat Lee",
        kind: "dm",
        preview: "Hey",
        sideTime: "10:30",
      },
      {
        id: "group-1",
        title: "Design",
        kind: "group",
        preview: "Review notes",
        sideTime: "09:15",
      },
    ]);
    expect(result["8:orgid:self"]).toEqual([
      {
        id: "group-1",
        title: "Design",
        kind: "group",
        preview: "Review notes",
        sideTime: "09:15",
      },
    ]);
  });

  it("prefers fetched member lists over sparse sidebar data", () => {
    const result = buildSharedConversationsByMri(
      [
        {
          id: "group-1",
          conversation: { id: "group-1" },
          title: "Design",
          preview: "Review notes",
          kind: "group",
          isFavorite: false,
          sideTime: "09:15",
          searchText: "design review notes group",
        },
      ] satisfies SharedConversationSource[],
      {
        "group-1": [{ id: "8:orgid:pat", role: "User", isMri: true }],
      },
      {},
      {},
    );

    expect(result["8:orgid:pat"]).toEqual([
      {
        id: "group-1",
        title: "Design",
        kind: "group",
        preview: "Review notes",
        sideTime: "09:15",
      },
    ]);
  });

  it("matches opaque members by email and display name without scanning every conversation for every mri", () => {
    const result = buildSharedConversationsByMri(
      [
        {
          id: "group-1",
          conversation: {
            id: "group-1",
            members: [
              {
                id: "29:opaque-member",
                role: "User",
                isMri: false,
                displayName: "Pat Lee",
                userPrincipalName: "pat@test.com",
              },
            ],
          },
          title: "Project alpha",
          preview: "Planning update",
          kind: "group",
          isFavorite: false,
          sideTime: "09:15",
          searchText: "project alpha planning update group",
        },
      ] satisfies SharedConversationSource[],
      {},
      { "8:orgid:pat": "Pat Lee" },
      { "8:orgid:pat": "pat@test.com" },
    );

    expect(result["8:orgid:pat"]).toEqual([
      {
        id: "group-1",
        title: "Project alpha",
        kind: "group",
        preview: "Planning update",
        sideTime: "09:15",
      },
    ]);
  });
});
