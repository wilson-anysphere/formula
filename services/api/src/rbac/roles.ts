export type DocumentRole = "owner" | "admin" | "editor" | "commenter" | "viewer";
export type OrgRole = "owner" | "admin" | "member";

export type DocumentAction = "read" | "edit" | "comment" | "share" | "admin";

export function canDocument(role: DocumentRole, action: DocumentAction): boolean {
  switch (role) {
    case "owner":
      return true;
    case "admin":
      return true;
    case "editor":
      return action === "read" || action === "edit" || action === "comment";
    case "commenter":
      return action === "read" || action === "comment";
    case "viewer":
      return action === "read";
  }
}

export function isOrgAdmin(role: OrgRole): boolean {
  return role === "owner" || role === "admin";
}
