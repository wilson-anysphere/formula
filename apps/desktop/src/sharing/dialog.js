import { DOCUMENT_ROLES } from "../../../../packages/collab/permissions/index.js";

export class SharingDialogController {
  constructor({ apiClient, docId, cookie }) {
    this.apiClient = apiClient;
    this.docId = docId;
    this.cookie = cookie;
  }

  async load() {
    const data = await this.apiClient.listCollaborators({ docId: this.docId, cookie: this.cookie });
    this.collaborators = data.members ?? [];
    return this.collaborators;
  }

  async invite({ email, role }) {
    return this.apiClient.inviteUser({ docId: this.docId, cookie: this.cookie, email, role });
  }

  async setRole({ email, role }) {
    if (!DOCUMENT_ROLES.includes(role)) throw new Error("Invalid role");
    return this.apiClient.changeRole({ docId: this.docId, cookie: this.cookie, email, role });
  }

  async generateShareLink({ visibility = "private", role = "viewer", expiresInSeconds } = {}) {
    return this.apiClient.createShareLink({
      docId: this.docId,
      cookie: this.cookie,
      visibility,
      role,
      expiresInSeconds
    });
  }
}
