export function createSharingApiClient({ baseUrl }) {
  if (!baseUrl) throw new Error("baseUrl is required");

  async function request({ path, method, cookie, body }) {
    const res = await fetch(`${baseUrl}${path}`, {
      method,
      credentials: "include",
      headers: {
        ...(cookie ? { cookie } : {}),
        "content-type": "application/json"
      },
      body: body ? JSON.stringify(body) : undefined
    });
    const json = await res.json().catch(() => ({}));
    if (!res.ok) {
      const error = new Error(json.error ?? `Request failed (${res.status})`);
      error.status = res.status;
      error.details = json.details;
      throw error;
    }
    return json;
  }

  return {
    listCollaborators: ({ docId, cookie }) =>
      request({ path: `/docs/${docId}/members`, method: "GET", cookie }),
    inviteUser: ({ docId, cookie, email, role }) =>
      request({
        path: `/docs/${docId}/invite`,
        method: "POST",
        cookie,
        body: { email, role }
      }),
    changeRole: ({ docId, cookie, email, role }) =>
      request({ path: `/docs/${docId}/invite`, method: "POST", cookie, body: { email, role } }),
    createShareLink: ({ docId, cookie, visibility, role, expiresInSeconds }) =>
      request({
        path: `/docs/${docId}/share-links`,
        method: "POST",
        cookie,
        body: { visibility, role, expiresInSeconds }
      })
  };
}
