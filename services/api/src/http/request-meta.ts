import type { FastifyRequest } from "fastify";

export function getClientIp(request: FastifyRequest): string | null {
  const ip = request.ip;
  return typeof ip === "string" ? ip : null;
}

export function getUserAgent(request: FastifyRequest): string | null {
  const ua = request.headers["user-agent"];
  return typeof ua === "string" ? ua : null;
}

