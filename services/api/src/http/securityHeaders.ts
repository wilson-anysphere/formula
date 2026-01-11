import type { FastifyInstance } from "fastify";

export function registerSecurityHeaders(app: FastifyInstance): void {
  const enableHsts = app.config.cookieSecure || process.env.NODE_ENV === "production";

  app.addHook("onSend", (request, reply, payload, done) => {
    // Only set headers when missing so individual endpoints can deliberately override them.
    if (!reply.hasHeader("x-content-type-options")) reply.header("x-content-type-options", "nosniff");
    if (!reply.hasHeader("x-frame-options")) reply.header("x-frame-options", "DENY");
    if (!reply.hasHeader("referrer-policy")) reply.header("referrer-policy", "no-referrer");
    if (!reply.hasHeader("permissions-policy")) {
      reply.header(
        "permissions-policy",
        [
          "accelerometer=()",
          "autoplay=()",
          "camera=()",
          "geolocation=()",
          "gyroscope=()",
          "magnetometer=()",
          "microphone=()",
          "payment=()",
          "usb=()"
        ].join(", ")
      );
    }

    if (enableHsts && !reply.hasHeader("strict-transport-security")) {
      // Only enable HSTS when we believe HTTPS is in use (cookieSecure=true), otherwise local dev
      // HTTP requests can get stuck due to browser caching.
      reply.header("strict-transport-security", "max-age=31536000; includeSubDomains");
    }

    done(null, payload);
  });
}

