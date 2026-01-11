FROM node:20-slim AS build

# Build from repo root:
#   docker build -f services/sync-server/Dockerfile .

WORKDIR /app

# Native deps (node-gyp) for optional LevelDB persistence:
# y-leveldb -> level -> leveldown (native).
RUN apt-get update \
  && apt-get install -y --no-install-recommends \
    ca-certificates \
    g++ \
    make \
    python3 \
  && rm -rf /var/lib/apt/lists/*

# Preserve monorepo-ish layout so relative imports from dist/ keep working.
COPY services/sync-server/package.json ./services/sync-server/package.json
COPY services/sync-server/package-lock.json ./services/sync-server/package-lock.json

WORKDIR /app/services/sync-server

# Optional: force native modules to compile from source (e.g. to validate node-gyp toolchain in CI):
#   docker build --build-arg npm_config_build_from_source=true -f services/sync-server/Dockerfile .
ARG npm_config_build_from_source

RUN npm ci
# y-leveldb is an optionalDependency (to keep local dev portable), but the Docker
# image must include it so the default persistence backend (leveldb) works in
# production.
RUN node -e 'require("y-leveldb")'

WORKDIR /app

COPY services/sync-server/tsconfig.json ./services/sync-server/tsconfig.json
COPY services/sync-server/src ./services/sync-server/src

# Shared crypto modules (encryption-at-rest).
# Note: we also copy packages/security/package.json so Node treats crypto/*.js as ESM.
COPY packages/security/package.json ./packages/security/package.json
COPY packages/security/crypto ./packages/security/crypto

WORKDIR /app/services/sync-server

# This image only copies the sync-server sources plus packages/security/crypto.
# Avoid importing other monorepo packages unless you also copy them here.
RUN node -e 'const fs=require("fs");const path=require("path");const bad=[];const deny=[/@formula\//,/packages\/(?!security\/crypto(?:\/|$))/];(function walk(d){for(const e of fs.readdirSync(d,{withFileTypes:true})){const p=path.join(d,e.name);if(e.isDirectory())walk(p);else if(e.isFile()&&p.endsWith(".ts")){const t=fs.readFileSync(p,"utf8");if(deny.some(r=>r.test(t)))bad.push(p);}}})("src");if(bad.length){console.error("sync-server Docker build only includes services/sync-server and packages/security/crypto. Offending files:",bad);process.exit(1);}'

RUN npm run build
RUN npm prune --omit=dev

FROM node:20-slim

WORKDIR /app/services/sync-server

# Runtime should include CA certificates so outbound TLS (e.g. future KMS providers)
# works reliably even in slim images.
RUN apt-get update \
  && apt-get install -y --no-install-recommends \
    ca-certificates \
  && rm -rf /var/lib/apt/lists/*

ENV NODE_ENV=production
ENV SYNC_SERVER_HOST=0.0.0.0
ENV SYNC_SERVER_PORT=1234

COPY --from=build --chown=node:node /app/services/sync-server/package.json ./package.json
COPY --from=build --chown=node:node /app/services/sync-server/node_modules ./node_modules
COPY --from=build --chown=node:node /app/services/sync-server/dist ./dist

COPY --from=build /app/packages/security/package.json /app/packages/security/package.json
COPY --from=build /app/packages/security/crypto /app/packages/security/crypto

# Default `SYNC_SERVER_DATA_DIR` resolves under this workdir.
RUN mkdir -p /app/services/sync-server/.sync-server-data \
  && chown -R node:node /app/services/sync-server/.sync-server-data

# Support `SYNC_SERVER_DATA_DIR=/data` (used by docker-compose). When a fresh
# named volume is mounted at `/data`, Docker copies the directory contents from
# the image into the volume. Keep a placeholder file so ownership is preserved
# (otherwise Docker creates the empty directory as root-owned, which prevents
# the non-root `node` user from writing persistence files).
RUN mkdir -p /data \
  && touch /data/.keep \
  && chown -R node:node /data

USER node

EXPOSE 1234
CMD ["node", "dist/index.js"]
