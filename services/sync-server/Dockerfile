FROM node:20-slim AS build

# Build from repo root:
#   docker build -f services/sync-server/Dockerfile .

WORKDIR /app

# Preserve monorepo-ish layout so relative imports from dist/ keep working.
COPY services/sync-server/package.json ./services/sync-server/package.json

WORKDIR /app/services/sync-server

RUN npm install

WORKDIR /app

COPY services/sync-server/tsconfig.json ./services/sync-server/tsconfig.json
COPY services/sync-server/src ./services/sync-server/src

# Shared crypto modules (encryption-at-rest).
# Note: we also copy packages/security/package.json so Node treats crypto/*.js as ESM.
COPY packages/security/package.json ./packages/security/package.json
COPY packages/security/crypto ./packages/security/crypto

WORKDIR /app/services/sync-server

RUN npm run build

FROM node:20-slim

WORKDIR /app/services/sync-server

ENV NODE_ENV=production
ENV SYNC_SERVER_HOST=0.0.0.0
ENV SYNC_SERVER_PORT=1234

COPY --from=build /app/services/sync-server/package.json ./package.json
COPY --from=build /app/services/sync-server/node_modules ./node_modules
COPY --from=build /app/services/sync-server/dist ./dist

COPY --from=build /app/packages/security/package.json /app/packages/security/package.json
COPY --from=build /app/packages/security/crypto /app/packages/security/crypto

EXPOSE 1234
CMD ["node", "dist/index.js"]
