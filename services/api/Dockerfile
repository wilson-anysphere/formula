FROM node:20-alpine AS deps

WORKDIR /app/services/api

COPY services/api/package.json services/api/package-lock.json ./

RUN npm ci

FROM node:20-alpine AS build

WORKDIR /app/services/api

COPY --from=deps /app/services/api/node_modules ./node_modules

COPY services/api/package.json ./package.json
COPY services/api/tsconfig.json ./tsconfig.json
COPY services/api/src ./src
COPY services/api/migrations ./migrations

# Runtime dependency for DLP evaluation (used via relative imports).
COPY shared/dlp-core /app/shared/dlp-core

RUN npm run build

FROM node:20-alpine

WORKDIR /app/services/api

ENV NODE_ENV=production

COPY --from=build /app/services/api/node_modules ./node_modules
COPY --from=build /app/services/api/dist ./dist
COPY --from=build /app/services/api/migrations ./migrations
COPY --from=build /app/services/api/package.json ./package.json

COPY --from=build /app/shared /app/shared

CMD ["node", "dist/index.js"]
