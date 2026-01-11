FROM node:20-alpine AS deps

WORKDIR /app/services/api

COPY services/api/package.json services/api/package-lock.json ./

# The API depends on some monorepo packages via `workspace:*` (pnpm protocol),
# which npm does not understand. For the Docker build, strip `workspace:*`
# dependencies and vendor those packages into `node_modules` manually.
RUN node -e 'const fs=require("fs");const pkg=JSON.parse(fs.readFileSync("package.json","utf8"));if(pkg.dependencies){for(const [name,ver] of Object.entries(pkg.dependencies)){if(typeof ver==="string"&&ver.startsWith("workspace:")) delete pkg.dependencies[name];}}fs.writeFileSync("package.json",JSON.stringify(pkg,null,2));'

RUN npm ci

# Vendored workspace deps (must match what we stripped above).
RUN mkdir -p node_modules/@formula
COPY packages/audit-core ./node_modules/@formula/audit-core

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
