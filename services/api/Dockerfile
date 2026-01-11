FROM node:20-alpine AS deps

WORKDIR /app/services/api

COPY services/api/package.json services/api/package-lock.json ./

# The API depends on some monorepo packages via `workspace:*` (pnpm protocol),
# which npm does not understand. For the Docker build, strip `workspace:*`
# dependencies and vendor those packages into `node_modules` manually.
RUN node -e 'const fs=require("fs");const pkg=JSON.parse(fs.readFileSync("package.json","utf8"));const vendored=new Set(["@formula/audit-core"]);for(const section of ["dependencies","devDependencies","optionalDependencies"]){const deps=pkg[section];if(!deps) continue;for(const [name,ver] of Object.entries(deps)){if(typeof ver==="string"&&ver.startsWith("workspace:")){if(!vendored.has(name)){console.error(`Unsupported workspace dependency (${section}): ${name}=${ver}. Update services/api/Dockerfile to vendor it.`);process.exit(1);}delete deps[name];}}}fs.writeFileSync("package.json",JSON.stringify(pkg,null,2));'

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

# Shared crypto modules (encryption-at-rest).
# Note: we also copy packages/security/package.json so Node treats crypto/*.js as ESM.
COPY packages/security/package.json /app/packages/security/package.json
COPY packages/security/crypto /app/packages/security/crypto

RUN npm run build

FROM node:20-alpine

WORKDIR /app/services/api

ENV NODE_ENV=production

COPY --from=build /app/services/api/node_modules ./node_modules
COPY --from=build /app/services/api/dist ./dist
COPY --from=build /app/services/api/migrations ./migrations
COPY --from=build /app/services/api/package.json ./package.json

COPY --from=build /app/shared /app/shared

COPY --from=build /app/packages/security/package.json /app/packages/security/package.json
COPY --from=build /app/packages/security/crypto /app/packages/security/crypto

CMD ["node", "dist/index.js"]
