FROM node:20-slim AS build

WORKDIR /app

COPY services/sync-server/package.json ./package.json
COPY services/sync-server/tsconfig.json ./tsconfig.json
COPY services/sync-server/src ./src

RUN npm install
RUN npm run build

FROM node:20-slim

WORKDIR /app

ENV NODE_ENV=production
ENV SYNC_SERVER_HOST=0.0.0.0
ENV SYNC_SERVER_PORT=1234

COPY --from=build /app/package.json ./package.json
COPY --from=build /app/node_modules ./node_modules
COPY --from=build /app/dist ./dist

EXPOSE 1234
CMD ["node", "dist/index.js"]

