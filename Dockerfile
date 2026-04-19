FROM node:20-alpine AS builder

WORKDIR /app
COPY package*.json ./
RUN npm ci

COPY tsconfig.json ./
COPY src ./src
RUN npm run build     # esbuild → dist/index.js (57 kB, <1 s)

# ── runtime stage ────────────────────────────────────────────────
FROM node:20-alpine AS runtime

WORKDIR /app

# Create directory for token cache (mount a volume here)
RUN mkdir -p /data && chown node:node /data

COPY package*.json ./
RUN npm ci --omit=dev

COPY --from=builder /app/dist ./dist

# Token cache lives at /data/tokens.json (override via TOKEN_CACHE_PATH)
VOLUME ["/data"]

USER node

ENV NODE_ENV=production
ENV TOKEN_CACHE_PATH=/data/tokens.json

ENTRYPOINT ["node", "dist/index.js"]
