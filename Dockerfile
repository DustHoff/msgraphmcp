FROM node:20-alpine

WORKDIR /app

# Create directory for token cache (mount a volume here)
RUN mkdir -p /data && chown node:node /data

COPY package*.json ./
RUN npm ci --omit=dev --no-fund --silent

# dist/ is pre-built by CI (esbuild runs natively on the CI runner, not under
# QEMU emulation) and copied into the build context before docker build runs.
COPY dist ./dist

# Token cache lives at /data/tokens.json (override TOKEN_CACHE_PATH at runtime)
VOLUME ["/data"]

USER node

ENV NODE_ENV=production

# When PORT is set the server listens on HTTP (Kubernetes mode).
# When PORT is unset the server uses stdio (Claude Code / local mode).
EXPOSE 8080

# K8s liveness / readiness probe (only active in HTTP mode)
HEALTHCHECK --interval=30s --timeout=5s --start-period=60s --retries=3 \
  CMD node -e "\
    const port = process.env.PORT || 8080;\
    require('http').get('http://localhost:' + port + '/health', r => {\
      process.exit(r.statusCode === 200 ? 0 : 1);\
    }).on('error', () => process.exit(1));"

ENTRYPOINT ["node", "dist/index.js"]
