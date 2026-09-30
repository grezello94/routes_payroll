FROM node:22-bookworm-slim

WORKDIR /app

ENV NODE_ENV=production
ENV HOST=0.0.0.0
ENV PORT=5501

RUN apt-get update \
  && apt-get install -y --no-install-recommends python3 make g++ ca-certificates \
  && rm -rf /var/lib/apt/lists/*

COPY package*.json ./
RUN npm install --omit=dev --build-from-source=sqlite3

COPY . .

RUN mkdir -p /app/data/monthly-backups \
  && chown -R node:node /app

USER node

EXPOSE 5501

HEALTHCHECK --interval=30s --timeout=5s --start-period=20s --retries=3 \
  CMD node -e "fetch('http://127.0.0.1:' + (process.env.PORT || 5501) + '/api/health').then(r => process.exit(r.ok ? 0 : 1)).catch(() => process.exit(1))"

CMD ["npm", "start"]
