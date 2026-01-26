FROM node:20-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci --ignore-scripts

COPY . .

# Only run generate if client.ts doesn't exist
# This allows pre-generated files to be used in Docker builds
RUN if [ ! -f "src/generated/client.ts" ]; then \
      echo "Generated client not found, running generation..."; \
      npm run generate || (echo "⚠️ Generation failed - ensure src/generated/client.ts exists" && exit 1); \
    else \
      echo "✅ Using pre-generated client file"; \
    fi

RUN npm run build

FROM node:20-alpine AS release

WORKDIR /app

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/package*.json ./

ENV NODE_ENV=production
RUN npm ci --ignore-scripts --omit=dev

ENTRYPOINT ["node", "dist/index.js"]

