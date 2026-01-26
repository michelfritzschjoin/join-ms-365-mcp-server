FROM node:24-alpine AS builder

WORKDIR /app

COPY package*.json ./
RUN npm ci --ignore-scripts

COPY . .

# Generate client code or use pre-generated file
# Note: Generation fails in qemu-emulated ARM64 environments due to native dependencies
# Solution: Commit src/generated/client.ts to the repository before building
RUN ARCH=$(uname -m) && \
    if [ "$ARCH" = "aarch64" ] && [ -f /.dockerenv ]; then \
      echo "⚠️ ARM64 emulated build detected - skipping generation (not supported in qemu)"; \
      if [ ! -f "src/generated/client.ts" ]; then \
        echo "❌ ERROR: src/generated/client.ts not found!"; \
        echo "   Generation is not supported in qemu-emulated ARM64 environments"; \
        echo "   Please ensure src/generated/client.ts is committed to the repository"; \
        echo "   Or build on a native ARM64 platform (not emulated)"; \
        exit 1; \
      else \
        echo "✅ Using pre-generated client file"; \
      fi; \
    elif [ ! -f "src/generated/client.ts" ]; then \
      echo "Generated client not found, attempting generation..."; \
      npm run generate || { \
        echo "❌ Generation failed"; \
        echo "   If building for ARM64, ensure src/generated/client.ts is committed"; \
        exit 1; \
      }; \
      echo "✅ Generation successful"; \
    else \
      echo "✅ Using pre-generated client file"; \
    fi

RUN npm run build

FROM node:24-alpine AS release

WORKDIR /app

COPY --from=builder /app/dist /app/dist
COPY --from=builder /app/package*.json ./

ENV NODE_ENV=production
RUN npm ci --ignore-scripts --omit=dev

ENTRYPOINT ["node", "dist/index.js"]

