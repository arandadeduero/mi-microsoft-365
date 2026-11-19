# Stage 1: Build & Dependencies
FROM node:25-alpine AS builder

WORKDIR /usr/src/app

# Copy package files
COPY package*.json ./

# Install only production dependencies and drop non-runtime TypeScript types / source maps
RUN npm ci --omit=dev --no-audit --no-fund \
  && find node_modules -type d -name dist-types -prune -exec rm -rf {} + \
  && find node_modules -type f \( -name '*.d.ts' -o -name '*.d.ts.map' -o -name '*.js.map' \) -delete

# Stage 2: Production Runner
FROM node:25-alpine AS runner

WORKDIR /usr/src/app

ENV NODE_ENV=production
ENV PORT=3000

# Copy node_modules from builder
COPY --from=builder --chown=node:node /usr/src/app/node_modules ./node_modules

# Copy app source
COPY --chown=node:node . .

# The directories the app writes to, owned by the user it runs as. A named
# volume mounted on one of them takes this ownership the first time
RUN mkdir -p sessions audit data && chown node:node sessions audit data

# Run as the unprivileged user the base image provides
USER node

# Expose port
EXPOSE 3000

# Start command (running node directly handles system signals correctly)
CMD ["node", "server.js"]
