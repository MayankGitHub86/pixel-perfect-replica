# Build stage
FROM node:20-alpine AS builder
WORKDIR /app

# Copy dependency definitions and prisma schema
COPY package*.json ./
COPY prisma ./prisma/

# Install dependencies including dev dependencies
RUN npm ci

# Generate Prisma client
RUN npx prisma generate

# Copy source code and build
COPY tsconfig.json ./
COPY src ./src
RUN npm run build

# Production stage
FROM node:20-alpine AS runner
WORKDIR /app
ENV NODE_ENV=production

COPY package*.json ./
COPY prisma ./prisma/

# Install only production dependencies
RUN npm ci --omit=dev

# Generate Prisma client for production runner
RUN npx prisma generate

# Copy compiled files from builder
COPY --from=builder /app/dist ./dist

EXPOSE 3000

CMD ["node", "dist/server.js"]
