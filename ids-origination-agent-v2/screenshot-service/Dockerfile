FROM node:20-slim

# Install Chromium and fonts for proper rendering
RUN apt-get update && apt-get install -y \
    chromium \
    fonts-liberation \
    fonts-noto-cjk \
    --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

# Install curl for health checks
RUN apt-get update && apt-get install -y curl --no-install-recommends && rm -rf /var/lib/apt/lists/*

# Set Puppeteer to use installed Chromium
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --omit=dev

# Copy application files
COPY server.js tableDetector.js ./

# Expose port
EXPOSE 3000

# Run server
CMD ["node", "server.js"]