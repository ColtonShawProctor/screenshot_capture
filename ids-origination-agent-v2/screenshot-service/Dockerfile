FROM node:20-slim

# Install LibreOffice, Chromium and fonts for Excel visual rendering
RUN apt-get update && apt-get install -y \
    libreoffice \
    chromium \
    fonts-liberation \
    fonts-noto-cjk \
    curl \
    imagemagick \
    poppler-utils \
    --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

# Configure ImageMagick to allow PDF processing
RUN sed -i 's/policy domain="coder" rights="none" pattern="PDF"/policy domain="coder" rights="read|write" pattern="PDF"/g' /etc/ImageMagick-6/policy.xml

# Set Puppeteer to use installed Chromium
ENV PUPPETEER_EXECUTABLE_PATH=/usr/bin/chromium
ENV PUPPETEER_SKIP_CHROMIUM_DOWNLOAD=true

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --omit=dev

# Copy application files
COPY server.js tableDetector.js excelVisualRenderer.js ./

# Expose port
EXPOSE 3000

# Run server
CMD ["node", "server.js"]