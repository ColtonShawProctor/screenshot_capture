FROM node:20-slim

# Install LibreOffice, Python, and other tools for Excel range extraction
RUN apt-get update && apt-get install -y \
    libreoffice-calc \
    python3 \
    python3-pip \
    poppler-utils \
    fonts-liberation \
    fonts-noto-cjk \
    curl \
    --no-install-recommends && \
    rm -rf /var/lib/apt/lists/*

# Install Python dependencies
RUN pip3 install openpyxl

# Remove old ImageMagick and Chromium configs (no longer needed)
# The new Python-based approach uses LibreOffice + poppler only

WORKDIR /app

# Copy package files
COPY package*.json ./

# Install dependencies
RUN npm install --omit=dev

# Copy application files
COPY server.js ./
COPY tableDetector.js ./
COPY excelVisualRenderer.js ./
COPY extract_table.py ./

# Make Python script executable  
RUN chmod +x extract_table.py

# Expose port
EXPOSE 3000

# Run server
CMD ["node", "server.js"]