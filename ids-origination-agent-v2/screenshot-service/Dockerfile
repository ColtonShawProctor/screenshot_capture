FROM ubuntu:22.04

# Install LibreOffice with UNO, Python, and fonts
RUN apt-get update && apt-get install -y \
    libreoffice-calc \
    libreoffice-writer \
    python3-uno \
    python3-pip \
    fonts-liberation \
    fonts-dejavu \
    curl \
    net-tools \
    iproute2 \
    poppler-utils \
    && rm -rf /var/lib/apt/lists/*

# Install Node.js for the HTTP service wrapper
RUN curl -fsSL https://deb.nodesource.com/setup_18.x | bash - \
    && apt-get install -y nodejs

WORKDIR /app
COPY package*.json ./
RUN npm install
COPY . .

# Expose port
EXPOSE 3000

# Start LibreOffice listener in background, then Node server
CMD soffice --headless --accept="socket,host=0.0.0.0,port=2002;urp;" & \
    sleep 5 && \
    node server.js