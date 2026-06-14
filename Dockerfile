# --- Stage 1: Build the React WebApp ---
FROM node:20-slim AS frontend-builder
WORKDIR /app/webapp
COPY webapp/package*.json ./
RUN npm install
COPY webapp/ ./
RUN npm run build

# --- Stage 2: Final Production Image ---
FROM python:3.11-slim

# Install system dependencies for LibreOffice and high-quality fonts.
# fonts-crosextra-carlito is metric-compatible with Calibri and fonts-crosextra-caladea
# with Cambria; fontconfig aliases them automatically so the optional PDF export
# (rendered server-side by LibreOffice) matches what clients see in the .pptx.
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-liberation \
    fonts-crosextra-carlito \
    fonts-crosextra-caladea \
    fonts-dejavu \
    fonts-freefont-ttf \
    fonts-noto \
    fonts-noto-cjk \
    fonts-noto-color-emoji \
    fonts-font-awesome \
    fontconfig \
    libcap2-bin \
    --no-install-recommends \
    && fc-cache -fv \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Install Python dependencies
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copy backend application code
COPY . .

# Copy built frontend from Stage 1
COPY --from=frontend-builder /app/webapp/dist ./webapp/dist

# Ensure the app can run without root if needed
ENV PYTHONUNBUFFERED=1

# Expose the port (Render uses PORT env var)
EXPOSE 10000

CMD ["python", "-m", "app.main"]
