FROM python:3.11-slim

# Install system dependencies for LibreOffice and high-quality fonts
RUN apt-get update && apt-get install -y \
    libreoffice \
    fonts-liberation \
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

# Copy application code
COPY . .

# Ensure the app can run without root if needed, but standard is fine
# Set environment variables
ENV PYTHONUNBUFFERED=1

CMD ["python", "-m", "app.main"]
