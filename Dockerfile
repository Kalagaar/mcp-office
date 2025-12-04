# syntax=docker/dockerfile:1

ARG PYTHON_VERSION=3.12.8
FROM python:${PYTHON_VERSION}-slim AS base

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

# System dependencies for python-docx/pptx (fonts) and PDF conversion
RUN apt-get update \ 
    && apt-get install -y --no-install-recommends \
        build-essential \
        libreoffice \
        poppler-utils \
        ghostscript \
        fonts-dejavu-core \
        fonts-liberation \
    && rm -rf /var/lib/apt/lists/*

# Install deps first to leverage Docker cache
COPY requirements.txt ./
RUN pip install --no-cache-dir -r requirements.txt

# Copy source
COPY . .

# Ensure runtime dirs exist
RUN mkdir -p output custom_templates config

# Expose MCP HTTP port
EXPOSE 8958

CMD ["python", "-m", "app.main"]
